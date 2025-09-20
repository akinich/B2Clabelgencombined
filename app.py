import streamlit as st
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from io import BytesIO

# === DEFAULT CONSTANTS ===
DEFAULT_WIDTH_MM = 50
DEFAULT_HEIGHT_MM = 30
FONT_ADJUSTMENT = 2  # for printer safety
LINE_SPACING = 2     # space between words

# Built-in fonts
AVAILABLE_FONTS = [
    "Helvetica",
    "Helvetica-Bold",
    "Times-Roman",
    "Times-Bold",
    "Courier",
    "Courier-Bold"
]

# === HELPER FUNCTIONS ===
def find_max_font_size_for_multiline(column_lines, max_width, max_height, font_name):
    """
    column_lines: list of lists, each sublist contains stacked words for a column
    """
    font_size = 1
    while True:
        total_height = 0
        max_line_width = 0
        for lines in column_lines:
            total_height += len(lines) * font_size + (len(lines) - 1) * LINE_SPACING
            max_line_width = max(max_line_width, max(pdfmetrics.stringWidth(line, font_name, font_size) for line in lines))
        # Add horizontal lines between columns
        total_height += (len(column_lines) - 1) * 2  # 2 units for each line
        if max_line_width > (max_width - 4) or total_height > (max_height - 4):
            return max(font_size - 1, 1)
        font_size += 1

def draw_label_pdf(c, column_texts, font_name, width, height, font_override=0):
    """
    column_texts: list of strings, one per selected column
    """
    # Split each column text into words (stacked)
    column_lines = []
    for text in column_texts:
        words = str(text).strip().split()
        column_lines.append(words if words else [""])
    
    # Determine max font size
    raw_font_size = find_max_font_size_for_multiline(column_lines, width, height, font_name)
    font_size = max(raw_font_size - FONT_ADJUSTMENT + font_override, 1)
    c.setFont(font_name, font_size)

    # Compute starting y to vertically center all columns with lines
    total_height = sum(len(lines) * font_size + (len(lines) - 1) * LINE_SPACING for lines in column_lines)
    total_height += (len(column_lines) - 1) * 2  # space for horizontal lines
    start_y = (height - total_height) / 2

    y = start_y
    for i, lines in enumerate(column_lines):
        # Draw words in top-to-bottom order
        for line in lines:
            line_width = pdfmetrics.stringWidth(line, font_name, font_size)
            x = (width - line_width) / 2
            c.drawString(x, y, line)
            y += font_size + LINE_SPACING
        # Draw horizontal line between columns, except after last column
        if i < len(column_lines) - 1:
            c.line(2, y + 1, width - 2, y + 1)  # small horizontal line with margin
            y += 2  # space after line

def create_pdf(data_list, font_name, width, height, font_override=0):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=(width, height))
    for row in data_list:
        # row is a list of column texts for that row
        draw_label_pdf(c, row, font_name, width, height, font_override)
        c.showPage()
    c.save()
    buffer.seek(0)
    return buffer

# === STREAMLIT UI ===
st.title("Excel/CSV to Label PDF Generator")
st.write("Generate multi-page PDF labels with stacked words and horizontal lines between columns.")

# --- User Inputs ---
selected_font = st.selectbox("Select font", AVAILABLE_FONTS, index=1)
font_override = st.slider("Font size override (+/- points)", min_value=-5, max_value=5, value=0)

width_mm = st.number_input("Label width (mm)", min_value=10, max_value=500, value=DEFAULT_WIDTH_MM)
height_mm = st.number_input("Label height (mm)", min_value=10, max_value=500, value=DEFAULT_HEIGHT_MM)

remove_duplicates = st.checkbox("Remove duplicate values", value=True)

# --- File Uploader ---
uploaded_file = st.file_uploader("Upload CSV or Excel file", type=["csv", "xlsx"])

# --- Load Data ---
df = None
if uploaded_file:
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file, engine="openpyxl")
        st.success("File loaded successfully!")
    except Exception as e:
        st.error(f"Error reading file: {e}")

# --- Column Selection ---
if df is not None:
    st.write("Preview of data:")
    st.dataframe(df)
    
    selected_columns = st.multiselect(
        "Select columns to generate labels",
        options=df.columns.tolist(),
        default=df.columns.tolist()
    )
    
    if selected_columns:
        # Prepare data row-wise
        data_list = df[selected_columns].fillna("").astype(str).values.tolist()
        
        # Remove duplicates if checked
        if remove_duplicates:
            seen = set()
            unique_data = []
            for row in data_list:
                row_tuple = tuple(row)
                if row_tuple not in seen:
                    seen.add(row_tuple)
                    unique_data.append(row)
            data_list = unique_data

        # --- Generate PDF ---
        if st.button("Generate PDF"):
            if not data_list:
                st.warning("No valid data found!")
            else:
                pdf_buffer = create_pdf(data_list, selected_font, width_mm * mm, height_mm * mm, font_override)
                st.download_button(
                    label="Download PDF",
                    data=pdf_buffer,
                    file_name="labels.pdf",
                    mime="application/pdf"
                )
