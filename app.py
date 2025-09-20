import streamlit as st
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.pdfmetrics import stringWidth
from io import BytesIO

# === DEFAULT CONSTANTS ===
DEFAULT_WIDTH_MM = 50
DEFAULT_HEIGHT_MM = 30
FONT_ADJUSTMENT = 2  # for printer safety

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
def wrap_text_to_width(text, font_name, font_size, max_width):
    """Wrap a single line of text into multiple lines that fit within max_width."""
    words = text.split()
    if not words:
        return [""]

    lines = []
    current_line = words[0]

    for word in words[1:]:
        test_line = f"{current_line} {word}"
        if stringWidth(test_line, font_name, font_size) <= max_width:
            current_line = test_line
        else:
            lines.append(current_line)
            current_line = word
    lines.append(current_line)
    return lines

def find_max_font_size_for_multiline(lines, max_width, max_height, font_name):
    font_size = 1
    while True:
        wrapped_lines = []
        for line in lines:
            wrapped_lines.extend(wrap_text_to_width(line, font_name, font_size, max_width))
        total_height = len(wrapped_lines) * font_size + (len(wrapped_lines) - 1) * 2
        max_line_width = max(stringWidth(line, font_name, font_size) for line in wrapped_lines)
        if max_line_width > (max_width - 4) or total_height > (max_height - 4):
            return max(font_size - 1, 1)
        font_size += 1

def draw_label_pdf(c, text, font_name, width, height, font_override=0):
    # Split text by newlines only
    lines = text.split("\n")

    raw_font_size = find_max_font_size_for_multiline(lines, width, height, font_name)
    font_size = max(raw_font_size - FONT_ADJUSTMENT + font_override, 1)
    c.setFont(font_name, font_size)

    # Wrap lines to fit width
    wrapped_lines = []
    for line in lines:
        wrapped_lines.extend(wrap_text_to_width(line, font_name, font_size, width))

    total_height = len(wrapped_lines) * font_size + (len(wrapped_lines) - 1) * 2
    start_y = (height - total_height) / 2

    for i, line in enumerate(wrapped_lines):
        line_width = stringWidth(line, font_name, font_size)
        x = (width - line_width) / 2
        y = start_y + (len(wrapped_lines) - i - 1) * (font_size + 2)
        c.drawString(x, y, line)

def create_pdf(data_list, font_name, width, height, font_override=0):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=(width, height))
    for value in data_list:
        text = str(value).strip()
        if not text or text.lower() == "nan":
            continue
        draw_label_pdf(c, text, font_name, width, height, font_override)
        c.showPage()
    c.save()
    buffer.seek(0)
    return buffer

# === STREAMLIT UI ===
st.title("Excel/CSV to Label PDF Generator")
st.write("Generate multi-page PDF labels with wrapped text per cell.")

# --- User Inputs ---
selected_font = st.selectbox("Select font", AVAILABLE_FONTS, index=1)
font_override = st.slider("Font size override (+/- points)", min_value=-5, max_value=5, value=0)

width_mm = st.number_input("Label width (mm)", min_value=10, max_value=500, value=DEFAULT_WIDTH_MM)
height_mm = st.number_input("Label height (mm)", min_value=10, max_value=500, value=DEFAULT_HEIGHT_MM)

remove_duplicates = st.checkbox("Remove duplicate values", value=True)

# --- Separator Input ---
separator = st.text_input("Separator for combined columns (use \\n for stacked text)", value=" ")

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
        # Replace escaped newline string with actual newline
        actual_separator = separator.replace("\\n", "\n")
        
        # Combine selected columns row-wise
        combined_values = df[selected_columns].astype(str).agg(actual_separator.join, axis=1)
        
        # Remove empty/NaN
        combined_values = [val.strip() for val in combined_values if val.strip() != ""]
        if remove_duplicates:
            combined_values = list(dict.fromkeys(combined_values))

        # --- Generate PDF ---
        if st.button("Generate PDF"):
            if not combined_values:
                st.warning("No valid data found!")
            else:
                with st.spinner("Generating PDF..."):
                    pdf_buffer = create_pdf(combined_values, selected_font, width_mm*mm, height_mm*mm, font_override)
                st.success("PDF generated!")
                st.download_button(
                    label="Download PDF",
                    data=pdf_buffer,
                    file_name="labels.pdf",
                    mime="application/pdf"
                )
