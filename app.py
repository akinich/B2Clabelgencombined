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
    font_size = 1
    while True:
        total_height = 0
        max_line_width = 0
        for lines in column_lines:
            total_height += len(lines) * font_size + (len(lines) - 1) * LINE_SPACING
            max_line_width = max(max_line_width, max(pdfmetrics.stringWidth(line, font_name, font_size) for line in lines))
        if max_line_width > (max_width - 4) or total_height > (max_height - 4):
            return max(font_size - 1, 1)
        font_size += 1

def draw_label_pdf(c, order_no_text, customer_name_text, font_name, width, height, font_override=0):
    """
    Draws a label with order number on top and customer name below.
    """
    # Ensure order: order number first, customer name second
    column_lines = []
    for text in [order_no_text, customer_name_text]:  # Order enforced here
        words = str(text).strip().split()
        column_lines.append(words if words else [""])

    # Determine max font size
    raw_font_size = find_max_font_size_for_multiline(column_lines, width, height, font_name)
    font_size = max(raw_font_size - FONT_ADJUSTMENT + font_override, 1)
    c.setFont(font_name, font_size)

    # Compute starting y to vertically center content
    total_height = sum(len(lines) * font_size + (len(lines) - 1) * LINE_SPACING for lines in column_lines)
    start_y = (height - total_height) / 2

    y = start_y
    for lines in column_lines:
        for line in lines:
            line_width = pdfmetrics.stringWidth(line, font_name, font_size)
            x = (width - line_width) / 2
            c.drawString(x, y, line)
            y += font_size + LINE_SPACING

def create_pdf(data_list, font_name, width, height, font_override=0):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=(width, height))
    for order_no, customer_name in data_list:
        draw_label_pdf(c, order_no, customer_name, font_name, width, height, font_override)
        c.showPage()
    c.save()
    buffer.seek(0)
    return buffer

# === STREAMLIT UI ===
st.title("Order Label PDF Generator")
st.write("Generates PDF labels with Order No on top and Customer Name below.")

# --- User Inputs ---
selected_font = st.selectbox("Select font", AVAILABLE_FONTS, index=1)
font_override = st.slider("Font size override (+/- points)", min_value=-5, max_value=5, value=0)

width_mm = st.number_input("Label width (mm)", min_value=10, max_value=500, value=DEFAULT_WIDTH_MM)
height_mm = st.number_input("Label height (mm)", min_value=10, max_value=500, value=DEFAULT_HEIGHT_MM)

remove_duplicates = st.checkbox("Remove duplicate values", value=True)

# --- File Uploader ---
uploaded_file = st.file_uploader("Upload CSV or Excel file", type=["csv", "xlsx"])

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

# --- Process Data ---
if df is not None:
    st.write("Preview of data:")
    st.dataframe(df)

    # Identify relevant columns (case-insensitive)
    df_columns_lower = [col.lower() for col in df.columns]
    try:
        order_no_col = df.columns[df_columns_lower.index("order no")]
        customer_name_col = df.columns[df_columns_lower.index("customer name")]
    except ValueError:
        st.error("Uploaded file must have 'order no' and 'customer name' columns (case-insensitive).")
        st.stop()

    # Extract relevant data
    data_list = df[[order_no_col, customer_name_col]].fillna("").astype(str).values.tolist()

    # Remove duplicates if needed
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
                file_name="order_labels.pdf",
                mime="application/pdf"
            )
