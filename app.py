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

def draw_label_pdf(c, lines, font_name, width, height, font_override=0):
    """Draws multiple lines on PDF, centered vertically and horizontally."""
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

def create_pdf(df, font_name, width, height, font_override=0):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=(width, height))

    for idx, row in df.iterrows():
        order_no = str(row["order no"]).strip()
        customer_name = str(row["customer name"]).strip()
        lines = [order_no, customer_name]
        draw_label_pdf(c, lines, font_name, width, height, font_override)
        c.showPage()

    c.save()
    buffer.seek(0)
    return buffer

# === STREAMLIT UI ===
st.title("Excel/CSV to Label PDF Generator (Order No + Customer Name)")
st.write("Generates PDF labels with Order No on top and Customer Name below.")

# --- User Inputs ---
selected_font = st.selectbox("Select font", AVAILABLE_FONTS, index=1)
font_override = st.slider("Font size override (+/- points)", min_value=-5, max_value=5, value=0)

width_mm = st.number_input("Label width (mm)", min_value=10, max_value=500, value=DEFAULT_WIDTH_MM)
height_mm = st.number_input("Label height (mm)", min_value=10, max_value=500, value=DEFAULT_HEIGHT_MM)

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

        # Normalize column names to lowercase
        df.columns = [col.strip().lower() for col in df.columns]

        # Ensure required columns exist
        required_cols = ["order no", "customer name"]
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            st.error(f"Missing required columns: {', '.join(missing_cols)}")
            df = None

    except Exception as e:
        st.error(f"Error reading file: {e}")

if df is not None:
    st.write("Preview of data:")
    st.dataframe(df[["order no", "customer name"]])

    # --- Generate PDF ---
    if st.button("Generate PDF"):
        if df.empty:
            st.warning("No valid data found!")
        else:
            with st.spinner("Generating PDF..."):
                pdf_buffer = create_pdf(df, selected_font, width_mm*mm, height_mm*mm, font_override)
            st.success("PDF generated!")
            st.download_button(
                label="Download PDF",
                data=pdf_buffer,
                file_name="labels.pdf",
                mime="application/pdf"
            )
