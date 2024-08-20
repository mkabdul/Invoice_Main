

import pandas as pd
import aspose.words as aw
import io
from io import BytesIO
import streamlit as st
from langchain_core.messages import HumanMessage
from langchain_google_genai import ChatGoogleGenerativeAI
from PIL import Image, ImageDraw, ImageFont
import os, fitz
import json, time
from docx import Document


os.environ["GOOGLE_API_KEY"] = 'AIzaSyBbepUh8x3CqpkxNFnJ1IX0dFc0UNTwwbU'


# Initialize the ChatGoogleGenerativeAI model
llm = ChatGoogleGenerativeAI(model="gemini-1.5-pro-latest")

st.markdown("""<style>
        .stButton > button {
            display: block;
            margin: 0 auto;}</style>
        """, unsafe_allow_html=True)

# Custom CSS to display radio options
st.markdown("""
    <style>
    div[role="radiogroup"] > label > div {
        display: flex;
        flex-direction: row;
    }
    div[role="radiogroup"] > label > div > div {
        margin-right: 10px;
    }
    </style>
    <style>
    div[role="radiogroup"] {
        display: flex;
        flex-direction: row;
    }
    div[role="radiogroup"] > label {
        margin-right: 20px;
    }
    input[type="radio"]:div {
        background-color: white;
        border-color: lightblue;
    }
    </style>
""", unsafe_allow_html=True)

def process_invoice(image_path):
    message = HumanMessage(
        content=[
            {
                "type": "text",
                "text": """You will carefully analyze the invoice only and get the output in pure JSON within a Python list format with relevant details
                with invoice id as default ino [{}], your response shall not contain ' ```python ' and ' ``` '
                        """,
            },
            {"type": "image_url", "image_url": image_path}
        ]
    )
    response = llm.invoke([message])
    return response.content

def convert_pdf_to_images_with_pymupdf(pdf_path, output_folder, zoom_x=2.0, zoom_y=2.0):
    """Convert PDF pages to images with high resolution."""
    doc = fitz.open(pdf_path)
    image_paths = []
    for page_number in range(len(doc)):
        page = doc.load_page(page_number)
        mat = fitz.Matrix(zoom_x, zoom_y)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        unique_name = f'output_image_{int(time.time())}_{page_number + 1}.png'
        image_path = os.path.join(output_folder, unique_name)
        pix.save(image_path)
        image_paths.append(image_path)
        print(f'Generated image: {image_path}')
    return image_paths

def convert_docx_to_images(docx_file, invoice_dir):
    """Convert DOCX file pages to high-quality images."""
    doc = aw.Document(docx_file)
    images = []

    for i in range(doc.page_count):
        options = aw.saving.ImageSaveOptions(aw.SaveFormat.JPEG)
        options.page_set = aw.saving.PageSet(i)
        options.horizontal_resolution = 300
        options.vertical_resolution = 300

        # Save the page image to a BytesIO buffer
        buffer = io.BytesIO()
        doc.save(buffer, options)
        buffer.seek(0)

        # Open the image from the buffer
        image = Image.open(buffer)
        images.append(image)

    # Get the width and height of the combined image
    total_width = max(image.width for image in images)
    total_height = sum(image.height for image in images)

    # Create a new blank image with the combined size
    combined_image = Image.new('RGB', (total_width, total_height))

    # Paste each page image into the combined image
    y_offset = 0
    for image in images:
        combined_image.paste(image, (0, y_offset))
        y_offset += image.height

    # Save the combined image with the same name as the DOCX file
    docx_filename = os.path.splitext(os.path.basename(docx_file))[0]
    combined_image_file = os.path.join(invoice_dir, f"{docx_filename}.jpg")
    combined_image.save(combined_image_file, quality=100)

    return combined_image_file

def txt_to_image(txt_file, invoice_dir, custom_font_path=None):
    """Convert text file to a high-resolution image."""
    with open(txt_file, 'r') as f:
        text = f.read()

    # Create a new image with higher resolution
    image = Image.new('RGB', (1600, 1200), color=(255, 255, 255))
    draw = ImageDraw.Draw(image)

    # Use the custom font if provided, else fallback to default
    if custom_font_path:
        try:
            font = ImageFont.truetype(custom_font_path, 24)
        except IOError:
            font = ImageFont.load_default()
    else:
        font = ImageFont.load_default()

    # Draw the text on the image
    draw.text((10, 10), text, font=font, fill=(0, 0, 0))

    # Save the image with high quality
    txt_image_path = os.path.join(invoice_dir, f"{os.path.splitext(os.path.basename(txt_file))[0]}.png")
    image.save(txt_image_path, quality=100)

    return txt_image_path

def clear_invoice_dir(invoice_dir):
    """Clear all files in the specified directory."""
    for filename in os.listdir(invoice_dir):
        file_path = os.path.join(invoice_dir, filename)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)  # Remove file
            elif os.path.isdir(file_path):
                os.rmdir(file_path)  # Remove directory if it's empty
        except Exception as e:
            print(f'Failed to delete {file_path}. Reason: {e}')

def parse_invoice_to_table(output):
    # Log or print the raw output for debugging
    # st.write(f"Raw Output: {output}")  # Removed raw output display

    # Ensure output is a dictionary or list
    try:
        if isinstance(output, str):
            json_output = json.loads(output)
        else:
            json_output = output
    except json.JSONDecodeError as e:
        st.error(f"Failed to decode JSON from the output: {e}")
        return pd.DataFrame()

    # Handle cases where json_output might be a list
    if isinstance(json_output, list):
        if len(json_output) > 0 and isinstance(json_output[0], dict):
            json_output = json_output[0]  # Take the first item if it's a dictionary
        else:
            st.error("JSON list does not contain dictionaries.")
            return pd.DataFrame()

    if not isinstance(json_output, dict):
        st.error("Output is neither a dictionary nor a valid list of dictionaries.")
        return pd.DataFrame()

    # Define the expected fields for the table
    fields = [
        "Invoice #", "Invoice Date",  "amount_due", "bill_to",
        "shipping_to", "name", "phone", "payment_terms", "items",
        "description", "quantity", "rate", "amount", "product_based",
        "subtotal", "tax", "grand_total"
    ]

    # Initialize an empty dictionary for the table data
    table_data = {field: json_output.get(field, None) for field in fields}

    # Add items to the table if they exist
    if "items" in json_output:
        items = json_output["items"]
        for item in items:
            for field in ["description", "quantity", "rate", "amount"]:
                table_data.setdefault(f"Item_{item.get('item_number', '')}_{field}", item.get(field, None))

    return pd.DataFrame([table_data])

def main():
    # Load and display the company logo
    logo_path = "logo.png"  # Replace with your logo path
    logo = Image.open(logo_path)

    # Layout for logo and title
    col1, col2 = st.columns([1, 5])  # Adjust column widths as needed

    with col1:
        st.image(logo, width=100)  # Adjust the width as needed

    with col2:
        st.title("Invoice Data Analyzer")

    option = st.radio(
        "Select an option:",
        ("Upload Invoice Images, PDFs, TXT Files", "Select Existing Images")
    )

    use_custom_font = st.checkbox('Use custom font for txt to image')

    custom_font_path = None

    if use_custom_font:
        font_dir = '/tmp/fonts/'
        if not os.path.exists(font_dir):
            os.makedirs(font_dir)

        existing_fonts = [f for f in os.listdir(font_dir) if f.endswith('.ttf')]

        if existing_fonts:
            font_choice = st.radio("Choose a font", existing_fonts)
            custom_font_path = os.path.join(font_dir, font_choice)
            st.write(f"Now using font: {font_choice}")

        uploaded_font = st.file_uploader("Upload .ttf for custom font", type=["ttf"])
        if uploaded_font:
            custom_font_path = os.path.join(font_dir, uploaded_font.name)
            with open(custom_font_path, "wb") as f:
                f.write(uploaded_font.getbuffer())
            st.write(f"Now using font: {uploaded_font.name}")

    if 'table_outputs' not in st.session_state:
        st.session_state.table_outputs = []

    invoice_dir = "invoice_images/"
    if not os.path.exists(invoice_dir):
        os.makedirs(invoice_dir)

    if option == "Upload Invoice Images, PDFs, TXT Files":
        uploaded_files = st.file_uploader("Choose PDF, Image, or TXT files", type=["pdf", "png", "jpg", "jpeg", "txt"], accept_multiple_files=True)

        for uploaded_file in uploaded_files:
            if uploaded_file.type == "application/pdf":
                pdf_path = os.path.join(invoice_dir, uploaded_file.name)
                with open(pdf_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                image_paths = convert_pdf_to_images_with_pymupdf(pdf_path, invoice_dir)
                for image_path in image_paths:
                    st.image(image_path)
                    json_output = process_invoice(image_path)
                    table_df = parse_invoice_to_table(json_output)
                    st.write("Extracted Data:", table_df)

            elif uploaded_file.type in ["image/png", "image/jpeg"]:
                image_path = os.path.join(invoice_dir, uploaded_file.name)
                with open(image_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                st.image(image_path)
                json_output = process_invoice(image_path)
                table_df = parse_invoice_to_table(json_output)
                st.write("Extracted Data:", table_df)

            elif uploaded_file.type == "text/plain":
                txt_file = os.path.join(invoice_dir, uploaded_file.name)
                with open(txt_file, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                image_path = txt_to_image(txt_file, invoice_dir, custom_font_path)
                st.image(image_path)
                json_output = process_invoice(image_path)
                table_df = parse_invoice_to_table(json_output)
                st.write("Extracted Data:", table_df)

    elif option == "Select Existing Images":
        existing_images = [f for f in os.listdir(invoice_dir) if f.endswith(('.png', '.jpg', '.jpeg'))]
        selected_images = st.multiselect("Select images", existing_images)

        for image_name in selected_images:
            image_path = os.path.join(invoice_dir, image_name)
            st.image(image_path)
            json_output = process_invoice(image_path)
            table_df = parse_invoice_to_table(json_output)
            st.write("Extracted Data:", table_df)

    if st.button('Clear Invoice Directory'):
        clear_invoice_dir(invoice_dir)
        st.write("Invoice directory cleared.")

if __name__ == "__main__":
    main()
