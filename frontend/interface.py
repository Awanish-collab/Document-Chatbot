import streamlit as st
from pathlib import Path
import tempfile
import os
import pdfplumber
from PIL import Image
import io
import base64

def analyze_pdf(file_content):
    """Analyze PDF content including visual elements"""
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
        tmp_file.write(file_content)
        tmp_file_path = tmp_file.name

    try:
        with pdfplumber.open(tmp_file_path) as pdf:
            total_pages = len(pdf.pages)
            st.write("Total pages: ", total_pages)
            
            # Prepare HTML container with specific size and scroll
            html_content = """
            <div style="width: 100%; height: 400px; overflow-y: scroll; border: 1px solid black; padding: 10px;">
            """
            
            for i in range(total_pages):
                # Convert page to an image
                page_image = page_to_image(i, tmp_file_path)
                if page_image:
                    # Convert PIL image to base64 for HTML display
                    buffered = io.BytesIO()
                    page_image.save(buffered, format="PNG")
                    img_str = base64.b64encode(buffered.getvalue()).decode("utf-8")
                    html_content += f'<img src="data:image/png;base64,{img_str}" style="width:100%; margin-bottom: 20px;" />'
                else:
                    html_content += f'<p>Could not render page {i + 1}</p>'

            html_content += "</div>"
            
            # Display scrollable container in Streamlit
            st.markdown(html_content, unsafe_allow_html=True)
            
    finally:
        try:
            os.unlink(tmp_file_path)
        except:
            pass

def page_to_image(page_number, pdf_path, width=800):
    """Convert a PDF page to an image"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[page_number]
            # Convert page to image
            img = page.to_image().original
            
            # Calculate height maintaining aspect ratio
            aspect_ratio = img.height / img.width
            height = int(width * aspect_ratio)
            
            # Resize image
            resized_image = img.resize((width, height))
            return resized_image
            
    except Exception as e:
        st.error(f"Error rendering page {page_number + 1}: {str(e)}")
        return None

def main():
    st.title("PDF Content Viewer")
    st.write("Upload a PDF file to view its content.")

    uploaded_file = st.file_uploader(
        "Choose a PDF file", 
        type=["pdf"],
        help="Upload a PDF file"
    )

    if uploaded_file is not None:
        # Create temporary directory for processing
        with tempfile.TemporaryDirectory() as temp_dir:
            with st.spinner("Processing file..."):
                # Get file content
                pdf_content = uploaded_file.getvalue()
                
                # Analyze the PDF content
                analyze_pdf(pdf_content)

if __name__ == "__main__":
    main()

# =============================================================================================
# =============================================================================================

import streamlit as st
from pathlib import Path
import tempfile
import os
import win32com.client
import pythoncom
import mammoth
import time

def convert_doc_to_docx(input_path):
    st.write("file path is: ", input_path)
    """Convert DOC file to DOCX"""
    pythoncom.CoInitialize()
    word = None
    temp_output_path = None
    
    try:
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(input_path)
        temp_output_path = str(Path(input_path).parent / f"temp_{int(time.time())}.docx")
        doc.SaveAs2(temp_output_path, FileFormat=16)  # 16 is the format for docx
        doc.Close()
        
        with open(temp_output_path, 'rb') as f:
            content = f.read()
        return content
    except Exception as e:
        st.error(f"Error converting file: {str(e)}")
        return None
    finally:
        if word:
            word.Quit()
        pythoncom.CoUninitialize()
        if temp_output_path and os.path.exists(temp_output_path):
            try:
                os.remove(temp_output_path)
            except:
                pass

def analyze_document(file_content):
    """Display document content preserving original formatting"""
    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
        tmp_file.write(file_content)
        tmp_file_path = tmp_file.name

    try:
        # Custom style map to preserve more formatting
        style_map = """
        p[style-name='Normal'] => p:fresh
        p[style-name='Heading 1'] => h1:fresh
        p[style-name='Heading 2'] => h2:fresh
        p[style-name='Heading 3'] => h3:fresh
        r[style-name='Strong'] => strong
        r[style-name='Emphasis'] => em
        """

        # Convert DOCX to HTML with mammoth
        with open(tmp_file_path, "rb") as docx_file:
            result = mammoth.convert_to_html(docx_file, style_map=style_map)
            html = result.value

        # Add custom CSS for better document display
        custom_css = """
        <style>
            .document-container {
                width: 100%;
                height: 600px;
                overflow-y: scroll;
                padding: 40px;
                background-color: white;
                border: 1px solid #ccc;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            }
            .document-content {
                max-width: 800px;
                margin: 0 auto;
                font-family: 'Calibri', sans-serif;
                line-height: 1.5;
            }
            table {
                border-collapse: collapse;
                width: 100%;
                margin: 10px 0;
            }
            td, th {
                border: 1px solid #ddd;
                padding: 8px;
            }
            img {
                max-width: 100%;
                height: auto;
            }
            h1 { font-size: 24px; margin: 20px 0; }
            h2 { font-size: 20px; margin: 16px 0; }
            h3 { font-size: 16px; margin: 14px 0; }
            p { margin: 12px 0; }
        </style>
        """

        # Combine CSS and HTML content
        full_html = f"""
        {custom_css}
        <div class="document-container">
            <div class="document-content">
                {html}
            </div>
        </div>
        """

        # Display the document
        st.markdown(full_html, unsafe_allow_html=True)

    except Exception as e:
        st.error(f"Error processing document: {str(e)}")
    finally:
        try:
            os.unlink(tmp_file_path)
        except:
            pass

def main():
    st.title("Word Document Viewer")
    st.write("Upload a Word file (DOC or DOCX) to view its content.")

    uploaded_file = st.file_uploader(
        "Choose a Word file", 
        type=["doc", "docx"],
        help="Upload a DOC or DOCX file"
    )

    if uploaded_file is not None:
        file_extension = Path(uploaded_file.name).suffix.lower()
        
        with st.spinner("Processing file..."):
            if file_extension == '.doc':
                with tempfile.NamedTemporaryFile(delete=False, suffix='.doc') as tmp_file:
                    tmp_file.write(uploaded_file.getvalue())
                    tmp_file_path = tmp_file.name
                
                try:
                    docx_content = convert_doc_to_docx(tmp_file_path)
                    if docx_content is None:
                        st.error("Failed to convert DOC file.")
                        st.stop()
                finally:
                    try:
                        os.unlink(tmp_file_path)
                    except:
                        pass
            else:
                docx_content = uploaded_file.getvalue()

            # Display the document content
            analyze_document(docx_content)

if __name__ == "__main__":
    main()