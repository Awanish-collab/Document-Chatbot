import streamlit as st
import os
import time
from collections import defaultdict
import hashlib
from datetime import datetime
import base64
import fitz  # PyMuPDF for better PDF handling

class FileStatus:
    PROCESSING = "processing"
    COMPLETED = "completed"
    FAILED = "failed"

class DocumentProcessor:
    MAX_PREVIEW_SIZE = 10 * 1024 * 1024  # 10MB limit for preview
    MAX_PDF_PAGES_PREVIEW = 5  # Maximum pages to show in preview

def get_file_hash(file_content):
        return hashlib.md5(file_content).hexdigest()

def initialize_session_state():
    if 'messages' not in st.session_state:
        st.session_state.messages = []
    if 'uploaded_files' not in st.session_state:
        st.session_state.uploaded_files = defaultdict(list)
    if 'file_contents' not in st.session_state:
        st.session_state.file_contents = {}
    if 'file_hashes' not in st.session_state:
        st.session_state.file_hashes = {}
    if 'file_status' not in st.session_state:
        st.session_state.file_status = {}
    if 'file_metadata' not in st.session_state:
        st.session_state.file_metadata = {}
    if 'file_data' not in st.session_state:
        st.session_state.file_data = {}

def truncate_filename(filename, max_length=15):
    name, ext = os.path.splitext(filename)
    if len(name) > max_length:
        return name[:max_length] + "..." + ext
    return filename

def get_file_type_icon(file_type):
    icons = {
        'pdf': '📄',
        'docx': '📝',
        'mp4': '🎥',
        'pptx': '📊',
        'mp3': '🎵',
        'wav': '🎵'
    }
    return icons.get(file_type, '📎')

def get_file_type(filename):
    return filename.split('.')[-1].lower()

def format_size(size):
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size < 1024:
            return f"{size:.2f} {unit}"
        size /= 1024
    return f"{size:.2f} TB"

def create_file_preview(filename):
    file_type = get_file_type(filename)
    file_data = st.session_state.file_data.get(filename)

    try:
        if file_type == 'pdf':
            # For smaller PDFs, show the full file
            data_url = f"data:application/pdf;base64,{base64.b64encode(file_data).decode()}"
            #st.write("data url: ", data_url)
            return f'<iframe src="{data_url}" width="100%" height="500px" style="border: 1px solid #ddd; border-radius: 5px;"></iframe>'
    
        elif file_type == 'pptx' or file_type == 'ppt':
            st.write("filename is: ", filename)
            return "PPT file preview not yet implemented"
        
        elif file_type in ['mp4', 'mp3', 'wav']:
            return "MP4 file preview not yet implemented"

        elif file_type == 'docx':
            return "DOCX file preview not yet implemented" 
            
    except Exception as e:
        return f"Preview generation error: {str(e)}"

def process_file(file):
    file_type = get_file_type(file.name)
    st.session_state.file_status[file.name] = FileStatus.PROCESSING
    
    try:
        st.session_state.file_status[file.name] = FileStatus.COMPLETED
        st.session_state.file_metadata[file.name] = {
            'processed_at': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'size': len(file.getvalue()),
            'type': file_type
        }
        
        # Store the file data for preview
        st.session_state.file_data[file.name] = file.getvalue()
        #st.write("process file get value: ", file.getvalue())
        
        return "File successfully processed"

    except Exception as e:
        st.session_state.file_status[file.name] = FileStatus.FAILED
        st.session_state.file_metadata[file.name] = {
            'error_message': str(e),
            'failed_at': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'type': file_type
        }
        raise

def handle_chat_input(prompt):
    if not st.session_state.file_contents:
        return "Please upload and process some documents before asking questions."

    # Clear previous messages when new files are uploaded
    st.session_state.messages = []
    
    response = "Based on the processed documents:\n\n"
    for file_type, filenames in st.session_state.uploaded_files.items():
        successfully_processed = [
            f for f in filenames 
            if st.session_state.file_status.get(f) == FileStatus.COMPLETED
        ]
        if successfully_processed:
            response += f"\n{get_file_type_icon(file_type)} {file_type.upper()} files:\n"
            for filename in successfully_processed:
                content = st.session_state.file_contents[filename]
                content_preview = content[0][:100] + "..." if isinstance(content, tuple) else content[:100] + "..."
                response += f"- {truncate_filename(filename)}\n  Preview: {content_preview}\n"
    
    return response

def main():
    st.set_page_config(
        page_title="Document Q&A Chatbot",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    initialize_session_state()

    with st.sidebar:
        st.title("📁 Document Manager")
        
        uploaded_files = st.file_uploader(
            "Upload your documents",
            type=['pdf', 'docx', 'mp4', 'pptx', 'ppt', 'mp3', 'wav'],
            accept_multiple_files=True
        )

        if uploaded_files:
            for file in uploaded_files:
                file_content = file.read()
                file_hash = get_file_hash(file_content)
                file.seek(0)
                
                if file.name in st.session_state.file_hashes and st.session_state.file_hashes[file.name] == file_hash:
                    st.warning(f'{truncate_filename(file.name)} is a duplicate and was skipped.')
                    continue
                
                st.session_state.file_hashes[file.name] = file_hash
                file_type = get_file_type(file.name)
                
                with st.spinner(f'Processing {truncate_filename(file.name)}...'):
                    try:
                        text_content = process_file(file)
                        if file_type not in st.session_state.uploaded_files:
                            st.session_state.uploaded_files[file_type] = []
                        if file.name not in st.session_state.uploaded_files[file_type]:
                            st.session_state.uploaded_files[file_type].append(file.name)
                        st.session_state.file_contents[file.name] = text_content
                        st.success(f'Successfully processed {truncate_filename(file.name)}')
                    except Exception as e:
                        st.error(f'Error processing {truncate_filename(file.name)}: {str(e)}')
                        if file.name in st.session_state.file_hashes:
                            del st.session_state.file_hashes[file.name]

        if st.session_state.uploaded_files:
            st.write("### Processed Files")
            
            file_types = list(st.session_state.uploaded_files.keys())
            if file_types:
                tabs = st.tabs([f"{get_file_type_icon(ft)} {ft.upper()}" for ft in file_types])
                
                for tab, file_type in zip(tabs, file_types):
                    with tab:
                        for filename in st.session_state.uploaded_files[file_type]:
                            col1, col2 = st.columns([3, 1])
                            with col1:
                                status = st.session_state.file_status.get(filename, "Unknown")
                                metadata = st.session_state.file_metadata.get(filename, {})
                                
                                # Show file info with better formatting
                                st.markdown(f"""
                                **{filename}**  
                                Status: {status}  
                                Size: {format_size(metadata.get('size', 0))}  
                                Processed: {metadata.get('processed_at', 'N/A')}
                                """)
                                
                                if status == FileStatus.FAILED:
                                    st.error(f"Error: {metadata.get('error_message', 'Unknown error')}")
                            
                            with col2:
                                # Preview button with unique key
                                preview_key = f"preview_{filename}"
                                preview_btn = st.button("👁️", key=preview_key)
                            
                            if preview_btn:
                                st.markdown("### File Preview")
                                preview_html = create_file_preview(filename)
                                st.markdown(preview_html, unsafe_allow_html=True)

            st.write("### Summary")
            total_files = sum(len(files) for files in st.session_state.uploaded_files.values())
            completed_files = sum(1 for status in st.session_state.file_status.values() 
                                if status == FileStatus.COMPLETED)
            failed_files = sum(1 for status in st.session_state.file_status.values() 
                             if status == FileStatus.FAILED)
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Files", total_files)
            with col2:
                st.metric("Successfully Processed", completed_files)
            with col3:
                st.metric("Failed", failed_files)

            if st.button("Clear All Files"):
                st.session_state.uploaded_files = defaultdict(list)
                st.session_state.file_contents = {}
                st.session_state.file_hashes = {}
                st.session_state.file_status = {}
                st.session_state.file_metadata = {}
                st.session_state.file_data = {}
                st.session_state.messages = []
                st.experimental_rerun()

    st.title("💬 Document Q&A Chatbot")
    
    # Display initial instructions if no files are uploaded
    if not st.session_state.uploaded_files:
        st.info("👈 Please start by uploading documents in the sidebar")
    else:
        # Display chat messages
        for message in st.session_state.messages:
            with st.chat_message(message["role"]):
                st.write(message["content"])

        if prompt := st.chat_input("Ask a question about your documents"):
            st.session_state.messages.append({"role": "user", "content": prompt})
            
            with st.chat_message("user"):
                st.write(prompt)

            with st.chat_message("assistant"):
                message_placeholder = st.empty()
                full_response = handle_chat_input(prompt)
                
                # Simulate typing
                displayed_response = ""
                for chunk in full_response.split():
                    displayed_response += chunk + " "
                    time.sleep(0.05)
                    message_placeholder.markdown(displayed_response + "▌")
                message_placeholder.markdown(displayed_response)
            
            st.session_state.messages.append({"role": "assistant", "content": full_response})

if __name__ == "__main__":
    main()