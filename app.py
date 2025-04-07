import streamlit as st
import os
import tempfile
import time
from ppt_reader import read_ppt, modify_ppt_text

def main():
    st.set_page_config(
        page_title="PPT Regenerator",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("PPT Regenerator")
    st.subheader("PowerPoint Reader Test")
    
    # File uploader for PowerPoint files
    uploaded_file = st.file_uploader("Upload a PowerPoint file", type=["pptx"])
    
    if uploaded_file is not None:
        # Create a temporary file to save the uploaded content
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name
        
        try:
            # Read the PowerPoint file
            with st.spinner("Reading PowerPoint file..."):
                ppt_info = read_ppt(tmp_file_path)
            
            # Display basic information
            st.success(f"Successfully read PowerPoint with {ppt_info['slide_count']} slides!")
            
            # Add text modification section
            st.write("---")
            st.subheader("Modify PowerPoint Text")
            
            # Add prefix input
            prefix = st.text_input("Text Prefix", value="CHANGE", 
                                  help="This text will be added before each text element in the PowerPoint")
            
            # Add button to modify and download
            if st.button("Modify Text and Download"):
                with st.spinner("Modifying PowerPoint..."):
                    # Create output file path
                    timestamp = int(time.time())
                    output_filename = f"modified_{timestamp}.pptx"
                    output_dir = os.path.dirname(tmp_file_path)
                    output_path = os.path.join(output_dir, output_filename)
                    
                    try:
                        # Modify the PowerPoint
                        success = modify_ppt_text(tmp_file_path, output_path, prefix)
                        
                        if success:
                            # Read the modified file for download
                            with open(output_path, "rb") as file:
                                modified_ppt_bytes = file.read()
                            
                            # Provide download button
                            st.success("PowerPoint modified successfully!")
                            st.download_button(
                                label="Download Modified PowerPoint",
                                data=modified_ppt_bytes,
                                file_name=f"modified_presentation.pptx",
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                            )
                        else:
                            st.error("Failed to modify PowerPoint. Please try again.")
                    except Exception as e:
                        st.error(f"Error processing file: {str(e)}")
                    finally:
                        # Clean up the output file if it exists
                        if os.path.exists(output_path):
                            os.unlink(output_path)
            
            st.write("---")
            
            # Create tabs for different views
            tab1, tab2 = st.tabs(["Formatted View", "Raw Data"])
            
            with tab1:
                # Option to view all slides at once
                view_option = st.radio(
                    "View mode:", 
                    [
                        "All slides at once", 
                        "Expandable slides"
                    ], 
                    horizontal=True
                )
                
                # Display slide information
                if view_option == "Expandable slides":
                    for slide in ppt_info["slides"]:
                        with st.expander(f"Slide {slide['slide_number']}"):
                            if slide['texts']:
                                for i, text in enumerate(slide['texts']):
                                    st.text_area(
                                        f"Text {i+1}",
                                        value=text, 
                                        height=100,
                                        disabled=True,
                                        key=f"expander_slide_{slide['slide_number']}_text_{i}"
                                    )
                            else:
                                st.write("No text content found in this slide.")
                else:
                    # Show all slides without expanders
                    for slide in ppt_info["slides"]:
                        st.markdown(f"## Slide {slide['slide_number']}")
                        
                        if slide['texts']:
                            for i, text in enumerate(slide['texts']):
                                st.text_area(
                                    f"Text {i+1}",
                                    value=text, 
                                    height=100,
                                    disabled=True,
                                    key=f"all_slide_{slide['slide_number']}_text_{i}"
                                )
                        else:
                            st.write("No text content found in this slide.")
                        
                        # Add a divider between slides
                        if slide['slide_number'] < ppt_info['slide_count']:
                            st.markdown("---")
            
            with tab2:
                # Show raw data
                st.subheader("PowerPoint Data")
                st.json(ppt_info)
        
        except Exception as e:
            st.error(f"Error reading PowerPoint file: {str(e)}")
        
        finally:
            # Clean up the temporary file
            os.unlink(tmp_file_path)
    else:
        st.info("Please upload a PowerPoint file to view its content.")

if __name__ == "__main__":
    main()