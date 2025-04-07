import streamlit as st
import os
import tempfile
import time
import json
from datetime import datetime
import glob
import traceback
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

from ppt_reader import read_ppt, extract_content_with_mapping
from ppt_processor import PPTProcessor
from utils import ensure_dir

# Set up page configuration
st.set_page_config(
    page_title="PPT Regenerator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

def init_session_state():
    """Initialize session state variables if they don't exist."""
    if "processing_results" not in st.session_state:
        st.session_state.processing_results = None
    if "ppt_info" not in st.session_state:
        st.session_state.ppt_info = None
    if "content_map" not in st.session_state:
        st.session_state.content_map = None
    if "before_after" not in st.session_state:
        st.session_state.before_after = []
    if "active_tab" not in st.session_state:
        st.session_state.active_tab = "Upload"
    if "output_file_created" not in st.session_state:
        st.session_state.output_file_created = False
    if "output_bytes" not in st.session_state:
        st.session_state.output_bytes = None

def main():
    init_session_state()
    
    st.title("PPT Regenerator")
    
    # Sidebar for configuration
    with st.sidebar:
        st.header("Settings")
        
        # Mock mode for development
        use_mock = st.checkbox("Use Mock LLM (for testing)", 
                              value=False,
                              help="Use fake responses instead of calling the API",
                              key="use_mock_checkbox")
        
        # Get API key from environment only (no UI element)
        api_key = os.getenv("DEEPSEEK_API_KEY", "")
        
        # Processing settings
        st.subheader("Processing Settings")
        
        # Debug settings
        debug_mode = st.checkbox("Enable Debug Mode", value=True,
                               help="Save detailed logs of each processing step",
                               key="debug_mode_checkbox")
    
    # Create main navigation tabs
    main_tabs = st.tabs(["Upload & Process", "View & Analyze", "Results"])
    
    with main_tabs[0]:  # Upload & Process tab
        st.header("Upload PowerPoint")
        
        # Get user information
        user_info = st.text_area(
            "Tell us about your industry or use case (optional)",
            placeholder="Example: I work in healthcare technology focused on patient engagement. Our target audience is hospital administrators and clinicians.",
            help="This information will help tailor the content to your specific needs.",
            key="user_info_textarea"
        )
        
        # File uploader for PowerPoint files
        uploaded_file = st.file_uploader("Upload a PowerPoint file", type=["pptx"], key="pptx_uploader")
        
        if uploaded_file is not None:
            # Create a temporary file to save the uploaded content
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_file_path = tmp_file.name
            
            try:
                # Read the PowerPoint file for basic info
                with st.spinner("Analyzing PowerPoint file..."):
                    try:
                        ppt_info = read_ppt(tmp_file_path, debug=debug_mode)
                        content_map = extract_content_with_mapping(ppt_info)
                        
                        # Store in session state
                        st.session_state.ppt_info = ppt_info
                        st.session_state.content_map = content_map
                        
                        # Show basic info
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("Slides", ppt_info['slide_count'])
                        with col2:
                            st.metric("Text Elements", sum(len(slide.get('texts', [])) for slide in ppt_info['slides']))
                        with col3:
                            st.metric("Size", f"{len(uploaded_file.getvalue()) / 1024:.1f} KB")
                            
                    except Exception as e:
                        st.error(f"Error analyzing PowerPoint: {str(e)}")
                        if debug_mode:
                            st.error(f"Detailed error: {traceback.format_exc()}")
                
                # Create a prominent regenerate button
                st.markdown("### Generate New Content")
                if st.button("🔄 Regenerate PowerPoint Content", type="primary", key="regenerate_button", use_container_width=True):
                    # Get API key from environment only
                    api_key = os.getenv("DEEPSEEK_API_KEY", "")
                    
                    if not use_mock and not api_key:
                        st.error("⚠️ No API key found in .env file. Please add DEEPSEEK_API_KEY to your .env file.")
                        st.code("Add this to your .env file:\nDEEPSEEK_API_KEY=your_api_key_here", language="text")
                    else:
                        # Switch to the Results tab
                        st.session_state.active_tab = "Results"
                        st.experimental_rerun()
            
            except Exception as e:
                st.error(f"Error processing PowerPoint file: {str(e)}")
                if debug_mode:
                    st.error(f"Detailed error: {traceback.format_exc()}")
        else:
            st.info("Please upload a PowerPoint file to begin.")
    
    with main_tabs[1]:  # View & Analyze tab
        st.header("Analyze Presentation")
        
        if st.session_state.ppt_info is not None:
            ppt_info = st.session_state.ppt_info
            content_map = st.session_state.content_map
            
            st.success(f"Successfully analyzed PowerPoint with {ppt_info['slide_count']} slides!")
            
            # Create tabs for viewing content
            view_tabs = st.tabs(["Slide Content", "Structure Map", "Raw Data"])
            
            with view_tabs[0]:
                # Display all slides sequentially
                for slide in ppt_info["slides"]:
                    st.markdown(f"## Slide {slide['slide_number']} - {slide['slide_layout']}")
                    
                    if slide['texts']:
                        for i, text in enumerate(slide['texts']):
                            st.text_area(
                                f"Text {i+1}",
                                value=text, 
                                height=100,
                                disabled=True,
                                key=f"slide_{slide['slide_number']}_text_{i}"
                            )
                    else:
                        st.write("No text content found in this slide.")
                    
                    # Add a divider between slides
                    if slide['slide_number'] < ppt_info['slide_count']:
                        st.markdown("---")
            
            with view_tabs[1]:
                # Show structure mapping
                st.subheader("Content Mapping Structure")
                st.json(content_map)
            
            with view_tabs[2]:
                # Show raw data
                st.subheader("Raw PowerPoint Data")
                st.json(ppt_info)
        else:
            st.info("Please upload a PowerPoint file in the 'Upload & Process' tab first.")
    
    with main_tabs[2]:  # Results tab
        st.header("Regeneration Results")
        
        if "output_file_created" in st.session_state and st.session_state.output_file_created:
            # We already have results, just display them without reprocessing
            if "processing_results" in st.session_state and st.session_state.processing_results:
                results = st.session_state.processing_results
                st.success(f"Successfully processed {results['total_slides']} slides in {results['sections']} sections!")
                
                # Provide download button using stored bytes
                st.download_button(
                    label="📥 Download Regenerated PowerPoint",
                    data=st.session_state.output_bytes,
                    file_name="regenerated_presentation.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    key="download_button",
                    use_container_width=True
                )
                
                if st.button("Process Again", key="process_again_button"):
                    # Reset for new processing
                    st.session_state.output_file_created = False
                    st.session_state.active_tab = "Upload"
                    st.experimental_rerun()
                
                # Continue with showing results as before
                if st.session_state.before_after:
                    results_tabs = st.tabs(["Before & After", "Statistics", "Debug Info"])
                    
                    with results_tabs[0]:
                        st.subheader("Content Comparison")
                        
                        # Group slides by 5 for pagination
                        slide_changes = st.session_state.before_after
                        slides_per_page = 5
                        total_pages = (len(slide_changes) + slides_per_page - 1) // slides_per_page
                        
                        # Add pagination
                        page = st.selectbox("Page", range(1, total_pages + 1), format_func=lambda x: f"Page {x} of {total_pages}", key="pagination")
                        start_idx = (page - 1) * slides_per_page
                        end_idx = min(start_idx + slides_per_page, len(slide_changes))
                        
                        # Show slides for current page
                        for slide_idx in range(start_idx, end_idx):
                            slide_changes_data = slide_changes[slide_idx]
                            slide_num = slide_changes_data.get("slide_number", "Unknown")
                            changes = slide_changes_data.get("changes", [])
                            
                            if changes:
                                st.markdown(f"### Slide {slide_num}")
                                for i, change in enumerate(changes):
                                    col1, col2 = st.columns(2)
                                    
                                    with col1:
                                        st.markdown("**Original:**")
                                        st.text_area(
                                            label="",
                                            value=change.get("before", ""),
                                            height=120,
                                            disabled=True,
                                            key=f"before_{slide_num}_{i}_{slide_idx}"
                                        )
                                    
                                    with col2:
                                        st.markdown("**Regenerated:**")
                                        st.text_area(
                                            label="",
                                            value=change.get("after", ""),
                                            height=120,
                                            disabled=True,
                                            key=f"after_{slide_num}_{i}_{slide_idx}"
                                        )
                                
                                st.markdown("---")
                    
                    with results_tabs[1]:
                        # Show processing statistics
                        st.subheader("Processing Statistics")
                        
                        # Create statistics columns
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.metric("Total Slides", results['total_slides'])
                            st.metric("Number of Sections", results['sections'])
                            st.metric("Total Processing Time", f"{results['total_duration']:.2f}s")
                        
                        with col2:
                            if "key_concepts" in results:
                                st.write("Key Concepts Identified:")
                                concepts = list(results["key_concepts"].keys())
                                if concepts:
                                    for concept in concepts[:10]:  # Show top 10
                                        st.write(f"- {concept}")
                                    if len(concepts) > 10:
                                        st.write(f"- Plus {len(concepts) - 10} more...")
                                else:
                                    st.write("No key concepts identified")
                        
                        # Section details
                        st.subheader("Section Processing Details")
                        for i, section in enumerate(results.get("section_details", [])):
                            with st.expander(f"Section {i+1} - {len(section.get('slide_numbers', []))} slides"):
                                st.write(f"Processing time: {section.get('duration', 0):.2f}s")
                                st.write(f"Slides: {', '.join(map(str, section.get('slide_numbers', [])))}")
                                st.write(f"Regenerated text blocks: {section.get('regenerated_texts_count', 0)}")
                    
                    with results_tabs[2]:
                        # Debug information
                        if debug_mode:
                            st.subheader("Debug Information")
                            if "llm_debug_info" in results:
                                with st.expander("LLM Service Debug Info"):
                                    st.json(results["llm_debug_info"])
                            
                            # View debug files
                            debug_dir = "debug_logs"
                            ensure_dir(debug_dir)
                            debug_files = glob.glob(f"{debug_dir}/*.json")
                            if debug_files:
                                selected_file = st.selectbox(
                                    "Select debug file to view", 
                                    debug_files,
                                    format_func=lambda x: os.path.basename(x),
                                    key="debug_file_selectbox"
                                )
                                
                                if st.button("View Selected Debug File", key="view_debug_button"):
                                    with open(selected_file, 'r') as f:
                                        debug_data = json.load(f)
                                    st.json(debug_data)
        
        elif uploaded_file is not None and st.session_state.active_tab == "Results":
            # Process the PowerPoint
            try:
                with st.spinner("Regenerating PowerPoint content..."):
                    # Get API key from environment only
                    api_key = os.getenv("DEEPSEEK_API_KEY", "")
                    
                    # Create a processor
                    processor = PPTProcessor(
                        use_mock=use_mock, 
                        api_key=api_key,  # Pass the API key from environment
                        debug=debug_mode,
                        max_slides_per_section=50,  # Hardcoded value
                        max_total_slides=500  # Hardcoded total slides limit
                    )
                    
                    # Create output path
                    timestamp = int(time.time())
                    output_filename = f"regenerated_{timestamp}.pptx"
                    output_dir = os.path.dirname(tmp_file_path)
                    output_path = os.path.join(output_dir, output_filename)
                    
                    # Progress tracking
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    # Start progress monitoring
                    status_text.text("Starting processing...")
                    
                    # Process the presentation
                    results = processor.process_presentation(
                        tmp_file_path, 
                        output_path,
                        max_slides_per_section=50,  # Hardcoded value
                        user_info=user_info
                    )
                    
                    # Store results in session state
                    st.session_state.processing_results = results
                    
                    # Store before/after in session state
                    if "before_after" in results:
                        st.session_state.before_after = results["before_after"]
                    
                    # Update progress
                    progress_bar.progress(100)
                    status_text.text("Processing complete!")
                
                # Display results
                if os.path.exists(output_path):
                    st.success(f"Successfully processed {results['total_slides']} slides in {results['sections']} sections!")
                    
                    # Provide download button for the modified file
                    with open(output_path, "rb") as file:
                        modified_ppt_bytes = file.read()
                    
                    # Store the output bytes in session state to avoid reprocessing when downloading
                    st.session_state.output_bytes = modified_ppt_bytes
                    st.session_state.output_file_created = True
                    
                    # Refresh the page to show results from session state
                    st.experimental_rerun()
                else:
                    st.error("Processing failed. No output file was generated.")
            
            except Exception as e:
                st.error(f"Error during regeneration: {str(e)}")
                if debug_mode:
                    st.error(f"Detailed error: {traceback.format_exc()}")
        else:
            st.info("Upload a PowerPoint and click 'Regenerate PowerPoint Content' to see results here.")
    
    # Debug info display
    if debug_mode and hasattr(st.session_state, 'debug_view') and st.session_state.debug_view:
        with st.expander("Debug Data View", expanded=False):
            st.json(st.session_state.debug_view)

if __name__ == "__main__":
    main()