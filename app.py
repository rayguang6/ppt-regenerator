import streamlit as st
import os
import tempfile
import time
from datetime import datetime
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

from ppt_reader import read_ppt, extract_content_with_mapping
from ppt_processor import PPTProcessor
from utils import ensure_dir

# Custom CSS for improved appearance
custom_css = """
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: 700;
        color: #1E88E5;
        margin-bottom: 0.5rem;
    }
    .divider {
        margin: 1.5rem 0;
        border: none;
        height: 1px;
        background-color: #E0E0E0;
    }
    footer {
        margin-top: 2rem;
        padding-top: 1rem;
        border-top: 1px solid #E0E0E0;
        text-align: center;
        color: #757575;
        font-size: 0.9rem;
    }
</style>
"""

# Set up page configuration
st.set_page_config(
    page_title="PPT Regenerator",
    page_icon="📊",
    layout="wide",
    # initial_sidebar_state="collapsed"
)

# Apply custom CSS
st.markdown(custom_css, unsafe_allow_html=True)

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
    if "regeneration_started" not in st.session_state:
        st.session_state.regeneration_started = False
    if "current_slide_processing" not in st.session_state:
        st.session_state.current_slide_processing = 0

def main():
    init_session_state()
    
    # Custom header with logo-like styling
    st.markdown('<div class="main-header">📊 PPT Regenerator</div>', unsafe_allow_html=True)
    
    # Get API key from environment
    api_key = os.getenv("DEEPSEEK_API_KEY", "")
    
    # Create main navigation tabs with simpler naming
    main_tabs = st.tabs(["Upload", "Analyze", "Results"])
    
    with main_tabs[0]:  # Upload tab
        # File uploader with clearer instructions
        uploaded_file = st.file_uploader("Upload a PowerPoint file", type=["pptx"], key="pptx_uploader")
        
        # Simplified user input focusing on essentials
        st.subheader("Tell us about your course")

        course_topic = st.text_input(
            "What specific skill or knowledge does your course teach?",
            placeholder="e.g., Day trading stocks, Watercolor painting, Facebook ads, etc.",
            help="Be specific about what students will learn"
        )

        main_outcome = st.text_area(
            "What's the #1 result your students will achieve?",
            placeholder="e.g., Create profitable trading systems with just 30 minutes per day",
            help="The primary transformation or outcome students will experience"
        )

        # Optional third field for better results
        target_audience = st.text_input(
            "Who is this course for? (optional)",
            placeholder="e.g., Busy professionals, Beginners with no experience, etc.",
            help="Your ideal student profile"
        )

        # Combine all inputs into user_info for the prompt
        user_info = f"""
        COURSE INFORMATION:
        What This Course Teaches: {course_topic}

        PRIMARY OUTCOME:
        {main_outcome}
        """

        if target_audience:
            user_info += f"\nTARGET AUDIENCE:\n{target_audience}"
        
        if uploaded_file is not None:
            # Create a temporary file to save the uploaded content
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_file_path = tmp_file.name
            
            try:
                # Read the PowerPoint file for basic info
                with st.spinner("Analyzing your PowerPoint file..."):
                    try:
                        ppt_info = read_ppt(tmp_file_path)
                        content_map = extract_content_with_mapping(ppt_info)
                        
                        # Store in session state
                        st.session_state.ppt_info = ppt_info
                        st.session_state.content_map = content_map
                        
                        # Show basic info with improved styling
                        st.success(f"PowerPoint file analyzed successfully! Found {ppt_info['slide_count']} slides.")
                            
                    except Exception as e:
                        st.error(f"Error analyzing PowerPoint: {str(e)}")
                
                # Create a prominent regenerate button
                st.markdown('<hr class="divider">', unsafe_allow_html=True)
                
                # Section for regeneration controls and progress
                regenerate_section = st.container()
                
                with regenerate_section:
                    if not st.session_state.regeneration_started and not st.session_state.output_file_created:
                        if st.button("Regenerate PowerPoint Content", type="primary", key="regenerate_button", use_container_width=True):
                            # Get API key from environment only
                            api_key = os.getenv("DEEPSEEK_API_KEY", "")
                            
                            if not api_key:
                                st.error("⚠️ No API key found in .env file. Please add DEEPSEEK_API_KEY to your .env file.")
                                st.code("Add this to your .env file:\nDEEPSEEK_API_KEY=your_api_key_here", language="text")
                            else:
                                # Process the PowerPoint immediately on this tab
                                st.session_state.regeneration_started = True
                                st.rerun()
                    
                    # If regeneration has been started, show progress
                    if st.session_state.regeneration_started and not st.session_state.output_file_created:
                        # Process the PowerPoint
                        try:
                            progress_container = st.container()
                            with progress_container:
                                st.subheader("Regenerating PowerPoint")
                                progress_bar = st.progress(0)
                                status_text = st.empty()
                                
                                with st.spinner("Processing your presentation..."):
                                    # Get API key from environment only
                                    api_key = os.getenv("DEEPSEEK_API_KEY", "")
                                    
                                    # Create a processor
                                    processor = PPTProcessor(
                                        api_key=api_key,
                                        max_slides_per_section=50,  # Use default section size for better results
                                        max_total_slides=500
                                    )
                                    
                                    # Create output path
                                    timestamp = int(time.time())
                                    output_filename = f"regenerated_{timestamp}.pptx"
                                    output_dir = os.path.dirname(tmp_file_path)
                                    output_path = os.path.join(output_dir, output_filename)
                                    
                                    # Start progress monitoring
                                    status_text.text("Starting processing...")
                                    
                                    # Define a callback for slide processing updates
                                    def progress_callback(current_slide, total_slides):
                                        if current_slide == 0:
                                            progress_value = 0
                                            status_message = "Preparing to process slides..."
                                        else:
                                            progress_value = min(current_slide / total_slides, 1.0)
                                            status_message = f"Processing slide {current_slide} of {total_slides}..."
                                            
                                        progress_bar.progress(progress_value)
                                        status_text.text(status_message)
                                        st.session_state.current_slide_processing = current_slide
                                    
                                    # Process the presentation
                                    results = processor.process_presentation(
                                        tmp_file_path, 
                                        output_path,
                                        max_slides_per_section=50,  # Use default section size
                                        user_info=user_info,
                                        progress_callback=progress_callback
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
                                    # Provide download button for the modified file
                                    with open(output_path, "rb") as file:
                                        modified_ppt_bytes = file.read()
                                    
                                    # Store the output bytes in session state
                                    st.session_state.output_bytes = modified_ppt_bytes
                                    st.session_state.output_file_created = True
                                    st.session_state.regeneration_started = False
                                    
                                    # Success message after processing
                                    st.success(f"Successfully processed {results['total_slides']} slides!")
                                    
                                    # Refresh to show download button
                                    st.rerun()
                                else:
                                    st.error("Processing failed. No output file was generated.")
                                    st.session_state.regeneration_started = False
                        except Exception as e:
                            st.error(f"Error during regeneration: {str(e)}")
                            st.session_state.regeneration_started = False
                    
                    # Show download button after processing is complete
                    if st.session_state.output_file_created and st.session_state.output_bytes is not None:
                        st.success(f"PowerPoint regeneration complete! Your presentation is ready for download.")
                        
                        # Provide download button using stored bytes
                        st.download_button(
                            label="Download Regenerated PowerPoint",
                            data=st.session_state.output_bytes,
                            file_name="regenerated_presentation.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            key="download_button_upload_tab",
                            use_container_width=True
                        )
                        
                        # Show processing statistics
                        if "processing_results" in st.session_state and st.session_state.processing_results:
                            results = st.session_state.processing_results
                            
                            # Format processing time nicely
                            total_time = results.get("total_duration", 0)
                            mins = int(total_time // 60)
                            secs = int(total_time % 60)
                            time_str = f"{mins}m {secs}s" if mins > 0 else f"{secs}s"
                            
                            # Create two columns for stats
                            col1, col2 = st.columns(2)
                            
                            with col1:
                                st.metric("Total Processing Time", time_str)
                                st.metric("Slides Processed", results.get("total_slides", 0))
                            
                            with col2:
                                st.metric("Sections", results.get("sections", 0))
                                st.metric("Text Elements Changed", sum(len(slide.get("changes", [])) for slide in results.get("before_after", [])))
                            
                            # Show any warnings if present
                            if "warnings" in results and results["warnings"]:
                                with st.expander("Processing Warnings"):
                                    for warning in results["warnings"]:
                                        st.warning(warning)
                        
                        # Option to start over
                        if st.button("Start Over", key="start_over_button"):
                            # Reset the session state
                            st.session_state.output_file_created = False
                            st.session_state.regeneration_started = False
                            st.session_state.output_bytes = None
                            st.session_state.processing_results = None
                            st.session_state.before_after = []
                            st.rerun()
                            
                        # Prompt to view detailed results
                        st.info("For detailed content comparison, check the Results tab.")
            
            except Exception as e:
                st.error(f"Error processing PowerPoint file: {str(e)}")
        else:
            # Show a simple upload prompt
            st.info("Please upload a PowerPoint file to begin.")
    
    with main_tabs[1]:  # Analyze tab
        if st.session_state.ppt_info is not None:
            ppt_info = st.session_state.ppt_info
            
            st.success(f"Successfully analyzed presentation with {ppt_info['slide_count']} slides")
            
            # Simply show all slide content without tabs
            st.subheader("Slide Content Preview")
            
            # Display all slides
            for slide in ppt_info["slides"]:
                st.markdown(f"### Slide {slide['slide_number']} - {slide['slide_layout']}")
                
                if slide['texts']:
                    # Combine all text elements into a single text area
                    combined_text = "\n\n".join(slide['texts'])
                    st.text_area(
                        "Slide content",
                        value=combined_text, 
                        height=150,
                        disabled=True,
                        key=f"slide_{slide['slide_number']}_text"
                    )
                else:
                    st.info("No text content found in this slide.")
                
                st.markdown('<hr class="divider">', unsafe_allow_html=True)
            
        else:
            # Show a helpful prompt
            st.info("Please upload a PowerPoint file in the Upload tab first.")
    
    with main_tabs[2]:  # Results tab
        if "output_file_created" in st.session_state and st.session_state.output_file_created:
            # We already have results, just display them without reprocessing
            if "processing_results" in st.session_state and st.session_state.processing_results:
                results = st.session_state.processing_results
                
                # Show success message with statistics
                st.success(f"Successfully processed {results['total_slides']} slides")
                
                # Provide download button using stored bytes
                st.download_button(
                    label="Download Regenerated PowerPoint",
                    data=st.session_state.output_bytes,
                    file_name="regenerated_presentation.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    key="download_button",
                    use_container_width=True
                )
                
                # Show processing statistics
                if results:
                    # Format processing time nicely
                    total_time = results.get("total_duration", 0)
                    mins = int(total_time // 60)
                    secs = int(total_time % 60)
                    time_str = f"{mins}m {secs}s" if mins > 0 else f"{secs}s"
                    
                    # Create two columns for stats
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.metric("Total Processing Time", time_str)
                        st.metric("Slides Processed", results.get("total_slides", 0))
                    
                    with col2:
                        st.metric("Sections", results.get("sections", 0))
                        st.metric("Text Elements Changed", sum(len(slide.get("changes", [])) for slide in results.get("before_after", [])))
                    
                    # Show any warnings if present
                    if "warnings" in results and results["warnings"]:
                        with st.expander("Processing Warnings"):
                            for warning in results["warnings"]:
                                st.warning(warning)
                 
                st.button("Start Over", key="process_again_button", on_click=lambda: setattr(st.session_state, "output_file_created", False) or setattr(st.session_state, "active_tab", "Upload"))
                
                # Show before & after content without tabs
                if st.session_state.before_after:
                    st.subheader("Content Comparison")
                    
                    # Add column headers just once at the top
                    col1, col2 = st.columns(2)
                    with col1:
                        st.markdown("**Original**")
                    with col2:
                        st.markdown("**Regenerated**")
                    
                    st.markdown('<hr class="divider">', unsafe_allow_html=True)
                    
                    # Show all slides without pagination
                    slide_changes = st.session_state.before_after
                    
                    for slide_changes_data in slide_changes:
                        slide_num = slide_changes_data.get("slide_number", "Unknown")
                        changes = slide_changes_data.get("changes", [])
                        
                        if changes:
                            st.markdown(f"**Slide {slide_num}**")
                            
                            for i, change in enumerate(changes):
                                col1, col2 = st.columns(2)
                                
                                with col1:
                                    st.text_area(
                                        label=f"Original text {i+1} for slide {slide_num}",
                                        value=change.get("before", ""),
                                        height=120,
                                        disabled=True,
                                        key=f"before_{slide_num}_{i}",
                                        label_visibility="collapsed"
                                    )
                                
                                with col2:
                                    st.text_area(
                                        label=f"Regenerated text {i+1} for slide {slide_num}",
                                        value=change.get("after", ""),
                                        height=120,
                                        disabled=True,
                                        key=f"after_{slide_num}_{i}",
                                        label_visibility="collapsed"
                                    )
                            
                            st.markdown('<hr class="divider">', unsafe_allow_html=True)
        else:
            # Show helpful guide when no results yet
            st.info("Upload a PowerPoint file in the Upload tab and click 'Regenerate PowerPoint Content' to see results here.")
    
    # Add minimal footer
    st.markdown("""
    <footer>
        PPT Regenerator
    </footer>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()