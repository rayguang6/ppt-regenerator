from pptx import Presentation
import os

def read_ppt(file_path):
    """
    Read a PowerPoint file and return information about its content.
    
    Args:
        file_path (str): Path to the PowerPoint file
        
    Returns:
        dict: Dictionary containing information about the presentation
    """
    prs = Presentation(file_path)
    
    # Get presentation-level information
    presentation_info = {
        "slide_count": len(prs.slides),
        "slide_width": prs.slide_width,
        "slide_height": prs.slide_height,
        "slides": []
    }
    
    # Extract information from each slide
    for i, slide in enumerate(prs.slides):
        slide_info = {
            "slide_number": i + 1,
            "slide_id": slide.slide_id if hasattr(slide, "slide_id") else None,
            "slide_layout": slide.slide_layout.name if hasattr(slide.slide_layout, "name") else "Unknown",
            "shape_count": len(slide.shapes),
            "shapes": [],
            "texts": []
        }
        
        # Get information about each shape
        for j, shape in enumerate(slide.shapes):
            shape_info = {
                "shape_id": j + 1,
                "shape_type": str(shape.shape_type) if hasattr(shape, "shape_type") else "Unknown",
                "name": shape.name if hasattr(shape, "name") else "Unnamed",
                "has_text_frame": hasattr(shape, "text_frame"),
                "has_table": hasattr(shape, "table"),
                "has_chart": hasattr(shape, "chart")
            }
            
            # Get text from text frames
            if hasattr(shape, "text_frame"):
                text = shape.text.strip()
                if text:
                    slide_info["texts"].append(text)
                    
                    # Add text details to shape info
                    shape_info["text"] = text
                    shape_info["paragraphs"] = len(shape.text_frame.paragraphs)
            
            # Get table information
            if hasattr(shape, "table"):
                table = shape.table
                shape_info["table"] = {
                    "rows": len(table.rows),
                    "columns": len(table.columns) if len(table.rows) > 0 else 0
                }
            
            slide_info["shapes"].append(shape_info)
        
        presentation_info["slides"].append(slide_info)
    
    return presentation_info

def modify_ppt_text(input_path, output_path, prefix="CHANGE"):
    """
    Modify PowerPoint file by adding a prefix to all text while maintaining styling.
    
    Args:
        input_path (str): Path to the input PowerPoint file
        output_path (str): Path to save the modified PowerPoint file
        prefix (str): Text to prepend to each text element
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        prs = Presentation(input_path)
        
        # Track if any changes were made
        changes_made = False
        
        # Process each slide
        for slide in prs.slides:
            # Process each shape that might contain text
            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    text_frame = shape.text_frame
                    
                    # Process each paragraph in the text frame
                    for paragraph in text_frame.paragraphs:
                        if paragraph.text.strip():  # Only modify non-empty paragraphs
                            # We need to modify each run to preserve styling
                            for run in paragraph.runs:
                                if run.text.strip():  # Skip empty runs
                                    # Preserve the original text with styling
                                    run.text = f"{prefix} {run.text}"
                                    changes_made = True
        
        # Save the modified presentation
        prs.save(output_path)
        return changes_made
    
    except Exception as e:
        print(f"Error modifying PowerPoint: {str(e)}")
        return False