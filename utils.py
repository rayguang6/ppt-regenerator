import os
import re
from datetime import datetime

def get_timestamp():
    """Return a formatted timestamp string."""
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def ensure_dir(directory):
    """Ensure a directory exists, creating it if necessary."""
    if not os.path.exists(directory):
        os.makedirs(directory)
    return directory

def split_into_sections(slides, max_slides_per_section=30):
    """
    Split a large presentation into logical sections.
    
    In a real implementation, this would use more sophisticated heuristics
    to identify logical section boundaries based on slide content and titles.
    """
    sections = []
    current_section = []
    
    for i, slide in enumerate(slides):
        # Add to current section
        current_section.append(slide)
        
        # Check if we should end this section
        # In a real implementation, this would use more sophisticated logic
        # like checking for section title slides, etc.
        if (i + 1) % max_slides_per_section == 0 or i == len(slides) - 1:
            sections.append(current_section)
            current_section = []
    
    return sections

def summarize_section_content(section_content, max_length=500):
    """
    Create a brief summary of section content.
    
    This is a simple implementation. In a real application, this might
    use an LLM call to generate a proper summary.
    """
    # For now, just take the first 500 characters as a "summary"
    full_text = ""
    
    # Handle different content structures
    if isinstance(section_content, list):
        for item in section_content:
            if isinstance(item, dict):
                if "regenerated_texts" in item:
                    full_text += " ".join(item.get("regenerated_texts", []))
                elif "texts" in item:
                    full_text += " ".join(item.get("texts", []))
            elif isinstance(item, str):
                full_text += item
    elif isinstance(section_content, dict):
        if "regenerated_texts" in section_content:
            full_text += " ".join(section_content.get("regenerated_texts", []))
        elif "texts" in section_content:
            full_text += " ".join(section_content.get("texts", []))
    elif isinstance(section_content, str):
        full_text = section_content
    
    if len(full_text) <= max_length:
        return full_text
    
    return full_text[:max_length] + "..."

def extract_key_concepts(content):
    """
    Extract key concepts from content.
    
    In a real implementation, this would use NLP techniques or an LLM
    to identify important terms and concepts.
    """
    # Simple implementation - extract capitalized phrases
    full_text = ""
    
    # Handle different content structures
    if isinstance(content, list):
        for item in content:
            if isinstance(item, dict):
                if "regenerated_texts" in item:
                    full_text += " ".join(item.get("regenerated_texts", []))
                elif "texts" in item:
                    full_text += " ".join(item.get("texts", []))
            elif isinstance(item, str):
                full_text += item
    elif isinstance(content, dict):
        if "regenerated_texts" in content:
            full_text += " ".join(content.get("regenerated_texts", []))
        elif "texts" in content:
            full_text += " ".join(content.get("texts", []))
    elif isinstance(content, str):
        full_text = content
    
    # Find capitalized phrases (crude approximation of key concepts)
    capitalized_phrases = re.findall(r'\b[A-Z][A-Za-z0-9]+ [A-Za-z0-9 ]+\b', full_text)
    
    # Create a dictionary with the phrases and dummy values
    key_concepts = {}
    for phrase in capitalized_phrases:
        if len(phrase) > 5:  # Filter out very short phrases
            key_concepts[phrase] = True
    
    return key_concepts