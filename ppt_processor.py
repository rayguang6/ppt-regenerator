from typing import Dict, List, Any, Optional, Tuple
import os
import tempfile
import time

from utils import (
    split_into_sections, 
    summarize_section_content, 
    extract_key_concepts
)
from ppt_reader import (
    read_ppt, 
    extract_content_with_mapping, 
    modify_ppt_with_mapping
)
from llm_service import LLMService

class PPTProcessor:
    """Process PowerPoint presentations for content regeneration."""
    
    def __init__(self, use_mock=False, api_key=None, debug=False, 
                 max_slides_per_section: int = 50,   # Increased from 30
                 max_total_slides: int = 500):       # Added total slides limit
        """
        Initialize the PPT processor.
        
        Args:
            use_mock: Whether to use mock LLM responses
            api_key: API key for LLM service
            debug: Whether to save debug information
            max_slides_per_section: Maximum slides to process in one section
            max_total_slides: Maximum total slides allowed
        """
        self.debug = debug
        self.max_slides_per_section = max_slides_per_section
        self.max_total_slides = max_total_slides
        
        self.llm_service = LLMService(api_key=api_key, use_mock=use_mock)
        self.debug_info = []  # Store debug info in memory
    
    def process_presentation(self, input_path: str, output_path: str, 
                           max_slides_per_section: int = None,
                           user_info: str = "") -> Dict[str, Any]:
        """
        Process a full presentation, regenerating content while preserving style.
        
        Args:
            input_path: Path to the input PowerPoint file
            output_path: Path to save the modified PowerPoint file
            max_slides_per_section: Override default section size
            user_info: User's industry and use case information
            
        Returns:
            Dictionary with processing results and statistics
        """
        start_time = time.time()
        
        # Validate and set section size
        if max_slides_per_section is None:
            max_slides_per_section = self.max_slides_per_section
        
        # Step 1: Read the presentation and extract content
        presentation_info = read_ppt(input_path, debug=False)
        content_map = extract_content_with_mapping(presentation_info)
        
        # Validate total slides
        total_slides = content_map["slide_count"]
        if total_slides > self.max_total_slides:
            raise ValueError(
                f"Presentation exceeds maximum allowed slides. "
                f"Max: {self.max_total_slides}, Current: {total_slides}"
            )
        
        # Step 2: Split the presentation into manageable sections
        sections = self._split_presentation_into_sections(
            content_map["slides"], 
            max_slides_per_section
        )
        
        # Statistics and tracking
        stats = {
            "total_slides": total_slides,
            "sections": len(sections),
            "section_details": [],
            "key_concepts": {},
            "start_time": start_time,
            "processing_times": [],
            "before_after": [],
            "warnings": content_map.get("warnings", [])
        }
        
        self.debug_info.append({
            "stage": "processing_start",
            "stats": stats,
            "timestamp": time.strftime("%Y-%m-%d %H:%M:%S")
        })
        
        # Step 3: Process each section with context management
        previous_context = ""
        key_concepts = {}
        
        for section_idx, section in enumerate(sections):
            section_start = time.time()
            
            # Create a section record for tracking
            section_record = {
                "section_idx": section_idx,
                "slide_count": len(section),
                "slide_numbers": [slide["slide_number"] for slide in section],
                "start_time": section_start
            }
            
            # Process this section with retry mechanism
            section_info = self._process_section_with_retry(
                section, 
                previous_context, 
                key_concepts,
                user_info,
                section_idx
            )
            
            # Update the content map with regenerated text
            for slide_idx, slide in enumerate(section):
                slide_number = slide["slide_number"]
                # Find the corresponding slide in the content map
                for content_slide in content_map["slides"]:
                    if content_slide["slide_number"] == slide_number:
                        # Map regenerated texts to the content map
                        if "regenerated_texts" in section_info[slide_idx]:
                            content_slide["regenerated_texts"] = section_info[slide_idx]["regenerated_texts"]
                        else:
                            # Fallback if regeneration fails
                            content_slide["regenerated_texts"] = [
                                f"REGENERATION-FAILED: {text}" for text in slide["texts"]
                            ]
                        
                        # Track before/after changes
                        if "before_after" in section_info[slide_idx]:
                            content_slide["before_after"] = section_info[slide_idx]["before_after"]
                            stats["before_after"].append(section_info[slide_idx]["before_after"])
                        break
            
            # Update context and key concepts
            context_update = summarize_section_content(section_info)
            previous_context += f"\nSection {section_idx + 1} Summary: {context_update}"
            
            # Extract and accumulate key concepts
            new_concepts = extract_key_concepts(section_info)
            key_concepts.update(new_concepts)
            
            # Update section record with timing and details
            section_record["end_time"] = time.time()
            section_record["duration"] = section_record["end_time"] - section_record["start_time"]
            section_record["regenerated_texts_count"] = sum(
                len(slide.get("regenerated_texts", [])) for slide in section_info
            )
            
            stats["section_details"].append(section_record)
            stats["processing_times"].append(section_record["duration"])
            
            self.debug_info.append({
                "stage": f"section_{section_idx}_processing",
                "section_record": section_record,
                "timestamp": time.strftime("%Y-%m-%d %H:%M:%S")
            })
        
        # Step 4: Modify the PowerPoint with regenerated content
        modification_start = time.time()
        success = modify_ppt_with_mapping(input_path, output_path, content_map)
        
        # Update final statistics
        stats["end_time"] = time.time()
        stats["total_duration"] = stats["end_time"] - stats["start_time"]
        stats["modification_time"] = stats["end_time"] - modification_start
        stats["success"] = success
        stats["key_concepts"] = key_concepts
        
        self.debug_info.append({
            "stage": "processing_complete",
            "stats": stats,
            "timestamp": time.strftime("%Y-%m-%d %H:%M:%S")
        })
        
        # Also save the LLM service debug info
        stats["llm_debug_info"] = self.llm_service.debug_info
        
        return stats
    
    def _split_presentation_into_sections(
        self, 
        slides: List[Dict], 
        max_slides_per_section: int
    ) -> List[List[Dict]]:
        """
        Split presentation into sections of specified max size.
        
        Args:
            slides: List of all slides
            max_slides_per_section: Maximum slides per section
            
        Returns:
            List of slide sections
        """
        sections = []
        for i in range(0, len(slides), max_slides_per_section):
            section = slides[i:i + max_slides_per_section]
            sections.append(section)
        return sections
    
    def _process_section_with_retry(
        self, 
        section: List[Dict[str, Any]], 
        previous_context: str, 
        key_concepts: Dict[str, Any],
        user_info: str,
        section_idx: int,
        max_retries: int = 3
    ) -> List[Dict[str, Any]]:
        """
        Process a section with retry mechanism and error handling.
        
        Args:
            section: Slides to process
            previous_context: Context from previous sections
            key_concepts: Accumulated key concepts
            user_info: User context information
            section_idx: Index of current section
            max_retries: Maximum retry attempts
            
        Returns:
            Processed section with regenerated content
        """
        for attempt in range(max_retries):
            try:
                # Regenerate content for this section
                regenerated_section = self.llm_service.regenerate_content(
                    section, previous_context, key_concepts, user_info
                )
                
                # Validate regenerated section
                if not regenerated_section or len(regenerated_section) != len(section):
                    raise ValueError(
                        f"Incomplete regeneration in section {section_idx}. "
                        f"Expected {len(section)} slides, got {len(regenerated_section)}"
                    )
                
                # Enrich section with additional metadata
                for i, slide in enumerate(section):
                    # Track before/after text changes
                    before_after = {
                        "slide_number": slide.get("slide_number", i+1),
                        "changes": []
                    }
                    
                    # Get regenerated texts
                    regenerated_texts = regenerated_section[i].get("texts", [])
                    original_texts = slide.get("texts", [])
                    
                    # Track text changes
                    for j, (orig_text, new_text) in enumerate(
                        zip(original_texts, regenerated_texts)
                    ):
                        before_after["changes"].append({
                            "before": orig_text,
                            "after": new_text
                        })
                    
                    # Attach metadata to the slide
                    section[i]["regenerated_texts"] = regenerated_texts
                    section[i]["before_after"] = before_after
                
                return section
            
            except Exception as e:
                # Log the error and prepare for retry
                print(f"Error processing section {section_idx}, attempt {attempt + 1}: {str(e)}")
                
                # Exponential backoff
                time.sleep(2 ** attempt)
        
        # If all retries fail, return original section with failure markers
        fallback_section = [
            {
                **slide, 
                "regenerated_texts": [f"REGENERATION-FAILED: {text}" for text in slide.get("texts", [])],
                "before_after": {
                    "slide_number": slide.get("slide_number"),
                    "changes": [],
                    "error": "Max retries exceeded"
                }
            } 
            for slide in section
        ]
        
        return fallback_section