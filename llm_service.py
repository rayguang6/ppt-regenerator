import json
import os
import time
import requests
import re
from typing import Dict, List, Union, Optional
import random

class LLMService:
    """Service to handle interactions with LLM APIs."""
    
    def __init__(self, api_key=None):
        """
        Initialize the LLM service.
        
        Args:
            api_key: DeepSeek API key (or None to use environment variable)
        """
        self.api_key = api_key or os.environ.get("DEEPSEEK_API_KEY")
    
    def regenerate_content(self, 
                           content: List[Dict], 
                           previous_context: str = "", 
                           key_concepts: Dict = None,
                           user_info: str = "") -> List[Dict]:
        """
        Regenerate content using the LLM.
        
        Args:
            content: List of slides with text to regenerate
            previous_context: Summary of previous content for context
            key_concepts: Dictionary of key concepts to maintain
            user_info: User industry and use case information
        
        Returns:
            List of dictionaries with regenerated text
        """
        return self._api_regenerate_content(content, previous_context, key_concepts, user_info)
    
    def _build_prompt(self, content, previous_context, key_concepts, user_info):
        """Build the prompt for the LLM."""
        # Format the slides content
        slides_text = ""
        for i, slide in enumerate(content):
            slides_text += f"SLIDE {slide.get('slide_number', i+1)}:\n"
            
            if slide.get('texts'):
                for j, text_block in enumerate(slide['texts']):
                    slides_text += f"TEXT {j+1}: {text_block}\n"
            
            slides_text += "\n"
        
        # Format key concepts if available
        concepts_text = ""
        if key_concepts and len(key_concepts) > 0:
            concepts_text = "KEY CONCEPTS TO MAINTAIN CONSISTENCY WITH:\n"
            for concept in key_concepts:
                concepts_text += f"- {concept}\n"
            concepts_text += "\n"
        
        # Add user information if available
        user_context = ""
        if user_info and user_info.strip():
            user_context = f"""
USER CONTEXT:
The content should be tailored to the following industry/use case:
{user_info}

"""
        
        # Build the full prompt
        prompt = f"""You are a professional presentation content regenerator. Your task is to completely rewrite the content for a PowerPoint presentation while maintaining the original structure, purpose, and approximate length.
IMPORTANT: This is a PROVEN SALES FRAMEWORK for selling courses. Maintain all persuasive elements, psychological triggers, and call-to-action structures while changing only the specific topic and examples.

{user_context}
{concepts_text}
PREVIOUS CONTENT SUMMARY:
{previous_context}

CURRENT SLIDES TO REGENERATE:
{slides_text}

INSTRUCTIONS:
1. ADAPT each slide's text to the new course topic while PRESERVING the original persuasive structure
2. Keep the same number of text blocks per slide
3. Maintain approximately the same length for each text block
4. Never cut sentences in half - always complete thoughts
5. Preserve all sales psychology elements like:
   - Attention-grabbing hooks and questions
   - Credibility statements
   - Pain points and objection handling
   - Call-to-action language
   - Social proof references
   - Urgency and scarcity elements
6. Return your response as a properly formatted JSON array

FORMAT OF RESPONSE:
[
  {{
    "slide_number": 1,
    "texts": [
      "Regenerated text for the first text block",
      "Regenerated text for the second text block"
    ]
  }},
  {{
    "slide_number": 2,
    "texts": [
      "Regenerated text for this slide's text block"
    ]
  }}
]

Remember: Your goal is NOT to create entirely new content, but to ADAPT the proven sales framework to the new course topic. The structure and persuasive elements are what make this framework effective.
"""
        return prompt
    
    def _calculate_prompt_tokens(self, prompt: str) -> int:
        """
        More accurate token estimation using a combination of methods.
        
        Args:
            prompt: The input prompt
        
        Returns:
            Estimated number of tokens
        """
        # Rough token estimation with more nuanced approach
        # Approximation based on common token counting methods
        words = prompt.split()
        
        # Base estimation: words + special tokens
        base_tokens = len(words)
        
        # Add extra tokens for punctuation and special characters
        punctuation_tokens = len(re.findall(r'[.,!?;:]', prompt)) // 2
        
        # Add tokens for code-like structures or JSON-like content
        code_tokens = len(re.findall(r'[{}[\]:"]', prompt)) // 3
        
        # Estimate tokens for numbers and special symbols
        special_tokens = len(re.findall(r'\d+|[#@%&*()]', prompt)) // 3
        
        # Combine estimations with a slight buffer
        total_tokens = base_tokens + punctuation_tokens + code_tokens + special_tokens
        
        # Ensure a minimum token count and add a small buffer
        return max(10, int(total_tokens * 1.2))
    
    def _split_content_by_tokens(self, 
                                 content: List[Dict], 
                                 max_tokens: int = 3500) -> List[List[Dict]]:
        """
        Split content into batches based on more accurate token estimation.
        
        Args:
            content: List of slides to process
            max_tokens: Maximum tokens per batch
        
        Returns:
            List of content batches
        """
        batches = []
        current_batch = []
        current_tokens = 0
        
        for slide in content:
            # Estimate tokens for this slide with more context
            slide_text = "\n".join(slide.get('texts', []))
            slide_tokens = self._calculate_prompt_tokens(slide_text)
            
            # Dynamically adjust batch size based on slide complexity
            dynamic_max_tokens = max_tokens - 500  # Reserve tokens for context and instructions
            
            # If adding this slide would exceed max tokens, start a new batch
            if current_tokens + slide_tokens > dynamic_max_tokens:
                if current_batch:
                    batches.append(current_batch)
                current_batch = []
                current_tokens = 0
            
            current_batch.append(slide)
            current_tokens += slide_tokens
        
        # Add the last batch if not empty
        if current_batch:
            batches.append(current_batch)
        
        # Log batching information for debugging
        print(f"Split {len(content)} slides into {len(batches)} batches")
        for i, batch in enumerate(batches):
            print(f"Batch {i+1}: {len(batch)} slides")
        
        return batches
    
    def _api_regenerate_content(self, content, previous_context, key_concepts, user_info=""):
        """
        Enhanced API call with comprehensive error handling and logging.
        
        Args:
            content: List of slides to regenerate
            previous_context: Context from previous processing
            key_concepts: Key concepts to maintain
            user_info: User-specific context
        
        Returns:
            List of regenerated slides
        """
        if not self.api_key:
            raise ValueError("DeepSeek API key not provided. Please set DEEPSEEK_API_KEY in your .env file.")
        
        # Comprehensive logging setup
        error_log = []
        
        try:
            # Split content into batches with improved token estimation
            content_batches = self._split_content_by_tokens(content)
            
            regenerated_content = []
            
            for batch_idx, batch in enumerate(content_batches):
                batch_error = None
                
                try:
                    # Build prompt for this batch
                    prompt = self._build_prompt(batch, previous_context, key_concepts, user_info)
                    
                    headers = {
                        "Authorization": f"Bearer {self.api_key}",
                        "Content-Type": "application/json"
                    }
                    
                    payload = {
                        "model": "deepseek-chat",
                        "messages": [
                            {"role": "system", "content": "You are a professional presentation content regenerator."},
                            {"role": "user", "content": prompt}
                        ],
                        "temperature": 0.7,
                        "max_tokens": 4000
                    }
                    
                    # Detailed debug logging
                    print(f"\n==== API CALL DETAILS (Batch {batch_idx + 1}/{len(content_batches)} ====")
                    print(f"Batch slides: {len(batch)}")
                    print(f"Prompt length: {len(prompt)} characters")
                    print(f"Estimated tokens: {self._calculate_prompt_tokens(prompt)}")
                    
                    # Enhanced retry mechanism with comprehensive error tracking
                    max_retries = 3
                    for attempt in range(max_retries):
                        try:
                            response = requests.post(
                                "https://api.deepseek.com/v1/chat/completions", 
                                headers=headers, 
                                json=payload,
                                timeout=180  # Increased timeout to 3 minutes
                            )
                            
                            # Comprehensive response handling
                            if response.status_code == 200:
                                result = response.json()
                                assistant_message = result.get("choices", [{}])[0].get("message", {}).get("content", "")
                                
                                # Advanced JSON parsing with multiple fallback strategies
                                try:
                                    # Primary parsing method
                                    batch_regenerated = json.loads(assistant_message)
                                    
                                    # Validate regenerated content structure
                                    if not isinstance(batch_regenerated, list):
                                        raise ValueError("Regenerated content is not a list")
                                    
                                    regenerated_content.extend(batch_regenerated)
                                    break  # Successful parsing, exit retry loop
                                
                                except (json.JSONDecodeError, ValueError) as parse_error:
                                    # Fallback parsing strategies
                                    print(f"JSON parsing failed on attempt {attempt + 1}: {parse_error}")
                                    
                                    # Try extracting JSON from markdown code block
                                    json_match = re.search(r'```json(.*?)```', assistant_message, re.DOTALL)
                                    if json_match:
                                        try:
                                            batch_regenerated = json.loads(json_match.group(1).strip())
                                            regenerated_content.extend(batch_regenerated)
                                            break
                                        except Exception as e:
                                            print(f"Markdown JSON extraction failed: {e}")
                                    
                                    # Last resort: manual parsing
                                    if attempt == max_retries - 1:
                                        raise parse_error
                            
                            else:
                                # Detailed error logging for API errors
                                error_details = {
                                    "status_code": response.status_code,
                                    "response_text": response.text,
                                    "batch_index": batch_idx
                                }
                                error_log.append(error_details)
                                print(f"API error details: {error_details}")
                                raise requests.RequestException(f"API returned status {response.status_code}")
                        
                        except (requests.Timeout, requests.ConnectionError) as connection_error:
                            # Detailed connection error logging
                            connection_error_details = {
                                "error_type": type(connection_error).__name__,
                                "error_message": str(connection_error),
                                "attempt": attempt + 1,
                                "batch_index": batch_idx
                            }
                            error_log.append(connection_error_details)
                            print(f"Connection error details: {connection_error_details}")
                            
                            if attempt == max_retries - 1:
                                raise
                            
                            # Exponential backoff with jitter
                            time.sleep((2 ** attempt) + random.random())
                
                except Exception as batch_error:
                    # Comprehensive batch processing error handling
                    error_details = {
                        "error_type": type(batch_error).__name__,
                        "error_message": str(batch_error),
                        "batch_index": batch_idx
                    }
                    error_log.append(error_details)
                    print(f"Batch processing error: {error_details}")
                    
                    # Add placeholder regeneration for failed batch
                    batch_placeholder = [
                        {
                            **slide, 
                            "texts": [f"REGENERATION-FAILED: {text}" for text in slide.get('texts', [])]
                        } 
                        for slide in batch
                    ]
                    regenerated_content.extend(batch_placeholder)
            
            # Final validation of regenerated content
            if not regenerated_content:
                raise ValueError("No content was regenerated")
            
            return regenerated_content
        
        except Exception as final_error:
            # Ultimate fallback for any unhandled errors
            final_error_details = {
                "error_type": type(final_error).__name__,
                "error_message": str(final_error),
                "error_log": error_log
            }
            print(f"Final processing error: {final_error_details}")
            
            # Return placeholder content for all slides
            placeholder_content = [
                {
                    **slide, 
                    "texts": [f"FINAL-ERROR: {text}" for text in slide.get('texts', [])]
                } 
                for slide in content
            ]
            
            return placeholder_content