"""
claude_form_processor.py
Claude AI integration for processing sensory evaluation forms
"""

import anthropic
import base64
import json
import os
from PIL import Image
import io

class ClaudeFormProcessor:
    """
    Processor that uses Claude AI to extract sensory ratings from form images.
    
    This approach leverages Claude's vision capabilities to understand the form
    layout and extract ratings without needing rigid region boundaries or
    extensive training data. It's more flexible and can handle layout variations
    much better than traditional computer vision approaches.
    """
    
    def __init__(self):
        # Try to get API key from environment variable
        self.api_key = os.getenv('ANTHROPIC_API_KEY')
        
        if not self.api_key:
            raise ValueError(
                "Anthropic API key not found. Please set the ANTHROPIC_API_KEY environment variable.\n\n"
                "You can get an API key from: https://console.anthropic.com/\n\n"
                "Set it in your environment:\n"
                "Windows: set ANTHROPIC_API_KEY=your_key_here\n"
                "Mac/Linux: export ANTHROPIC_API_KEY=your_key_here"
            )
        
        self.client = anthropic.Anthropic(api_key=self.api_key)
        
        # Expected metrics from your sensory evaluation form
        self.metrics = [
            "Burnt Taste",
            "Vapor Volume", 
            "Overall Flavor",
            "Smoothness",
            "Overall Liking"
        ]
        
    def prepare_image(self, image_path):
        """
        Prepare image for Claude API by encoding it as base64.
        
        Claude can accept images in base64 format. We also resize large images
        to ensure they fit within API limits while maintaining readability.
        """
        
        try:
            # Open and potentially resize the image
            with Image.open(image_path) as img:
                # Convert to RGB if necessary
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                
                # Resize if too large (Claude has size limits)
                max_size = 1568  # Claude's recommended max dimension
                if max(img.width, img.height) > max_size:
                    ratio = max_size / max(img.width, img.height)
                    new_size = (int(img.width * ratio), int(img.height * ratio))
                    img = img.resize(new_size, Image.Resampling.LANCZOS)
                    print(f"Resized image to {new_size} for API processing")
                
                # Convert to base64
                buffer = io.BytesIO()
                img.save(buffer, format='JPEG', quality=90)
                image_data = base64.b64encode(buffer.getvalue()).decode('utf-8')
                
                return image_data
                
        except Exception as e:
            raise Exception(f"Failed to prepare image: {e}")
    
    def process_form_image(self, image_path):
        """
        Process a sensory evaluation form using Claude AI.
        
        This method sends the image to Claude along with detailed instructions
        about what to look for and how to format the response. Claude's vision
        capabilities allow it to understand the form structure and extract
        ratings accurately.
        """
        
        print(f"Processing form with Claude AI: {os.path.basename(image_path)}")
        
        # Prepare the image
        try:
            image_data = self.prepare_image(image_path)
            print("Image prepared for AI processing")
        except Exception as e:
            raise Exception(f"Image preparation failed: {e}")
        
        # Create the prompt for Claude
        prompt = self.create_extraction_prompt()
        
        try:
            # Send request to Claude
            response = self.client.messages.create(
                model="claude-3-5-sonnet-20241022",  # Latest Claude model with vision
                max_tokens=1500,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "text",
                                "text": prompt
                            },
                            {
                                "type": "image",
                                "source": {
                                    "type": "base64",
                                    "media_type": "image/jpeg",
                                    "data": image_data
                                }
                            }
                        ]
                    }
                ]
            )
            
            # Extract and parse the response
            response_text = response.content[0].text
            print("Received response from Claude AI")
            
            # Parse the JSON response
            extracted_data = self.parse_claude_response(response_text)
            
            print(f"Successfully extracted data for {len(extracted_data)} samples")
            return extracted_data
            
        except anthropic.APIError as e:
            raise Exception(f"Claude API error: {e}")
        except Exception as e:
            raise Exception(f"AI processing failed: {e}")
    
    def create_extraction_prompt(self):
        """
        Create a detailed prompt that instructs Claude how to analyze the form.
        
        The key to getting good results from Claude is providing clear, specific
        instructions about what to look for and how to format the response.
        """
        
        prompt = f"""
Please analyze this sensory evaluation form image and extract the ratings data. 

FORM STRUCTURE:
- The form contains 4 samples arranged in a 2x2 grid
- Each sample has ratings for 5 attributes: {', '.join(self.metrics)}
- Ratings are on a 1-9 scale where participants circle their chosen number
- Look for circled numbers, checkmarks, or other clear markings indicating the selected rating

EXTRACTION INSTRUCTIONS:
1. Identify each sample section (should be 4 total: Sample 1, Sample 2, Sample 3, Sample 4)
2. For each sample, find the rating for each of the 5 attributes
3. Look for circled numbers, heavy marks, checkmarks, or other clear indicators
4. If a rating is unclear or unmarked, use null

RESPONSE FORMAT:
Return the data as a JSON object with this exact structure:

{{
    "Sample 1": {{
        "Burnt Taste": 5,
        "Vapor Volume": 7,
        "Overall Flavor": 6,
        "Smoothness": 8,
        "Overall Liking": 7,
        "comments": ""
    }},
    "Sample 2": {{
        "Burnt Taste": 4,
        "Vapor Volume": 6,
        "Overall Flavor": 5,
        "Smoothness": 7,
        "Overall Liking": 6,
        "comments": ""
    }},
    "Sample 3": {{
        "Burnt Taste": 6,
        "Vapor Volume": 8,
        "Overall Flavor": 7,
        "Smoothness": 6,
        "Overall Liking": 8,
        "comments": ""
    }},
    "Sample 4": {{
        "Burnt Taste": 3,
        "Vapor Volume": 5,
        "Overall Flavor": 4,
        "Smoothness": 5,
        "Overall Liking": 4,
        "comments": ""
    }}
}}

IMPORTANT:
- Use only numbers 1-9 for ratings, or null if unclear
- Include all 5 attributes for each sample
- Use the exact attribute names I provided
- Sample names should be "Sample 1", "Sample 2", "Sample 3", "Sample 4"
- Return only the JSON object, no additional text
- If you see handwritten comments, include them in the comments field
"""
        
        return prompt
    
    def parse_claude_response(self, response_text):
        """
        Parse Claude's response and convert it to the format expected by the GUI.
        
        This method handles the JSON parsing and validates that the response
        contains the expected structure and data types.
        """
        
        try:
            # Find JSON in the response (Claude sometimes includes explanation text)
            response_text = response_text.strip()
            
            # Look for JSON object boundaries
            start_idx = response_text.find('{')
            end_idx = response_text.rfind('}') + 1
            
            if start_idx == -1 or end_idx == 0:
                raise ValueError("No JSON object found in response")
            
            json_text = response_text[start_idx:end_idx]
            
            # Parse the JSON
            data = json.loads(json_text)
            
            # Validate the structure
            validated_data = {}
            
            for sample_name, sample_data in data.items():
                if not isinstance(sample_data, dict):
                    continue
                    
                validated_sample = {}
                
                # Validate each metric
                for metric in self.metrics:
                    rating = sample_data.get(metric)
                    
                    # Convert and validate rating
                    if rating is None or rating == "null":
                        validated_sample[metric] = 5  # Default to middle rating
                    elif isinstance(rating, (int, float)):
                        # Ensure rating is in valid range
                        rating = max(1, min(9, int(rating)))
                        validated_sample[metric] = rating
                    else:
                        print(f"Warning: Invalid rating '{rating}' for {sample_name} - {metric}, using default")
                        validated_sample[metric] = 5
                
                # Include comments
                validated_sample['comments'] = str(sample_data.get('comments', ''))
                
                validated_data[sample_name] = validated_sample
            
            if not validated_data:
                raise ValueError("No valid sample data found in response")
            
            return validated_data
            
        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid JSON in Claude response: {e}\n\nResponse: {response_text[:200]}...")
        except Exception as e:
            raise ValueError(f"Failed to parse Claude response: {e}")