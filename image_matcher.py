import os
import cv2
import numpy as np
from pathlib import Path
import json
from datetime import datetime

class ImageMatcher:
    """
    A class for matching partial images (memorias) against screenshots and scoring the matches.
    """
    
    def __init__(self, memoria_dir='memorias', screenshots_dir='screenshots', threshold=0.7, custom_scores=None):
        """
        Initialize the ImageMatcher.
        
        Args:
            memoria_dir: Directory containing the memoria template images
            screenshots_dir: Directory containing the screenshots to match against
            threshold: Minimum confidence threshold for a match (0.0 to 1.0)
            custom_scores: Dictionary mapping memoria names to custom point values
                           e.g. {'yingying-ss1': 20, 'another-memoria': 15}
        """
        self.memoria_dir = Path(memoria_dir)
        self.screenshots_dir = Path(screenshots_dir)
        self.threshold = threshold
        self.custom_scores = custom_scores or {}
        self.memorias = self._load_memorias()
        
    def _load_memorias(self):
        """Load all memoria template images from the memoria directory."""
        memorias = {}
        
        if not self.memoria_dir.exists():
            print(f"Warning: Memoria directory {self.memoria_dir} does not exist")
            return memorias
            
        for img_path in self.memoria_dir.glob('*.png'):
            try:
                img = cv2.imread(str(img_path))
                if img is not None:
                    memorias[img_path.stem] = {
                        'image': img,
                        'path': str(img_path),
                        'custom_score': self.custom_scores.get(img_path.stem, 0)  # Default to 0 if not specified
                    }
            except Exception as e:
                print(f"Error loading memoria {img_path}: {e}")
                
        return memorias
        
    def match_screenshot(self, screenshot_path, scoring_criteria=None):
        """
        Match memorias against a single screenshot.
        
        Args:
            screenshot_path: Path to the screenshot image
            scoring_criteria: Dictionary of criteria for scoring matches
                              (e.g. {'position_weight': 0.3, 'size_weight': 0.2, 'color_weight': 0.5})
                              
        Returns:
            List of dictionaries containing match information
        """
        # Default scoring criteria if none provided
        if scoring_criteria is None:
            scoring_criteria = {
                'match_confidence_weight': 0.7,
                'position_weight': 0.1,
                'size_weight': 0.1,
                'color_similarity_weight': 0.1
            }
            
        # Extract email from screenshot filename
        screenshot_name = Path(screenshot_path).name
        email = self._extract_email_from_filename(screenshot_name)
        
        # Load screenshot
        screenshot = cv2.imread(str(screenshot_path))
        if screenshot is None:
            print(f"Error: Could not load screenshot {screenshot_path}")
            return []
            
        matches = []
        
        # Process each memoria template
        for memoria_name, memoria_data in self.memorias.items():
            memoria_img = memoria_data['image']
            
            # Use template matching to find the memoria in the screenshot
            result = cv2.matchTemplate(screenshot, memoria_img, cv2.TM_CCOEFF_NORMED)
            
            # Get the best match location and confidence
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
            
            # If match confidence exceeds threshold, calculate score
            if max_val >= self.threshold:
                # Get the position of the match
                top_left = max_loc
                h, w = memoria_img.shape[:2]
                bottom_right = (top_left[0] + w, top_left[1] + h)
                
                # Calculate position score (center of screen is better)
                center_x, center_y = screenshot.shape[1] / 2, screenshot.shape[0] / 2
                match_center_x = top_left[0] + w / 2
                match_center_y = top_left[1] + h / 2
                
                # Normalized distance from center (0 = center, 1 = corner)
                max_distance = np.sqrt((center_x**2) + (center_y**2))
                distance = np.sqrt((center_x - match_center_x)**2 + (center_y - match_center_y)**2)
                position_score = 1 - (distance / max_distance)
                
                # Calculate size score (larger is better, up to a point)
                size_ratio = (w * h) / (screenshot.shape[1] * screenshot.shape[0])
                size_score = min(size_ratio * 10, 1.0)  # Cap at 1.0
                
                # Calculate color similarity
                match_region = screenshot[top_left[1]:bottom_right[1], top_left[0]:bottom_right[0]]
                color_similarity = self._calculate_color_similarity(match_region, memoria_img)
                
                # Calculate final score based on weights
                match_quality_score = (
                    max_val * scoring_criteria.get('match_confidence_weight', 0.7) +
                    position_score * scoring_criteria.get('position_weight', 0.1) +
                    size_score * scoring_criteria.get('size_weight', 0.1) +
                    color_similarity * scoring_criteria.get('color_similarity_weight', 0.1)
                )
                
                # Create match record
                match = {
                    'memoria_name': memoria_name,
                    'memoria_path': memoria_data['path'],
                    'confidence': max_val,
                    'position': top_left,
                    'size': (w, h),
                    'position_score': position_score,
                    'size_score': size_score,
                    'color_similarity': color_similarity,
                    'match_quality_score': match_quality_score,
                    'custom_score': memoria_data['custom_score'],  # Use the custom score from memoria data
                    'final_score': memoria_data['custom_score']  # Add final_score for backward compatibility
                }
                
                matches.append(match)
        
        # Sort matches by custom_score (highest first), then by match_quality_score as a tiebreaker
        matches.sort(key=lambda x: (x['custom_score'], x['match_quality_score']), reverse=True)
        
        return {
            'email': email,
            'screenshot_path': str(screenshot_path),
            'matches': matches,
            'timestamp': datetime.now().isoformat()
        }
    
    def _calculate_color_similarity(self, region1, region2):
        """Calculate color similarity between two image regions."""
        # Resize region2 to match region1 if needed
        if region1.shape != region2.shape:
            region2 = cv2.resize(region2, (region1.shape[1], region1.shape[0]))
            
        # Calculate histogram for both regions
        hist1 = cv2.calcHist([region1], [0, 1, 2], None, [8, 8, 8], [0, 256, 0, 256, 0, 256])
        hist2 = cv2.calcHist([region2], [0, 1, 2], None, [8, 8, 8], [0, 256, 0, 256, 0, 256])
        
        # Normalize histograms
        cv2.normalize(hist1, hist1, 0, 1, cv2.NORM_MINMAX)
        cv2.normalize(hist2, hist2, 0, 1, cv2.NORM_MINMAX)
        
        # Compare histograms
        similarity = cv2.compareHist(hist1, hist2, cv2.HISTCMP_CORREL)
        
        # Return similarity (0 to 1, where 1 is identical)
        return max(0, similarity)
    
    def _extract_email_from_filename(self, filename):
        """Extract email address from screenshot filename."""
        # Assuming format like "email_at_domain.com_timestamp.png"
        parts = filename.split('_')
        if len(parts) >= 3:
            email_parts = parts[:-2]  # Skip timestamp and extension
            email = '_'.join(email_parts).replace('_at_', '@')
            return email
        return None
    
    def batch_match_screenshots(self, scoring_criteria=None, email_filter=None, skip_processed=True):
        """
        Match memorias against all screenshots in the screenshots directory.
        
        Args:
            scoring_criteria: Dictionary of criteria for scoring matches
            email_filter: Optional filter to only process screenshots with matching email
            skip_processed: If True, skip screenshots that have already been processed
                           and have results in match_results.json
            
        Returns:
            Dictionary with emails as keys and lists of matches as values
        """
        results = {}
        
        if not self.screenshots_dir.exists():
            print(f"Error: Screenshots directory {self.screenshots_dir} does not exist")
            return results
        
        # Load existing results if skip_processed is True
        processed_screenshots = set()
        if skip_processed and os.path.exists('match_results.json'):
            try:
                with open('match_results.json', 'r') as f:
                    existing_results = json.load(f)
                    
                # Extract paths of already processed screenshots
                for email_results in existing_results.values():
                    for match_result in email_results:
                        processed_screenshots.add(match_result['screenshot_path'])
                        
                print(f"Found {len(processed_screenshots)} already processed screenshots")
            except (json.JSONDecodeError, KeyError) as e:
                print(f"Error loading existing results: {e}")
                processed_screenshots = set()
            
        # Process each screenshot
        for screenshot_path in self.screenshots_dir.glob('*.png'):
            # Skip if doesn't match email filter
            if email_filter and email_filter not in screenshot_path.name:
                continue
                
            # Skip if already processed
            if skip_processed and str(screenshot_path) in processed_screenshots:
                print(f"Skipping already processed screenshot: {screenshot_path.name}")
                continue
                
            match_result = self.match_screenshot(screenshot_path, scoring_criteria)
            
            if match_result['email']:
                email = match_result['email']
                
                if email not in results:
                    results[email] = []
                    
                results[email].append(match_result)
        
        return results
    
    def save_results(self, results, output_file='match_results.json'):
        """Save match results to a JSON file."""
        # Load existing results if file exists
        existing_results = {}
        if os.path.exists(output_file):
            try:
                with open(output_file, 'r') as f:
                    existing_results = json.load(f)
            except json.JSONDecodeError:
                existing_results = {}
        
        # Merge new results with existing ones
        merged_results = existing_results.copy()
        for email, match_results in results.items():
            if email in merged_results:
                # Add only new screenshots
                existing_paths = {r['screenshot_path'] for r in merged_results[email]}
                for match_result in match_results:
                    if match_result['screenshot_path'] not in existing_paths:
                        merged_results[email].append(match_result)
            else:
                merged_results[email] = match_results
        
        # Save merged results
        with open(output_file, 'w') as f:
            json.dump(merged_results, f, indent=2)
        
        print(f"Results saved to {output_file}")

    def load_results(self, input_file='match_results.json'):
        """Load match results from a JSON file."""
        if not os.path.exists(input_file):
            print(f"Error: Results file {input_file} does not exist")
            return {}
            
        with open(input_file, 'r') as f:
            return json.load(f)


def match_memorias(custom_scores=None, scoring_criteria=None, email_filter=None, threshold=0.7, skip_processed=True):
    """
    Convenience function to match memorias against screenshots.
    
    Args:
        custom_scores: Dictionary mapping memoria names to custom point values
                       e.g. {'yingying-ss1': 20, 'another-memoria': 15}
        scoring_criteria: Dictionary of criteria for scoring matches
        email_filter: Optional filter to only process screenshots with matching email
        threshold: Minimum confidence threshold for a match (0.0 to 1.0)
        skip_processed: If True, skip screenshots that have already been processed
        
    Returns:
        Dictionary with emails as keys and lists of matches as values
    """
    matcher = ImageMatcher(threshold=threshold, custom_scores=custom_scores)
    results = matcher.batch_match_screenshots(scoring_criteria, email_filter, skip_processed)
    matcher.save_results(results)
    return results


if __name__ == "__main__":
    # Example usage with custom scores
    custom_scores = {
        'yingying-ss1': 20,
        'aoi-ss2': 5,
        'yuina-ss3': 1,
        'seira-ss1': 1,
        'chie-ss2': 1,
        'seika-ss2': 20,
        'seika-ss1': 3,
        'tama-ss1': 1,
        'tama-ss4': 2
    }
    
    scoring_criteria = {
        'match_confidence_weight': 0.6,
        'position_weight': 0.2,
        'size_weight': 0.1,
        'color_similarity_weight': 0.1
    }
    
    results = match_memorias(custom_scores, scoring_criteria)
    
    # Print summary of results
    for email, matches in results.items():
        print(f"\nEmail: {email}")
        for match_result in matches:
            print(f"  Screenshot: {os.path.basename(match_result['screenshot_path'])}")
            print(f"  Matches found: {len(match_result['matches'])}")
            
            for i, match in enumerate(match_result['matches'][:3]):  # Show top 3 matches
                print(f"    Match {i+1}: {match['memoria_name']} " +
                      f"(Custom Score: {match['custom_score']}, Match Quality: {match['match_quality_score']:.2f})")