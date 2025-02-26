import json
import os
from pathlib import Path
import sys

def extract_email_from_filename(filename):
    """Extract email address from screenshot filename."""
    # Assuming format like "email_at_domain.com_timestamp.png"
    parts = filename.split('_')
    if len(parts) >= 3:
        email_parts = parts[:-2]  # Skip timestamp and extension
        email = '_'.join(email_parts).replace('_at_', '@')
        return email
    return None

def save_readable_results():
    """Save match results in a more readable format"""
    # Load match_results.json
    with open('match_results.json', 'r') as f:
        results = json.load(f)
    
    readable_results = {}
    
    # First, get all unique emails from screenshots directory
    all_emails = set()
    screenshots_dir = os.path.join(os.getcwd(), 'screenshots')
    if os.path.exists(screenshots_dir):
        for filename in os.listdir(screenshots_dir):
            if filename.endswith('.png') and '@' in filename:
                email = extract_email_from_filename(filename)
                if email:
                    all_emails.add(email)
    
    # Initialize all emails with zero scores
    for email in all_emails:
        readable_results[email] = {
            'matching_memorias': [],
            'score': 0
        }
    
    # Update with actual match results
    for email, match_results in results.items():
        if email not in readable_results:
            readable_results[email] = {
                'matching_memorias': [],
                'score': 0
            }
        
        # Collect all memoria matches for this email
        all_matches = []
        for match_result in match_results:
            for match in match_result.get('matches', []):
                all_matches.append({
                    'memoria': match['memoria_name'],
                    'custom_score': match['custom_score'],
                    'match_quality_score': match['match_quality_score']
                })
        
        # Get unique memorias with their best scores
        memoria_scores = {}
        for match in all_matches:
            memoria = match['memoria']
            custom_score = match['custom_score']
            match_quality = match['match_quality_score']
            
            if memoria not in memoria_scores or match_quality > memoria_scores[memoria]['match_quality_score']:
                memoria_scores[memoria] = {
                    'custom_score': custom_score,
                    'match_quality_score': match_quality
                }
        
        # Calculate overall score (sum of custom scores of top 3 memorias)
        top_memorias = sorted(memoria_scores.items(), key=lambda x: (x[1]['custom_score'], x[1]['match_quality_score']), reverse=True)[:3]
        if top_memorias:
            total_score = sum(info['custom_score'] for _, info in top_memorias)
            readable_results[email]['score'] = total_score
            readable_results[email]['matching_memorias'] = [
                {'name': name, 'score': info['custom_score']} for name, info in top_memorias
            ]
    
    # Save to file
    with open('memoria_match_results.json', 'w') as f:
        json.dump(readable_results, f, indent=2)
    
    print("Saved readable results to memoria_match_results.json")

if __name__ == "__main__":
    save_readable_results()
