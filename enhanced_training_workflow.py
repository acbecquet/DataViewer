# enhanced_training_workflow.py
"""
Complete workflow for Claude-enhanced training data generation
"""

import os
import sys
from claude_enhanced_training_assistant import ClaudeEnhancedTrainingAssistant
from enhanced_ml_processor_updater import MLProcessorUpdater

def run_enhanced_training_workflow():
    """
    Complete workflow that combines Claude's vision with my ML training pipeline.
    """
    
    print("="*60)
    print("CLAUDE-ENHANCED TRAINING WORKFLOW")
    print("="*60)
    print("This workflow will:")
    print("1. Use Claude to detect precise sample boundaries")
    print("2. Extract training data with improved accuracy") 
    print("3. Generate configuration for your ML processor")
    print("4. Optionally update your ML processor code")
    print()
    
    # Get training images folder
    training_folder = input("Enter path to folder with training form images: ").strip()
    
    if not os.path.exists(training_folder):
        print(f"Error: Folder {training_folder} not found")
        return
    
    try:
        # Step 1: Run Claude-enhanced training extraction
        print("\nStep 1: Running Claude-enhanced boundary detection and training extraction...")
        assistant = ClaudeEnhancedTrainingAssistant()
        assistant.process_training_images_with_claude(training_folder)
        
        print("\nStep 2: Training data extraction complete!")
        
        # Step 3: Ask if user wants to update ML processor
        update_choice = input("\nDo you want to update your ML processor with the improved boundaries? (y/n): ").strip().lower()
        
        if update_choice == 'y':
            # Find the most recent boundaries file
            claude_analysis_dir = "training_data/claude_analysis"
            if os.path.exists(claude_analysis_dir):
                boundary_files = [f for f in os.listdir(claude_analysis_dir) if f.startswith('improved_boundaries')]
                if boundary_files:
                    latest_file = max(boundary_files)
                    full_path = os.path.join(claude_analysis_dir, latest_file)
                    
                    print(f"\nStep 3: Updating ML processor with {latest_file}...")
                    updater = MLProcessorUpdater()
                    updater.update_ml_processor_boundaries(full_path)
                    
                    print("\nWorkflow complete! Your ML processor has been enhanced with Claude's intelligence.")
                    print("Next steps:")
                    print("1. Train your ML model using the new training data")
                    print("2. Test the improved accuracy on new forms")
                    print("3. Compare performance with the original hardcoded boundaries")
                else:
                    print("No boundary files found to update with")
            else:
                print("Claude analysis directory not found")
        else:
            print("\nWorkflow complete! Training data extracted with Claude-enhanced boundaries.")
            print("You can manually update your ML processor later if desired.")
            
    except Exception as e:
        print(f"Error in workflow: {e}")
        print("Check that:")
        print("1. ANTHROPIC_API_KEY environment variable is set")
        print("2. Training images are in the specified folder")
        print("3. You have write permissions in the current directory")

if __name__ == "__main__":
    run_enhanced_training_workflow()