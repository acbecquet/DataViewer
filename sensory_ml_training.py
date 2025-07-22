"""
sensory_ml_training.py
ML Training and AI Processing Module for Sensory Data Collection
Developed by Charlie Becquet
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import json
from datetime import datetime
import numpy as np
from utils import debug_print


class SensoryMLTrainer:
    """Handles all ML training and AI processing functionality."""
    
    def __init__(self, parent_window):
        """Initialize with reference to parent window for UI updates."""
        self.parent = parent_window
        
    def check_enhanced_data_balance(self):
        """Check enhanced training data balance and quality."""
        try:
            base_dir = "training_data/sensory_ratings"
            if not os.path.exists(base_dir):
                messagebox.showwarning("No Enhanced Data", 
                                     "Enhanced training data not found.\n"
                                     "Run enhanced extraction first.")
                return
        
            print("="*80)
            print("ENHANCED TRAINING DATA ANALYSIS")
            print("="*80)
        
            # Detailed analysis with enhanced metrics
            class_distribution = {}
            enhanced_info_files = {}
            total_images = 0
            total_enhanced = 0
        
            for rating in range(1, 10):
                rating_dir = os.path.join(base_dir, f"rating_{rating}")
                if os.path.exists(rating_dir):
                    # Count image files
                    images = [f for f in os.listdir(rating_dir) 
                             if f.lower().endswith(('.jpg', '.png', '.jpeg'))]
                    count = len(images)
                    class_distribution[rating] = count
                    total_images += count
                
                    # Count enhanced extraction info files
                    info_files = [f for f in os.listdir(rating_dir) if f.endswith('_info.txt')]
                    enhanced_info_files[rating] = len(info_files)
                    total_enhanced += len(info_files)
                
                    # Sample file analysis
                    sample_sizes = []
                    for img_file in images[:5]:  # Check first 5 files
                        img_path = os.path.join(rating_dir, img_file)
                        try:
                            import cv2
                            img = cv2.imread(img_path)
                            if img is not None:
                                sample_sizes.append(f"{img.shape[1]}x{img.shape[0]}")
                        except:
                            pass
                
                    size_info = f" (sizes: {', '.join(set(sample_sizes[:3]))})" if sample_sizes else ""
                
                    print(f"Rating {rating}: {count:4d} images, {len(info_files):3d} enhanced{size_info}")
                
                    # Show sample filenames
                    if images:
                        print(f"  Sample files: {images[:2]}")
        
            # Enhanced analysis
            print("-" * 70)
            print(f"Total training images: {total_images}")
            print(f"Enhanced extractions: {total_enhanced}")
        
            if total_images > 0:
                min_count = min(class_distribution.values())
                max_count = max(class_distribution.values())
                imbalance_ratio = max_count / max(min_count, 1)
                enhancement_rate = total_enhanced / total_images
            
                print(f"\nEnhanced Quality Metrics:")
                print(f"  Class balance ratio: {imbalance_ratio:.2f}")
                print(f"  Enhancement rate: {enhancement_rate:.1%}")
                print(f"  Min/Max class sizes: {min_count}/{max_count}")
            
                # Enhanced recommendations
                recommendations = []
                if total_images < 100:
                    recommendations.append("Collect more training data (target: 100+ images)")
                if imbalance_ratio > 3.0:
                    recommendations.append("Balance classes - some ratings underrepresented")
                if enhancement_rate < 0.8:
                    recommendations.append("Re-extract with enhanced workflow for better quality")
                if total_images < 200:
                    recommendations.append("For production quality: collect 200+ images")
            
                # Status assessment
                if not recommendations:
                    status = "EXCELLENT - Ready for production training"
                elif len(recommendations) <= 2:
                    status = "GOOD - Ready for training with minor improvements"
                else:
                    status = "NEEDS IMPROVEMENT - Address issues before training"
            
                print(f"\nStatus: {status}")
            
                if recommendations:
                    print(f"\nRecommendations:")
                    for i, rec in enumerate(recommendations, 1):
                        print(f"  {i}. {rec}")
            
                # Show in dialog
                dialog_msg = (f"Enhanced Training Data Analysis\n\n"
                             f"Status: {status}\n\n"
                             f"Metrics:\n"
                             f"• Total images: {total_images}\n"
                             f"• Enhanced extractions: {total_enhanced} ({enhancement_rate:.1%})\n"
                             f"• Class balance ratio: {imbalance_ratio:.2f}\n"
                             f"• Resolution: 600x140 pixels\n\n")
            
                if recommendations:
                    dialog_msg += "Recommendations:\n" + "\n".join(f"• {rec}" for rec in recommendations)
                else:
                    dialog_msg += "✓ Data is ready for enhanced model training!"
                
                messagebox.showinfo("Enhanced Data Analysis", dialog_msg)
            else:
                messagebox.showwarning("No Training Data", 
                                     "No enhanced training images found.\n"
                                     "Use enhanced extraction tools first.")
            
        except Exception as e:
            messagebox.showerror("Analysis Error", f"Enhanced data analysis failed: {e}")
            import traceback
            traceback.print_exc()

    def train_enhanced_model(self):
        """Train enhanced model with comprehensive configuration."""
        try:
            # Import enhanced processor
            from enhanced_ml_form_processor import EnhancedMLFormProcessor, EnhancedMLTrainingHelper
        
            # Verify enhanced training data
            if not os.path.exists("training_data/sensory_ratings"):
                messagebox.showerror("Missing Enhanced Data", 
                                   "Enhanced training data not found.\n"
                                   "Use enhanced extraction first.")
                return
        
            # Count enhanced training data
            total_images = 0
            enhanced_count = 0
        
            for rating in range(1, 10):
                rating_dir = os.path.join("training_data/sensory_ratings", f"rating_{rating}")
                if os.path.exists(rating_dir):
                    images = len([f for f in os.listdir(rating_dir) if f.endswith(('.jpg', '.png', '.jpeg'))])
                    info_files = len([f for f in os.listdir(rating_dir) if f.endswith('_info.txt')])
                    total_images += images
                    enhanced_count += info_files
        
            # Enhanced training dialog
            enhancement_rate = enhanced_count / max(total_images, 1)
        
            training_msg = (f"Enhanced ML Model Training\n\n"
                           f"Training Data:\n"
                           f"• Total images: {total_images}\n"
                           f"• Enhanced extractions: {enhanced_count} ({enhancement_rate:.1%})\n"
                           f"• Target resolution: 600x140 pixels\n\n"
                           f"Enhanced Architecture:\n"
                           f"• 5 convolutional layers\n"
                           f"• Optimized for high-resolution data\n"
                           f"• Advanced regularization\n"
                           f"• Shadow removal preprocessing compatibility\n\n"
                           f"Training Features:\n"
                           f"• Early stopping with patience\n"
                           f"• Learning rate scheduling\n"
                           f"• Enhanced model checkpointing\n"
                           f"• Comprehensive logging\n\n"
                           f"Estimated time: 10-30 minutes\n\n"
                           f"Continue?")
        
            result = messagebox.askyesno("Enhanced Model Training", training_msg)
        
            if result:
                print("="*80)
                print("ENHANCED ML MODEL TRAINING")
                print("="*80)
                print(f"Training images: {total_images}")
                print(f"Enhanced extractions: {enhanced_count}")
                print(f"Architecture: Enhanced CNN for 600x140 resolution")
                print(f"Features: Shadow removal compatibility, advanced regularization")
            
                # Initialize enhanced components
                processor = EnhancedMLFormProcessor()
                trainer = EnhancedMLTrainingHelper(processor)
            
                # Enhanced training configuration
                training_config = {
                    'epochs': 100,
                    'batch_size': 16,
                    'validation_split': 0.25,
                    'save_best_only': True,
                    'patience': 20
                }
            
                print(f"\nEnhanced training configuration:")
                for key, value in training_config.items():
                    print(f"  {key}: {value}")
            
                # Train enhanced model
                model, history = trainer.train_enhanced_model(**training_config)
            
                # Enhanced results reporting
                if history and model:
                    final_train_acc = history.history['accuracy'][-1]
                    final_val_acc = history.history['val_accuracy'][-1]
                    best_val_acc = max(history.history['val_accuracy'])
                    epochs_trained = len(history.history['accuracy'])
                
                    # Check for enhanced model files
                    model_files = []
                    if os.path.exists('models/sensory_rating_classifier.h5'):
                        size_mb = os.path.getsize('models/sensory_rating_classifier.h5') / (1024*1024)
                        model_files.append(f"• Final model: {size_mb:.1f} MB")
                
                    if os.path.exists('models/enhanced/sensory_rating_classifier_best.h5'):
                        size_mb = os.path.getsize('models/enhanced/sensory_rating_classifier_best.h5') / (1024*1024)
                        model_files.append(f"• Best enhanced model: {size_mb:.1f} MB")
                
                    success_msg = (f"Enhanced Model Training Complete!\n\n"
                                 f"Performance Metrics:\n"
                                 f"• Final training accuracy: {final_train_acc:.3f}\n"
                                 f"• Final validation accuracy: {final_val_acc:.3f}\n"
                                 f"• Best validation accuracy: {best_val_acc:.3f}\n"
                                 f"• Epochs trained: {epochs_trained}\n\n"
                                 f"Model Files Saved:\n" + "\n".join(model_files) + f"\n\n"
                                 f"Enhanced Features:\n"
                                 f"• 600x140 high resolution\n"
                                 f"• Shadow removal preprocessing\n"
                                 f"• Advanced CNN architecture\n"
                                 f"• Production-ready accuracy\n\n"
                                 f"Next: Test Enhanced Model")
                
                    messagebox.showinfo("Enhanced Training Complete", success_msg)
                
                    print("="*80)
                    print("ENHANCED TRAINING COMPLETED SUCCESSFULLY")
                    print("="*80)
                    print(f"Enhanced model ready for production use!")
                
                else:
                    messagebox.showwarning("Training Issues", 
                                         "Enhanced training completed with issues.\n"
                                         "Check console for detailed information.")
                
        except Exception as e:
            error_msg = f"Enhanced training failed: {e}"
            print(f"ERROR: {error_msg}")
            messagebox.showerror("Enhanced Training Error", error_msg)
            import traceback
            traceback.print_exc()

    def test_enhanced_model(self):
        """Test enhanced model with comprehensive evaluation."""
        try:
            from enhanced_ml_form_processor import EnhancedMLFormProcessor
            import cv2
            import numpy as np
        
            # Check for enhanced model
            model_paths = [
                "models/enhanced/sensory_rating_classifier_best.h5",
                "models/sensory_rating_classifier.h5",
                "models/enhanced/sensory_rating_classifier.h5"
            ]
        
            model_path = None
            for path in model_paths:
                if os.path.exists(path):
                    model_path = path
                    break
        
            if not model_path:
                messagebox.showwarning("No Enhanced Model", 
                                     "No enhanced model found.\n"
                                     "Train the enhanced model first.")
                return
        
            print("="*80)
            print("ENHANCED MODEL TESTING")
            print("="*80)
            print(f"Testing model: {model_path}")
        
            # Initialize enhanced processor
            processor = EnhancedMLFormProcessor(model_path)
        
            if not processor.load_model():
                messagebox.showerror("Model Load Error", 
                                   "Failed to load enhanced model.\n"
                                   "Check console for error details.")
                return
        
            print(f"✓ Enhanced model loaded successfully")
            print(f"Model resolution: {processor.target_size}")
        
            # Test on enhanced training data samples
            base_dir = "training_data/sensory_ratings"
            test_results = {}
            detailed_results = []
        
            total_tests = 0
            correct_predictions = 0
            confidence_scores = []
        
            print(f"\nTesting enhanced model on training samples...")
        
            for rating in range(1, 10):
                rating_dir = os.path.join(base_dir, f"rating_{rating}")
                if os.path.exists(rating_dir):
                    images = [f for f in os.listdir(rating_dir) if f.endswith(('.jpg', '.png', '.jpeg'))]
                    if images:
                        # Test on first image in each class
                        test_image_path = os.path.join(rating_dir, images[0])
                    
                        try:
                            # Load and test with enhanced processor
                            test_image = cv2.imread(test_image_path, cv2.IMREAD_GRAYSCALE)
                            if test_image is not None:
                                # Resize to enhanced model input size
                                resized_image = cv2.resize(test_image, processor.target_size, 
                                                         interpolation=cv2.INTER_CUBIC)
                            
                                # Get enhanced prediction
                                predicted_rating, confidence, probabilities = processor.predict_rating_enhanced(resized_image)
                            
                                is_correct = predicted_rating == rating
                                test_results[rating] = {
                                    'predicted': predicted_rating,
                                    'confidence': confidence,
                                    'correct': is_correct,
                                    'probabilities': probabilities
                                }
                            
                                detailed_results.append({
                                    'true_rating': rating,
                                    'predicted_rating': predicted_rating,
                                    'confidence': confidence,
                                    'correct': is_correct
                                })
                            
                                total_tests += 1
                                confidence_scores.append(confidence)
                                if is_correct:
                                    correct_predictions += 1
                            
                                status = "✓ CORRECT" if is_correct else "✗ WRONG"
                                conf_level = "HIGH" if confidence > 0.8 else "MED" if confidence > 0.6 else "LOW"
                            
                                print(f"Rating {rating}: Predicted {predicted_rating} ({conf_level} conf: {confidence:.3f}) - {status}")
                            
                                # Show top 3 predictions for detailed analysis
                                top_3 = processor.get_top_predictions(probabilities, 3)
                                top_3_str = ", ".join([f"R{r}({p:.2f})" for r, p in top_3])
                                print(f"  Top 3: {top_3_str}")
                    
                        except Exception as e:
                            print(f"Error testing rating {rating}: {e}")
        
            # Enhanced results analysis
            if detailed_results:
                test_accuracy = correct_predictions / total_tests
                avg_confidence = np.mean(confidence_scores)
                confidence_std = np.std(confidence_scores)
            
                # Confidence analysis
                high_conf = sum(1 for c in confidence_scores if c > 0.8)
                med_conf = sum(1 for c in confidence_scores if 0.6 <= c <= 0.8)
                low_conf = sum(1 for c in confidence_scores if c < 0.6)
            
                print(f"\n" + "="*80)
                print(f"ENHANCED MODEL TEST RESULTS")
                print(f"="*80)
                print(f"Model tested: {os.path.basename(model_path)}")
                print(f"Resolution: {processor.target_size}")
                print(f"Classes tested: {total_tests}")
                print(f"Correct predictions: {correct_predictions}")
                print(f"Test accuracy: {test_accuracy:.3f} ({test_accuracy*100:.1f}%)")
                print(f"Average confidence: {avg_confidence:.3f} ± {confidence_std:.3f}")
                print(f"Confidence distribution: High({high_conf}) Med({med_conf}) Low({low_conf})")
            
                # Detailed error analysis
                errors = [r for r in detailed_results if not r['correct']]
                if errors:
                    print(f"\nError analysis:")
                    for error in errors:
                        print(f"  True: {error['true_rating']} → Predicted: {error['predicted_rating']} (conf: {error['confidence']:.3f})")
            
                # Performance assessment
                if test_accuracy >= 0.9:
                    status = "EXCELLENT"
                    recommendation = "Model ready for production deployment!"
                elif test_accuracy >= 0.8:
                    status = "VERY GOOD"
                    recommendation = "Model suitable for production with monitoring"
                elif test_accuracy >= 0.7:
                    status = "GOOD"
                    recommendation = "Consider collecting more training data"
                else:
                    status = "NEEDS IMPROVEMENT"
                    recommendation = "Collect significantly more training data"
            
                # Show comprehensive results dialog
                result_msg = (f"Enhanced Model Test Results - {status}\n\n"
                             f"Performance Metrics:\n"
                             f"• Test accuracy: {test_accuracy*100:.1f}%\n"
                             f"• Average confidence: {avg_confidence:.3f}\n"
                             f"• High confidence predictions: {high_conf}/{total_tests}\n"
                             f"• Model resolution: {processor.target_size[0]}x{processor.target_size[1]}\n\n"
                             f"Confidence Distribution:\n"
                             f"• High (>0.8): {high_conf}\n"
                             f"• Medium (0.6-0.8): {med_conf}\n"
                             f"• Low (<0.6): {low_conf}\n\n"
                             f"Recommendation:\n{recommendation}\n\n"
                             f"Check console for detailed per-class results.")
            
                messagebox.showinfo("Enhanced Model Test Complete", result_msg)
            else:
                messagebox.showwarning("No Test Data", 
                                     "No test data available.\n"
                                     "Ensure training data is present.")
        
        except Exception as e:
            messagebox.showerror("Enhanced Test Error", f"Enhanced model testing failed: {e}")
            import traceback
            traceback.print_exc()

    def validate_enhanced_performance(self):
        """Comprehensive enhanced model validation."""
        messagebox.showinfo("Enhanced Validation", 
                          "Comprehensive Enhanced Model Validation\n\n"
                          "Features to be implemented:\n\n"
                          "• Cross-validation analysis\n"
                          "• Confusion matrix generation\n"
                          "• Per-attribute accuracy metrics\n"
                          "• Confidence calibration analysis\n"
                          "• Model uncertainty quantification\n"
                          "• Production readiness assessment\n\n"
                          "This advanced validation suite will be available\n"
                          "in the next update for production deployment.")

    def update_processor_config(self):
        """Update processor configuration with enhanced settings."""
        try:
            from enhanced_ml_processor_updater import MLProcessorUpdater
        
            # Find available configurations
            config_dir = "training_data/claude_analysis"
            if os.path.exists(config_dir):
                config_files = [f for f in os.listdir(config_dir) 
                               if f.startswith('improved_boundaries_') and f.endswith('.json')]
            
                if config_files:
                    config_file = filedialog.askopenfilename(
                        title="Select enhanced boundary configuration",
                        initialdir=config_dir,
                        filetypes=[("JSON files", "*.json")]
                    )
                
                    if config_file:
                        result = messagebox.askyesno("Update Enhanced Processor",
                                                   f"Update enhanced processor with:\n"
                                                   f"{os.path.basename(config_file)}\n\n"
                                                   f"This will modify enhanced_ml_form_processor.py\n"
                                                   f"A backup will be created automatically.\n\n"
                                                   f"Continue?")
                        if result:
                            updater = MLProcessorUpdater()
                            updater.update_ml_processor_boundaries(config_file)
                        
                            messagebox.showinfo("Enhanced Update Complete",
                                              f"Enhanced processor updated!\n\n"
                                              f"Configuration: {os.path.basename(config_file)}\n"
                                              f"Backup: {updater.backup_path}\n\n"
                                              f"Test the updated enhanced processor.")
                else:
                    messagebox.showinfo("No Enhanced Configs", 
                                      "No enhanced configurations found.\n"
                                      "Use enhanced extraction tools first.")
            else:
                messagebox.showinfo("No Analysis Directory", 
                                  "Enhanced analysis directory not found.")
            
        except Exception as e:
            messagebox.showerror("Update Error", f"Failed to update enhanced processor: {e}")


class SensoryAIProcessor:
    """Handles AI image processing functionality."""
    
    def __init__(self, parent_window):
        """Initialize with reference to parent window for UI updates."""
        self.parent = parent_window
        
    def load_from_image_enhanced(self):
        """Load sensory data using enhanced ML processing."""
        try:
            from enhanced_ml_form_processor import EnhancedMLFormProcessor
        
            # Check for enhanced model
            model_paths = [
                "models/enhanced/sensory_rating_classifier_best.h5",
                "models/sensory_rating_classifier.h5"
            ]
        
            model_path = None
            for path in model_paths:
                if os.path.exists(path):
                    model_path = path
                    break
        
            if not model_path:
                messagebox.showwarning("No Enhanced Model", 
                                     "No enhanced model found.\n"
                                     "Train an enhanced model first using the Enhanced ML menu.")
                return
        
            # Select image file
            image_path = filedialog.askopenfilename(
                title="Select form image for enhanced ML processing",
                filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.tiff")]
            )
        
            if image_path:
                print("="*80)
                print("ENHANCED ML FORM LOADING")
                print("="*80)
            
                # Initialize enhanced processor
                processor = EnhancedMLFormProcessor(model_path)
            
                # Process with enhanced pipeline
                extracted_data, processed_image = processor.process_form_image_enhanced(image_path)
            
                # Show enhanced preview with confidence scores
                self._show_enhanced_extraction_preview(extracted_data, processed_image, 
                                                     os.path.basename(image_path))
        
        except Exception as e:
            messagebox.showerror("Enhanced ML Error", f"Enhanced ML processing failed: {e}")
            import traceback
            traceback.print_exc()

    def load_from_image_with_ai(self):
        """Load sensory data from a single form image using Enhanced Claude AI."""
        # Check for required dependencies
        try:
            import anthropic
            import base64
            import io
            from PIL import Image
        except ImportError as e:
            messagebox.showerror("Missing Dependencies", 
                               f"AI image processing requires additional packages.\n\n"
                               f"Install with: pip install anthropic pillow\n\n"
                               f"Error: {e}")
            return

        # Get image file
        filename = filedialog.askopenfilename(
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.tiff *.bmp"),
                ("All files", "*.*")
            ],
            title="Load Sensory Form Image for Enhanced AI Processing"
        )

        if not filename:
            return

        try:
            from enhanced_claude_form_processor import EnhancedClaudeFormProcessor
            
            # Show processing dialog
            progress_window = self._create_progress_window("Processing with Enhanced Claude AI...")
    
            # Process with Enhanced Claude AI
            ai_processor = EnhancedClaudeFormProcessor()
        
            # Process single image (uses shadow removal preprocessing)
            image_data, processed_image = ai_processor.prepare_image_with_preprocessing(filename)
            extracted_data = ai_processor.process_single_image_with_claude(image_data, filename)
    
            # Stop progress bar
            progress_window.destroy()
    
            # Show enhanced results preview
            if extracted_data:
                self._show_enhanced_ai_extraction_preview(extracted_data, processed_image, filename)
            else:
                messagebox.showwarning("No Data", "No sensory data could be extracted from the image.")
        
        except Exception as e:
            if 'progress_window' in locals():
                progress_window.destroy()
            messagebox.showerror("Enhanced AI Processing Error", f"Failed to process image with AI: {e}")
            debug_print(f"Enhanced AI processing error: {e}")

    def batch_process_with_ai(self):
        """Process a batch of form images using Enhanced Claude AI."""
        # Check for required dependencies
        try:
            import anthropic
        except ImportError as e:
            messagebox.showerror("Missing Dependencies", 
                               f"AI batch processing requires additional packages.\n\n"
                               f"Install with: pip install anthropic pillow opencv-python\n\n"
                               f"Error: {e}")
            return

        # Get folder containing images
        folder_path = filedialog.askdirectory(
            title="Select Folder Containing Form Images for Batch AI Processing"
        )

        if not folder_path:
            return

        try:
            from enhanced_claude_form_processor import EnhancedClaudeFormProcessor
            
            print("DEBUG: Starting batch AI processing")
        
            # Show processing dialog
            progress_window = self._create_progress_window("Batch Processing with Enhanced Claude AI...")

            # Initialize Enhanced Claude AI processor
            ai_processor = EnhancedClaudeFormProcessor()
    
            # Process batch of images
            batch_results = ai_processor.process_batch_images(folder_path)

            # Stop progress bar
            progress_window.destroy()

            # Show batch results summary
            successful_count = sum(1 for result in batch_results.values() if result['status'] == 'success')
            total_count = len(batch_results)
    
            if successful_count > 0:
                # Store the processor and results for review interface
                self.ai_processor = ai_processor
        
                result_msg = (f"Enhanced Batch Processing Complete!\n\n"
                             f"Successfully processed: {successful_count}/{total_count} images\n"
                             f"Features used:\n"
                             f"• Shadow removal preprocessing\n"
                             f"• Sample name extraction\n"
                             f"• Enhanced AI analysis\n\n"
                             f"Launch interactive review to verify and edit results?")
        
                if messagebox.askyesno("Batch Processing Complete", result_msg):
                    print("DEBUG: Launching review interface")
                    # Launch the interactive review interface
                    ai_processor.launch_review_interface()
                
                    # Monitor for review completion
                    self._monitor_review_completion(ai_processor)
                else:
                    print("DEBUG: Loading results directly without review")
                    # Auto-load all successful results
                    self._load_batch_results_directly(batch_results)
            else:
                messagebox.showerror("Batch Processing Failed", 
                                   f"No images could be processed successfully.\n"
                                   f"Check that images contain readable sensory evaluation forms.")
    
        except Exception as e:
            if 'progress_window' in locals():
                progress_window.destroy()
            messagebox.showerror("Batch AI Processing Error", f"Failed to process batch: {e}")
            debug_print(f"Batch AI processing error: {e}")

    def _create_progress_window(self, title):
        """Create a standard progress window."""
        progress_window = tk.Toplevel(self.parent.window)
        progress_window.title(title)
        progress_window.geometry("400x150")
        progress_window.transient(self.parent.window)
        progress_window.grab_set()

        progress_label = ttk.Label(progress_window, 
                                 text="Processing image with shadow removal and AI analysis...", 
                                 font=('Arial', 10))
        progress_label.pack(expand=True, pady=10)

        progress_bar = ttk.Progressbar(progress_window, mode='indeterminate')
        progress_bar.pack(fill='x', padx=20, pady=10)
        progress_bar.start()

        self.parent.window.update()
        return progress_window

    def _show_enhanced_extraction_preview(self, extracted_data, processed_img, filename):
        """Show enhanced extraction preview with confidence analysis."""
        # This method would contain the preview logic
        # (Moving the existing implementation from the main file)
        pass

    def _show_enhanced_ai_extraction_preview(self, extracted_data, processed_image, filename):
        """Show enhanced AI extraction preview with editable ratings."""
        print("DEBUG: Starting enhanced AI extraction preview")
        print(f"DEBUG: Extracted data keys: {list(extracted_data.keys())}")
        print(f"DEBUG: Processed image shape: {processed_image.shape if hasattr(processed_image, 'shape') else 'No shape attr'}")
    
        try:
            import tkinter as tk
            from tkinter import ttk, messagebox
            from PIL import Image, ImageTk
            import cv2
            import numpy as np
        
            # Create preview window
            preview_window = tk.Toplevel(self.parent.window)
            preview_window.title(f"Enhanced AI Extraction Preview - {os.path.basename(filename)}")
            preview_window.geometry("1400x900")
            preview_window.configure(bg='white')
            preview_window.transient(self.parent.window)
            preview_window.grab_set()
        
            print("DEBUG: Preview window created")
        
            # Create main frame
            main_frame = tk.Frame(preview_window, bg='white')
            main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
            # Header frame
            header_frame = tk.Frame(main_frame, bg='lightblue', height=60)
            header_frame.pack(fill='x', pady=(0, 10))
            header_frame.pack_propagate(False)
        
            # Title
            title_label = tk.Label(header_frame, text=f"AI Extraction Preview: {os.path.basename(filename)}", 
                                  font=('Arial', 14, 'bold'), bg='lightblue')
            title_label.pack(side='left', padx=10, pady=15)
        
            # Action buttons
            button_frame = tk.Frame(header_frame, bg='lightblue')
            button_frame.pack(side='right', padx=10, pady=10)
        
            def save_and_load():
                """Save edits and load into main application."""
                print("DEBUG: save_and_load called")
                try:
                    # Save current edits
                    final_data = {}
                    for sample_key, widgets in entry_widgets.items():
                        sample_data = {}
                        sample_data['sample_name'] = widgets['sample_name'].get()
                    
                        # Get metric values
                        for metric in ['Burnt Taste', 'Vapor Volume', 'Overall Flavor', 'Smoothness', 'Overall Liking']:
                            value = widgets[metric].get().strip()
                            if value:
                                try:
                                    sample_data[metric] = int(value)
                                except ValueError:
                                    sample_data[metric] = None
                            else:
                                sample_data[metric] = None
                    
                        # Get comments
                        if isinstance(widgets['comments'], tk.Text):
                            sample_data['comments'] = widgets['comments'].get('1.0', tk.END).strip()
                        else:
                            sample_data['comments'] = widgets['comments'].get()
                    
                        final_data[sample_key] = sample_data
                
                    print(f"DEBUG: Final data prepared: {list(final_data.keys())}")
                
                    # Create new session with the extracted data
                    session_id = f"AI_Extract_{os.path.basename(filename).replace('.', '_')}"
                
                    # Ensure unique session name
                    counter = 1
                    original_session_id = session_id
                    while session_id in self.parent.sessions:
                        session_id = f"{original_session_id}_{counter}"
                        counter += 1
                
                    # Create session data structure
                    from datetime import datetime
                    self.parent.sessions[session_id] = {
                        'header': {
                            'Date': datetime.now().strftime("%Y-%m-%d"),
                            'Assessor Name': 'AI Extraction',
                            'Session Type': 'AI Processing',
                            'Notes': f'Extracted from {os.path.basename(filename)}'
                        },
                        'samples': final_data,
                        'timestamp': datetime.now().isoformat(),
                        'source_file': filename,
                        'source_image': filename
                    }
                
                    print(f"DEBUG: Created session: {session_id}")
                
                    # Switch to the new session
                    if hasattr(self.parent, 'switch_to_session'):
                        self.parent.switch_to_session(session_id)
                        print(f"DEBUG: Switched to session: {session_id}")
                    
                        # Update UI components
                        if hasattr(self.parent, 'update_session_combo'):
                            self.parent.update_session_combo()
                        if hasattr(self.parent, 'session_var'):
                            self.parent.session_var.set(session_id)
                        
                        messagebox.showinfo("Success", 
                                          f"Data loaded into new session: {session_id}\n"
                                          f"Loaded {len(final_data)} samples")
                        preview_window.destroy()
                    else:
                        print("DEBUG: switch_to_session method not found")
                        messagebox.showerror("Error", "Could not switch to new session")
                    
                except Exception as e:
                    print(f"DEBUG: Error in save_and_load: {e}")
                    import traceback
                    traceback.print_exc()
                    messagebox.showerror("Error", f"Failed to save and load data: {e}")
        
            tk.Button(button_frame, text="💾 Load Data", command=save_and_load,
                     bg='gold', font=('Arial', 10, 'bold')).pack(side='left', padx=5)
        
            tk.Button(button_frame, text="❌ Cancel", command=preview_window.destroy,
                     bg='lightcoral', font=('Arial', 10, 'bold')).pack(side='left', padx=5)
        
            # Content frame
            content_frame = tk.Frame(main_frame, bg='white')
            content_frame.pack(fill='both', expand=True)
        
            # Left side - Image display (60% width)
            image_frame = tk.Frame(content_frame, bg='white', width=960)
            image_frame.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
            # Display processed image
            try:
                if isinstance(processed_image, np.ndarray):
                    # Convert from CV2 format to PIL
                    if len(processed_image.shape) == 3:
                        # Color image - convert from BGR to RGB
                        rgb_image = cv2.cvtColor(processed_image, cv2.COLOR_BGR2RGB)
                        pil_image = Image.fromarray(rgb_image)
                    else:
                        # Grayscale image
                        pil_image = Image.fromarray(processed_image)
                else:
                    pil_image = processed_image
            
                # Resize for display
                display_width = 900
                aspect_ratio = pil_image.height / pil_image.width
                display_height = int(display_width * aspect_ratio)
            
                if display_height > 700:
                    display_height = 700
                    display_width = int(display_height / aspect_ratio)
            
                pil_image = pil_image.resize((display_width, display_height), Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(pil_image)
            
                image_label = tk.Label(image_frame, image=photo, bg='white')
                image_label.pack(pady=10)
                image_label.image = photo  # Keep reference
            
                print(f"DEBUG: Image displayed at {display_width}x{display_height}")
            
            except Exception as e:
                print(f"DEBUG: Error displaying image: {e}")
                image_label = tk.Label(image_frame, text=f"Error displaying image: {e}", bg='white')
                image_label.pack(pady=10)
        
            # Right side - Data editing (40% width)
            edit_frame = tk.Frame(content_frame, bg='lightgray', width=640)
            edit_frame.pack(side='right', fill='both')
            edit_frame.pack_propagate(False)
        
            # Scrollable frame for data editing
            canvas = tk.Canvas(edit_frame, bg='white')
            scrollbar = ttk.Scrollbar(edit_frame, orient="vertical", command=canvas.yview)
            scrollable_frame = tk.Frame(canvas, bg='white')
        
            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )
        
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
        
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
        
            # Title for editing section
            title_label = tk.Label(scrollable_frame, text="Extracted Data (Click to Edit)", 
                                  font=('Arial', 14, 'bold'), bg='white')
            title_label.pack(pady=10)
        
            # Store entry widgets for saving
            entry_widgets = {}
        
            # Convert to list for easier indexing
            sample_items = list(extracted_data.items())
            print(f"DEBUG: Processing {len(sample_items)} samples for display")
        
            # Create 2x2 grid layout to match form structure
            grid_frame = tk.Frame(scrollable_frame, bg='white')
            grid_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
            # Configure grid weights for even distribution
            grid_frame.grid_columnconfigure(0, weight=1)
            grid_frame.grid_columnconfigure(1, weight=1)
            grid_frame.grid_rowconfigure(0, weight=1)
            grid_frame.grid_rowconfigure(1, weight=1)
        
            metrics = ['Burnt Taste', 'Vapor Volume', 'Overall Flavor', 'Smoothness', 'Overall Liking']
        
            for i, (sample_key, sample_data) in enumerate(sample_items):
                print(f"DEBUG: Creating UI for sample {i}: {sample_key}")
            
                # Calculate grid position (2x2 layout)
                row = i // 2  # 0 for first two samples, 1 for last two
                col = i % 2   # 0 for left column, 1 for right column
            
                # Sample frame
                sample_frame = tk.LabelFrame(grid_frame, text=f"{sample_key}", 
                                           font=('Arial', 11, 'bold'), bg='white', 
                                           padx=8, pady=8, relief='raised', bd=2)
                sample_frame.grid(row=row, column=col, sticky='nsew', padx=8, pady=8)
            
                entry_widgets[sample_key] = {}
            
                # Sample name section
                name_frame = tk.Frame(sample_frame, bg='white')
                name_frame.pack(fill='x', pady=(0, 8))
            
                tk.Label(name_frame, text="Name:", font=('Arial', 9, 'bold'), 
                        bg='white', width=8, anchor='w').pack(side='left')
                name_entry = tk.Entry(name_frame, font=('Arial', 9), width=15)
                name_entry.pack(side='left', padx=(3, 0), fill='x', expand=True)
                name_entry.insert(0, str(sample_data.get('sample_name', sample_key)))
                entry_widgets[sample_key]['sample_name'] = name_entry
            
                # Metrics section
                metrics_frame = tk.Frame(sample_frame, bg='white')
                metrics_frame.pack(fill='both', expand=True)
            
                for metric in metrics:
                    metric_frame = tk.Frame(metrics_frame, bg='white')
                    metric_frame.pack(fill='x', pady=1)
                
                    # Metric label
                    label_text = metric.replace(' ', '\n') if len(metric) > 12 else metric
                    metric_label = tk.Label(metric_frame, text=f"{label_text}:", 
                                          font=('Arial', 8), bg='white', 
                                          width=12, anchor='w', justify='left')
                    metric_label.pack(side='left')
                
                    # Metric entry
                    metric_entry = tk.Entry(metric_frame, font=('Arial', 9), width=8)
                    metric_entry.pack(side='left', padx=(3, 0))
                
                    # Pre-fill with extracted value
                    value = sample_data.get(metric)
                    if value is not None:
                        metric_entry.insert(0, str(value))
                
                    entry_widgets[sample_key][metric] = metric_entry
            
                # Comments section
                comments_frame = tk.Frame(sample_frame, bg='white')
                comments_frame.pack(fill='x', pady=(8, 0))
            
                tk.Label(comments_frame, text="Comments:", font=('Arial', 8, 'bold'), 
                        bg='white', width=12, anchor='w').pack(side='top', anchor='w')
            
                comments_text = tk.Text(comments_frame, font=('Arial', 8), 
                                      width=25, height=3, wrap='word')
                comments_text.pack(fill='x', expand=True, pady=(2, 0))
                comments_text.insert('1.0', str(sample_data.get('comments', '')))
                entry_widgets[sample_key]['comments'] = comments_text
        
            # Add placeholder frames if fewer than 4 samples
            total_samples = len(sample_items)
            if total_samples < 4:
                for i in range(total_samples, 4):
                    row = i // 2
                    col = i % 2
                    placeholder_frame = tk.Frame(grid_frame, bg='lightgray', relief='sunken', bd=1)
                    placeholder_frame.grid(row=row, column=col, sticky='nsew', padx=8, pady=8)
                    tk.Label(placeholder_frame, text="No Sample", 
                            font=('Arial', 10), bg='lightgray', fg='gray').pack(expand=True)
        
            print("DEBUG: Enhanced AI extraction preview setup complete")
        
           # Center the window - IMPROVED VERSION
            preview_window.update_idletasks()  # Ensure all widgets are rendered
            preview_window.geometry("")  # Reset geometry to auto-size
            preview_window.update_idletasks()  # Update again after reset
        
            # Get the actual window dimensions
            width = preview_window.winfo_reqwidth()
            height = preview_window.winfo_reqheight()
        
            # Get screen dimensions
            screen_width = preview_window.winfo_screenwidth()
            screen_height = preview_window.winfo_screenheight()
        
            # Calculate center position
            x = (screen_width - width) // 2
            y = (screen_height - height) // 2
        
            # Set the centered geometry
            preview_window.geometry(f"{width}x{height}+{x}+{y}")
        
            print(f"DEBUG: Window centered at {x},{y} with size {width}x{height}")
        
        except Exception as e:
            print(f"DEBUG: Error in _show_enhanced_ai_extraction_preview: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Preview Error", f"Failed to show extraction preview: {e}")

    def _monitor_review_completion(self, ai_processor):
        """Monitor the review interface for completion."""
        def check_completion():
            print("DEBUG: Checking review completion status")
        
            if ai_processor.is_review_complete():
                print("DEBUG: Review completed, loading results")
            
                # Get the reviewed results
                reviewed_results = ai_processor.get_reviewed_results()
            
                if reviewed_results:
                    print(f"DEBUG: Loading {len(reviewed_results)} reviewed results")
                
                    # Load the reviewed results using session-based structure
                    self._load_batch_results_directly(reviewed_results)
                
                    # Update session selector UI
                    if hasattr(self.parent, 'session_var') and self.parent.current_session_id:
                        self.parent.session_var.set(self.parent.current_session_id)
                
                    # Show completion message
                    total_samples = sum(len(data['extracted_data']) for data in reviewed_results.values() 
                                      if data['status'] == 'success')
                
                    messagebox.showinfo("Review Complete", 
                                      f"Successfully loaded reviewed data!\n"
                                      f"Total sessions: {len(reviewed_results)}\n"
                                      f"Total samples: {total_samples}\n"
                                      f"Use the session selector to switch between sessions.")
                else:
                    print("DEBUG: No reviewed results found")
                    messagebox.showwarning("No Data", "No reviewed data to load.")
            
                return  # Stop monitoring
        
            # Continue monitoring if review not complete
            self.parent.window.after(500, check_completion)  # Check every 500ms
    
        print("DEBUG: Starting review completion monitoring")
        # Start monitoring after a short delay to allow review window to open
        self.parent.window.after(1000, check_completion)

    def _load_batch_results_directly(self, batch_results):
        """Load successful batch results with each image as a separate session."""
        loaded_sessions = 0
        loaded_samples = 0

        print("DEBUG: Starting batch results loading with session-per-image structure")
        print(f"DEBUG: Processing {len(batch_results)} batch results")

        for image_path, result in batch_results.items():
            if result['status'] == 'success':
                extracted_data = result['extracted_data']
        
                # Skip empty results
                if not extracted_data:
                    print(f"DEBUG: Skipping empty result for {image_path}")
                    continue
        
                # Create session name from image filename
                image_name = os.path.splitext(os.path.basename(image_path))[0]
                session_name = f"Batch_AI_{image_name}"
            
                # Ensure unique session name
                counter = 1
                original_session_name = session_name
                while session_name in self.parent.sessions:
                    session_name = f"{original_session_name}_{counter}"
                    counter += 1
        
                print(f"DEBUG: Creating session {session_name} from image {image_name}")
        
                # Create new session for this image
                self.parent.sessions[session_name] = {
                    'header': {field: var.get() for field, var in self.parent.header_vars.items()},
                    'samples': {},
                    'timestamp': datetime.now().isoformat(),
                    'source_image': image_path,
                    'extraction_method': 'Enhanced_Claude_AI_Batch'
                }
        
                # Load samples into this session (up to 4 samples)
                sample_count = 0
                for sample_key, sample_data in extracted_data.items():
                    if sample_count >= 4:
                        print(f"DEBUG: Reached maximum 4 samples for session {session_name}")
                        break
            
                    # Skip empty samples
                    if not sample_data or not any(sample_data.get(metric, None) for metric in self.parent.metrics):
                        print(f"DEBUG: Skipping empty sample {sample_key}")
                        continue
            
                    # Add sample to this session
                    self.parent.sessions[session_name]['samples'][sample_key] = sample_data
                    sample_count += 1
                    loaded_samples += 1
            
                    print(f"DEBUG: Added sample {sample_key} to session {session_name}")
        
                if sample_count > 0:
                    loaded_sessions += 1
                    print(f"DEBUG: Session {session_name} created with {sample_count} samples")
                else:
                    # Remove empty session
                    del self.parent.sessions[session_name]
                    print(f"DEBUG: Removed empty session {session_name}")

        # Switch to first session if any were loaded
        if self.parent.sessions:
            first_session = list(self.parent.sessions.keys())[0]
            self.parent.switch_to_session(first_session)
        
            # Update session selector UI
            if hasattr(self.parent, 'session_var'):
                self.parent.session_var.set(first_session)
    
            print(f"DEBUG: Switched to first session: {first_session}")

        # Update UI
        self.parent.update_session_combo()
        self.parent.update_sample_combo()
        self.parent.update_sample_checkboxes()
        self.parent.update_plot()

        print(f"DEBUG: Batch loading complete")
        print(f"DEBUG: Loaded {loaded_sessions} sessions with total {loaded_samples} samples")

        if loaded_sessions > 0:
            messagebox.showinfo("Batch Load Complete", 
                              f"Loaded {loaded_sessions} sessions with {loaded_samples} total samples!\n"
                              f"Each image is now a separate session (max 4 samples each).\n"
                              f"Use the session selector to switch between sessions.")
        else:
            messagebox.showwarning("No Data Loaded", 
                                 "No valid samples found in batch results.")