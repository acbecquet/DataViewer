"""
ml_form_processor.py
Machine Learning Form Processor for Sensory Data Collection
Integrates with the training data pipeline to provide automated form reading
"""

import cv2
import numpy as np
import tensorflow as tf
import os
import pickle
from datetime import datetime
import json

# Lazy imports for TensorFlow to avoid startup delays
def _lazy_import_tensorflow():
    """Lazy import TensorFlow to avoid loading it unless actually needed."""
    try:
        import tensorflow as tf
        from tensorflow import keras
        from tensorflow.keras import layers
        print("TIMING: Lazy loaded TensorFlow for ML processing")
        return tf, keras, layers
    except ImportError as e:
        print(f"Error importing TensorFlow: {e}")
        raise ImportError("TensorFlow is required for ML functionality. Install with: pip install tensorflow")


class MLFormProcessor:
    """
    Production ML processor for extracting sensory ratings from phone camera images.
    
    This class handles the complete pipeline from raw phone camera image to
    extracted sensory data, using the CNN model trained from your labeled examples.
    The architecture is designed to be robust to the natural variations in
    phone camera images while maintaining high accuracy.
    """
    
    def __init__(self, model_path='models/sensory_rating_classifier.h5'):
        self.model_path = model_path
        self.model = None
        self.is_loaded = False
        
        # These match your training data extractor settings
        self.target_size = (500,120)  # Standard input size for our CNN
        self.num_classes = 9  # Ratings 1-9
        
        # Form layout parameters (matches your extractor)
        self.attributes = [
            "Burnt Taste", "Vapor Volume", "Overall Flavor", 
            "Smoothness", "Overall Liking"
        ]
        
        print("MLFormProcessor initialized")
        print(f"Model path: {model_path}")
        print(f"Target image size: {self.target_size}")
        
    def load_model(self):
        """
        Load the trained CNN model for rating classification.
        
        This method handles loading the trained model and preparing it for inference.
        It includes error handling for common deployment issues and validates
        that the model architecture matches our expectations.
        """
        tf,keras,layers = _lazy_import_tensorflow()
        if self.is_loaded:
            return True
            
        if not os.path.exists(self.model_path):
            print(f"Model file not found: {self.model_path}")
            print("Please train the model first using the training pipeline")
            return False
            
        try:
            # Load the trained model
            self.model = keras.models.load_model(self.model_path)
            
            # Validate model architecture
            expected_input_shape = (None, self.target_size[0], self.target_size[1], 1)  # Grayscale
            actual_input_shape = self.model.input_shape
            
            if actual_input_shape != expected_input_shape:
                print(f"Warning: Model input shape {actual_input_shape} doesn't match expected {expected_input_shape}")
                
            expected_output_shape = (None, self.num_classes)
            actual_output_shape = self.model.output_shape
            
            if actual_output_shape != expected_output_shape:
                print(f"Warning: Model output shape {actual_output_shape} doesn't match expected {expected_output_shape}")
                
            self.is_loaded = True
            print(f"Model loaded successfully")
            print(f"Input shape: {actual_input_shape}")
            print(f"Output shape: {actual_output_shape}")
            
            return True
            
        except Exception as e:
            print(f"Error loading model: {e}")
            return False
    
    def preprocess_image(self, image_path):
        """
        Preprocess a phone camera image for region extraction.
        
        This preprocessing pipeline is designed to handle the characteristics
        of phone camera images while preparing them for accurate region extraction.
        The key is to enhance the image for computer vision while preserving
        the natural characteristics that the model was trained on.
        """
        
        # Load the image with error handling
        img = cv2.imread(image_path)
        if img is None:
            raise ValueError(f"Could not load image: {image_path}")
            
        print(f"Original image size: {img.shape[1]}x{img.shape[0]}")
        
        # Convert to grayscale (matches training data)
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        
        # Apply gentle preprocessing that enhances readability without
        # removing the natural characteristics the model expects
        
        # Enhance contrast using CLAHE (Contrast Limited Adaptive Histogram Equalization)
        # This helps with varying lighting conditions while preserving local details
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
        enhanced = clahe.apply(gray)
        
        # Light denoising to reduce camera sensor noise
        denoised = cv2.bilateralFilter(enhanced, 5, 50, 50)
        
        # Detect and correct slight skew (common in hand-held photography)
        corrected = self.correct_minor_skew(denoised)
        
        return corrected
    
    def correct_minor_skew(self, image):
        """
        Detect and correct minor skew that's common in phone camera images.
        
        This method handles the small rotational variations that naturally occur
        when photographing documents by hand, even when trying to be careful.
        We only correct significant skew to avoid over-processing.
        """
        
        try:
            # Find edges for line detection
            edges = cv2.Canny(image, 50, 150, apertureSize=3)
            
            # Use Hough transform to find dominant lines
            lines = cv2.HoughLines(edges, 1, np.pi/180, threshold=100)
            
            if lines is not None and len(lines) > 5:
                # Extract angles from detected lines
                angles = []
                for line in lines[:20]:  # Use first 20 lines for stability
                    rho, theta = line[0]
                    angle = theta * 180 / np.pi
                    
                    # Normalize to [-45, 45] range
                    if angle > 45:
                        angle = angle - 90
                    elif angle < -45:
                        angle = angle + 90
                        
                    angles.append(angle)
                
                # Use median to avoid outlier influence
                if angles:
                    skew_angle = np.median(angles)
                    
                    # Only correct if skew is significant (>1 degree)
                    if abs(skew_angle) > 1.0:
                        print(f"Correcting skew: {skew_angle:.2f} degrees")
                        return self.rotate_image(image, skew_angle)
                        
        except Exception as e:
            print(f"Skew correction failed: {e}")
            
        return image
    
    def rotate_image(self, image, angle):
        """Rotate image by specified angle with proper handling of boundaries."""
        
        (h, w) = image.shape[:2]
        center = (w // 2, h // 2)
        
        # Calculate rotation matrix
        M = cv2.getRotationMatrix2D(center, angle, 1.0)
        
        # Perform rotation with border reflection to avoid black edges
        rotated = cv2.warpAffine(image, M, (w, h), 
                               flags=cv2.INTER_CUBIC,
                               borderMode=cv2.BORDER_REFLECT)
        return rotated
    
# ml_form_processor.py - Update these specific sections

class MLFormProcessor:
    def __init__(self, model_path='models/sensory_rating_classifier.h5'):
        self.model_path = model_path
        self.model = None
        self.is_loaded = False
        
        # MUST match the training extractor exactly
        self.target_size = (500, 120)  # Updated to match training
        self.num_classes = 9
        
        self.attributes = [
            "Burnt Taste", "Vapor Volume", "Overall Flavor", 
            "Smoothness", "Overall Liking"
        ]
        
        print("MLFormProcessor initialized with high-resolution support")
        print(f"Target image size: {self.target_size}")

    def extract_regions_for_prediction(self, processed_image):
        """
        Extract regions using the EXACT same parameters as training.
        This must match the improved_training_extractor.py boundaries exactly.
        """
        
        height, width = processed_image.shape
        regions = {}
        
        # EXACT same boundaries as training extractor
        sample_regions = {
            1: {  # Top-left sample
                "y_start": int(height * 0.28), "y_end": int(height * 0.56), 
                "x_start": int(width * 0.05), "x_end": int(width * 0.50)  # Extended to 50%
            },
            2: {  # Top-right sample  
                "y_start": int(height * 0.28), "y_end": int(height * 0.56),
                "x_start": int(width * 0.50), "x_end": int(width * 0.98)  # Extended to 98%
            },
            3: {  # Bottom-left sample
                "y_start": int(height * 0.64), "y_end": int(height * 0.90),
                "x_start": int(width * 0.05), "x_end": int(width * 0.50)  # Extended to 50%
            },
            4: {  # Bottom-right sample
                "y_start": int(height * 0.64), "y_end": int(height * 0.90),
                "x_start": int(width * 0.50), "x_end": int(width * 0.98)  # Extended to 98%
            }
        }
        
        # Extract regions with EXACT same logic as training
        for sample_id, sample_region in sample_regions.items():
            sample_height = sample_region["y_end"] - sample_region["y_start"]
            attr_height = sample_height // len(self.attributes)
            
            regions[f"Sample {sample_id}"] = {}
            
            for i, attribute in enumerate(self.attributes):
                # Calculate boundaries exactly like training
                y_start = sample_region["y_start"] + i * attr_height
                y_end = sample_region["y_start"] + (i + 1) * attr_height
                x_start = sample_region["x_start"]
                x_end = sample_region["x_end"]
                
                # EXACT same padding as training extractor
                padding_y = int(attr_height * 0.02)  # Very small Y padding
                padding_x = int((x_end - x_start) * 0.01)  # Minimal X padding
                
                y_start = max(0, y_start - padding_y)
                y_end = min(height, y_end + padding_y)
                x_start = max(0, x_start - padding_x)
                x_end = min(width, x_end + padding_x)
                
                # Extract region
                region_img = processed_image[y_start:y_end, x_start:x_end]
                
                if region_img.size > 0:
                    # EXACT same resizing as training
                    resized = cv2.resize(region_img, self.target_size, interpolation=cv2.INTER_CUBIC)
                    regions[f"Sample {sample_id}"][attribute] = resized
                    
        return regions
    
    def predict_rating(self, region_image):
        """
        Predict the rating (1-9) for a single region image.
        
        This method takes a preprocessed region image and returns the predicted
        rating along with confidence scores. The confidence information helps
        identify predictions that might need human review.
        """
        
        if not self.is_loaded:
            if not self.load_model():
                raise RuntimeError("Model not loaded and failed to load")
                
        # Prepare image for model input
        # Add batch and channel dimensions: (1, height, width, 1)
        input_image = region_image.reshape(1, self.target_size[0], self.target_size[1], 1)
        
        # Normalize pixel values to [0, 1] range (matches training preprocessing)
        input_image = input_image.astype(np.float32) / 255.0
        
        # Get model predictions
        predictions = self.model.predict(input_image, verbose=0)
        
        # Convert to probabilities and extract results
        probabilities = predictions[0]
        predicted_class = np.argmax(probabilities)
        confidence = np.max(probabilities)
        
        # Convert from 0-indexed class to 1-9 rating
        predicted_rating = predicted_class + 1
        
        return predicted_rating, confidence, probabilities
    
    def process_form_image(self, image_path):
        """
        Complete pipeline: process a form image and extract all sensory ratings.
        
        This is the main method that orchestrates the entire ML pipeline,
        from raw phone camera image to structured sensory data that can be
        loaded into your data collection interface.
        """
        
        print(f"Processing form image: {os.path.basename(image_path)}")
        
        # Step 1: Preprocess the image
        try:
            processed_image = self.preprocess_image(image_path)
            print("Image preprocessing completed")
        except Exception as e:
            raise Exception(f"Image preprocessing failed: {e}")
        
        # Step 2: Extract all regions
        try:
            regions = self.extract_regions_for_prediction(processed_image)
            total_regions = sum(len(sample_regions) for sample_regions in regions.values())
            print(f"Extracted {total_regions} regions for prediction")
        except Exception as e:
            raise Exception(f"Region extraction failed: {e}")
        
        # Step 3: Predict ratings for all regions
        extracted_data = {}
        prediction_log = []
        
        for sample_name, sample_regions in regions.items():
            extracted_data[sample_name] = {}
            
            for attribute, region_img in sample_regions.items():
                try:
                    # Get prediction for this region
                    rating, confidence, probabilities = self.predict_rating(region_img)
                    
                    # Store the prediction
                    extracted_data[sample_name][attribute] = rating
                    
                    # Log prediction details for analysis
                    prediction_log.append({
                        'sample': sample_name,
                        'attribute': attribute,
                        'predicted_rating': rating,
                        'confidence': confidence,
                        'probabilities': probabilities.tolist()
                    })
                    
                    print(f"  {sample_name} - {attribute}: Rating {rating} (confidence: {confidence:.3f})")
                    
                except Exception as e:
                    print(f"  Error predicting {sample_name} - {attribute}: {e}")
                    # Use default rating if prediction fails
                    extracted_data[sample_name][attribute] = 5
            
            # Add empty comments field for consistency with manual entry
            extracted_data[sample_name]['comments'] = ''
        
        # Step 4: Generate processing summary
        avg_confidence = np.mean([log['confidence'] for log in prediction_log])
        low_confidence_count = sum(1 for log in prediction_log if log['confidence'] < 0.7)
        
        print(f"Processing complete:")
        print(f"  Average confidence: {avg_confidence:.3f}")
        print(f"  Low confidence predictions: {low_confidence_count}/{len(prediction_log)}")
        
        return extracted_data, processed_image

class MLTrainingHelper:
    """
    Training pipeline for the sensory rating classification CNN.
    
    This class handles model architecture design, training orchestration,
    and performance validation. The architecture is specifically designed
    for the characteristics of your phone camera training data.
    """
    
    def __init__(self, processor):
        self.processor = processor
        self.model = None
        self.target_size = (500,120)
        self.training_history = None
        print("MLTrainingHelper initialized")
        print("Ready to train CNN model from your labeled training data")
    def create_cnn_model(self):
        """
        Build CNN architecture optimized for the new higher resolution.
        """
        tf, keras, layers = _lazy_import_tensorflow()
        
        model = keras.Sequential([
            # Input layer - updated dimensions
            layers.Input(shape=(self.target_size[1], self.target_size[0], 1)),  # (120, 500, 1)
            
            # First conv block - adapted for wider images
            layers.Conv2D(32, (3, 3), activation='relu', padding='same'),
            layers.MaxPooling2D((2, 2)),
            layers.BatchNormalization(),
            
            # Second conv block
            layers.Conv2D(64, (3, 3), activation='relu', padding='same'),
            layers.MaxPooling2D((2, 2)),
            layers.BatchNormalization(),
            
            # Third conv block 
            layers.Conv2D(128, (3, 3), activation='relu', padding='same'),
            layers.MaxPooling2D((2, 2)),
            layers.BatchNormalization(),
            
            # Fourth conv block for higher resolution
            layers.Conv2D(256, (3, 3), activation='relu', padding='same'),
            layers.MaxPooling2D((2, 2)),
            layers.BatchNormalization(),
            
            # Global average pooling instead of flatten for better generalization
            layers.GlobalAveragePooling2D(),
            
            # Dense layers
            layers.Dense(512, activation='relu'),
            layers.Dropout(0.5),
            layers.Dense(256, activation='relu'),
            layers.Dropout(0.3),
            
            # Output layer - 9 classes for ratings 1-9
            layers.Dense(9, activation='softmax')
        ])
        
        # Compile with appropriate optimizer for higher resolution
        model.compile(
            optimizer=keras.optimizers.Adam(learning_rate=0.001),
            loss='categorical_crossentropy',
            metrics=['accuracy', 'top_2_accuracy']
        )
        
        return model
    
     def load_training_data(self):
        """
        Load and prepare training data from your organized directory structure.
        
        This method takes the carefully labeled and augmented training data
        that your training_data_helper.py created and prepares it for CNN training.
        The data loading process includes validation to ensure your training
        data is properly organized and balanced.
        """
        
        # Import TensorFlow components for data loading
        tf, keras, layers = _lazy_import_tensorflow()
        
        train_dir = "training_data/sensory_ratings"
        
        if not os.path.exists(train_dir):
            raise Exception(
                "Training data directory not found. Please run the training data "
                "extraction process first using training_data_helper.py"
            )
        
        # Check that we have data for all rating classes
        missing_classes = []
        class_counts = {}
        
        for rating in range(1, 10):
            class_dir = os.path.join(train_dir, f"rating_{rating}")
            if not os.path.exists(class_dir):
                missing_classes.append(rating)
                class_counts[rating] = 0
            else:
                count = len([f for f in os.listdir(class_dir) if f.endswith('.jpg')])
                class_counts[rating] = count
                if count == 0:
                    missing_classes.append(rating)
        
        if missing_classes:
            raise Exception(
                f"Missing training data for rating classes: {missing_classes}. "
                f"Please ensure you have labeled examples for all ratings 1-9."
            )
        
        total_samples = sum(class_counts.values())
        print(f"Training data summary:")
        for rating, count in class_counts.items():
            percentage = (count / total_samples) * 100
            print(f"  Rating {rating}: {count:4d} images ({percentage:5.1f}%)")
        print(f"  Total: {total_samples:4d} images")
        
        # Configure data generators with appropriate augmentation
        # We use lighter augmentation here since your training data already includes augmented examples
        train_datagen = keras.preprocessing.image.ImageDataGenerator(
            rescale=1./255,  # Normalize pixel values to [0,1]
            validation_split=0.2,  # Reserve 20% for validation
            
            # Light additional augmentation to improve robustness
            rotation_range=3,  # Small rotations to handle minor alignment variations
            width_shift_range=0.05,  # Minor position shifts
            height_shift_range=0.05,
            zoom_range=0.05,  # Slight zoom variations
            horizontal_flip=False,  # Don't flip - rating scales have orientation
            fill_mode='nearest'  # How to fill pixels after transformations
        )
        
        # Create training data generator
        train_generator = train_datagen.flow_from_directory(
            train_dir,
            target_size=self.processor.target_size,  # Resize to model input size
            batch_size=32,  # Process 32 images at a time
            class_mode='categorical',  # One-hot encoded labels for multi-class classification
            color_mode='grayscale',  # Single channel input
            subset='training',  # Use training split
            shuffle=True,  # Randomize order for better training
            seed=42  # Reproducible random seed
        )
        
        # Create validation data generator (no additional augmentation for validation)
        validation_generator = train_datagen.flow_from_directory(
            train_dir,
            target_size=self.processor.target_size,
            batch_size=32,
            class_mode='categorical',
            color_mode='grayscale',
            subset='validation',  # Use validation split
            shuffle=False,  # Keep consistent order for validation
            seed=42
        )
        
        print(f"Data generators created:")
        print(f"  Training samples: {train_generator.samples}")
        print(f"  Validation samples: {validation_generator.samples}")
        print(f"  Classes found: {train_generator.num_classes}")
        print(f"  Class labels: {list(train_generator.class_indices.keys())}")
        
        return train_generator, validation_generator
    
    def train_model(self, epochs=50, save_best_only=True):
        """
        Train the CNN model using your labeled training data.
        
        This method orchestrates the complete training process, transforming
        your human expertise (captured in the labeled training examples) into
        a neural network that can automatically recognize rating patterns.
        
        The training process includes several sophisticated techniques to ensure
        you get the best possible model from your training data, including
        early stopping to prevent overfitting and learning rate scheduling
        to optimize convergence.
        """
        
        print("="*60)
        print("STARTING CNN TRAINING FOR SENSORY RATING CLASSIFICATION")
        print("="*60)
        
        # Load and validate training data
        print("Loading training data...")
        train_gen, val_gen = self.load_training_data()
        
        if train_gen.samples < 100:
            print("WARNING: Small training dataset detected.")
            print("Consider collecting more training examples for better model performance.")
        
        # Create the CNN model architecture
        print("Creating model architecture...")
        self.model = self.create_cnn_model()
        
        # Display model summary for verification
        print("\nModel Architecture Summary:")
        self.model.summary()
        
        # Create models directory if it doesn't exist
        os.makedirs('models', exist_ok=True)
        
        # Configure training callbacks for optimal performance
        callbacks = []
        
        # Model checkpointing - save the best model during training
        if save_best_only:
            checkpoint_callback = keras.callbacks.ModelCheckpoint(
                'models/sensory_rating_classifier_best.h5',
                monitor='val_accuracy',  # Watch validation accuracy
                save_best_only=True,     # Only save when performance improves
                mode='max',              # Higher accuracy is better
                verbose=1,               # Print when saving
                save_weights_only=False  # Save complete model
            )
            callbacks.append(checkpoint_callback)
        
        # Early stopping - prevent overfitting by stopping when validation performance plateaus
        early_stopping = keras.callbacks.EarlyStopping(
            monitor='val_accuracy',   # Watch validation accuracy
            patience=15,              # Wait 15 epochs before stopping
            restore_best_weights=True, # Use the best weights found
            verbose=1,                # Print when stopping early
            min_delta=0.001           # Minimum improvement to count as progress
        )
        callbacks.append(early_stopping)
        
        # Learning rate reduction - automatically tune learning rate during training
        lr_reduction = keras.callbacks.ReduceLROnPlateau(
            monitor='val_loss',       # Watch validation loss
            factor=0.5,               # Reduce learning rate by half
            patience=8,               # Wait 8 epochs before reducing
            min_lr=1e-7,             # Don't go below this learning rate
            verbose=1                 # Print when reducing learning rate
        )
        callbacks.append(lr_reduction)
        
        # Training progress logging
        csv_logger = keras.callbacks.CSVLogger(
            'models/training_log.csv',
            append=True  # Add to existing log if present
        )
        callbacks.append(csv_logger)
        
        # Calculate training parameters
        steps_per_epoch = max(1, train_gen.samples // train_gen.batch_size)
        validation_steps = max(1, val_gen.samples // val_gen.batch_size)
        
        print(f"\nTraining Configuration:")
        print(f"  Epochs: {epochs}")
        print(f"  Steps per epoch: {steps_per_epoch}")
        print(f"  Validation steps: {validation_steps}")
        print(f"  Batch size: {train_gen.batch_size}")
        
        print(f"\nStarting training...")
        
        # Train the model - this is where your human expertise becomes machine intelligence
        self.training_history = self.model.fit(
            train_gen,
            epochs=epochs,
            steps_per_epoch=steps_per_epoch,
            validation_data=val_gen,
            validation_steps=validation_steps,
            callbacks=callbacks,
            verbose=1  # Show progress bar
        )
        
        # Save the final model
        final_model_path = 'models/sensory_rating_classifier.h5'
        self.model.save(final_model_path)
        
        # Analyze training results
        final_train_acc = self.training_history.history['accuracy'][-1]
        final_val_acc = self.training_history.history['val_accuracy'][-1]
        best_val_acc = max(self.training_history.history['val_accuracy'])
        
        print("="*60)
        print("TRAINING COMPLETED!")
        print("="*60)
        print(f"Final training accuracy: {final_train_acc:.4f}")
        print(f"Final validation accuracy: {final_val_acc:.4f}")
        print(f"Best validation accuracy: {best_val_acc:.4f}")
        
        # Check for potential overfitting
        acc_diff = final_train_acc - final_val_acc
        if acc_diff > 0.1:
            print(f"\nWARNING: Possible overfitting detected")
            print(f"Training accuracy ({final_train_acc:.4f}) significantly higher than validation ({final_val_acc:.4f})")
            print("Consider collecting more training data or adjusting regularization.")
        else:
            print(f"\nGood generalization: Training-validation gap = {acc_diff:.4f}")
        
        print(f"\nModel files saved:")
        print(f"  Final model: {final_model_path}")
        if save_best_only:
            print(f"  Best model: models/sensory_rating_classifier_best.h5")
        print(f"  Training log: models/training_log.csv")
        
        return self.model, self.training_history
    
    def create_training_data_structure(self):
        """
        Create the directory structure needed for training data organization.
        
        This method sets up the folder hierarchy that TensorFlow's data generators
        expect, making it easy to organize your labeled training examples by
        rating class for efficient loading during training.
        """
        
        base_dir = "training_data/sensory_ratings"
        os.makedirs(base_dir, exist_ok=True)
        
        # Create a folder for each rating class (1 through 9)
        for rating in range(1, 10):
            rating_dir = os.path.join(base_dir, f"rating_{rating}")
            os.makedirs(rating_dir, exist_ok=True)
            
        # Create additional directories for training management
        os.makedirs("training_data/logs", exist_ok=True)
        os.makedirs("training_data/augmentation_examples", exist_ok=True)
        os.makedirs("models", exist_ok=True)
        os.makedirs("logs/ml_processing", exist_ok=True)
        
        print("Training data structure created:")
        for rating in range(1, 10):
            print(f"  training_data/sensory_ratings/rating_{rating}/")
        print("  training_data/logs/")
        print("  training_data/augmentation_examples/")
        print("  models/")
        print("  logs/ml_processing/")
        
        print(f"\nNext steps:")
        print(f"1. Use training_data_helper.py to extract and label training regions")
        print(f"2. Train the model using the ML menu options")
        print(f"3. Test the trained model on new forms")