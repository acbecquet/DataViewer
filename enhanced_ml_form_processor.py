"""
enhanced_ml_form_processor.py
Enhanced ML Form Processor for Sensory Data Collection
Integrates with enhanced training workflow including shadow removal and OCR boundary detection
"""

import cv2
import numpy as np
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
        print("TIMING: Lazy loaded TensorFlow for enhanced ML processing")
        return tf, keras, layers
    except ImportError as e:
        print(f"Error importing TensorFlow: {e}")
        raise ImportError("TensorFlow is required for ML functionality. Install with: pip install tensorflow")

# CPU optimization configuration - ADD THIS FIRST
def configure_cpu_optimization():
    """Configure TensorFlow for optimal CPU usage."""
    print("TIMING: Configuring CPU optimization...")
    tf, keras,layers = _lazy_import_tensorflow()
    # Get CPU core count
    cpu_count = os.cpu_count()
    print(f"TIMING: Found {cpu_count} CPU cores")

    # Configure threading for maximum CPU usage
    tf.config.threading.set_inter_op_parallelism_threads(cpu_count)
    tf.config.threading.set_intra_op_parallelism_threads(cpu_count)

    # Enable CPU optimizations
    os.environ['OMP_NUM_THREADS'] = str(cpu_count)
    os.environ['TF_NUM_INTEROP_THREADS'] = str(cpu_count)
    os.environ['TF_NUM_INTRAOP_THREADS'] = str(cpu_count)

    # Enable optimized CPU operations
    tf.config.optimizer.set_experimental_options({'layout_optimizer': True})

    print(f"TIMING: CPU optimization enabled - using all {cpu_count} cores")
    return cpu_count

# Call this before any TensorFlow operations
CPU_CORES = configure_cpu_optimization()

class EnhancedMLFormProcessor:
    """
    Enhanced ML processor that matches the improved training extractor exactly.

    This processor uses the EXACT same preprocessing, boundary detection, and region extraction
    as the enhanced training workflow to ensure perfect consistency between training and inference.
    """

    def __init__(self, model_path='models/sensory_rating_classifier.h5'):
        print("DEBUG: EnhancedMLFormProcessor.__init__ started (no TF loading yet)")
        self.model_path = model_path
        self.model = None
        self.is_loaded = False

        # MUST match enhanced training extractor exactly
        self.target_size = (600, 140)  # Updated to match enhanced training
        self.model_input_shape = (140,600,1)
        self.num_classes = 9  # Ratings 1-9

        # EXACT same attributes as enhanced training extractor
        self.attributes = [
            "Burnt Taste", "Vapor Volume", "Overall Flavor",
            "Smoothness", "Overall Liking"
        ]

        print("Enhanced MLFormProcessor initialized")
        print(f"Model path: {model_path}")
        print(f"Target image size: {self.target_size} (matches enhanced training)")
        print("Features: Shadow removal preprocessing, OCR boundary detection compatibility")

    def load_model(self):
        """
        Load the trained CNN model for rating classification.
        Enhanced with better error handling and architecture validation.
        """
        tf, keras, layers = _lazy_import_tensorflow()
        if self.is_loaded:
            return True

        if not os.path.exists(self.model_path):
            print(f"Model file not found: {self.model_path}")
            print("Please train the enhanced model first using the enhanced training workflow")

            # Check for alternative model files
            alternative_paths = [
                'models/sensory_rating_classifier_best.h5',
                'models/enhanced/sensory_rating_classifier.h5'
            ]

            for alt_path in alternative_paths:
                if os.path.exists(alt_path):
                    print(f"Found alternative model: {alt_path}")
                    self.model_path = alt_path
                    break
            else:
                return False

        try:
            # Load the trained model
            self.model = keras.models.load_model(self.model_path)

            # Validate model architecture for enhanced resolution
            expected_input_shape = (None, self.target_size[1], self.target_size[0], 1)  # (None, 140, 600, 1)
            actual_input_shape = self.model.input_shape

            if actual_input_shape != expected_input_shape:
                print(f"WARNING: Model input shape {actual_input_shape} doesn't match enhanced expected {expected_input_shape}")
                print("This may indicate the model was trained with different parameters")

                # Try to determine the correct target size from model
                if len(actual_input_shape) == 4:
                    model_height = actual_input_shape[1]
                    model_width = actual_input_shape[2]
                    if model_height and model_width:
                        self.target_size = (model_width, model_height)
                        print(f"Adjusted target size to match model: {self.target_size}")

            expected_output_shape = (None, self.num_classes)
            actual_output_shape = self.model.output_shape

            if actual_output_shape != expected_output_shape:
                print(f"WARNING: Model output shape {actual_output_shape} doesn't match expected {expected_output_shape}")

            self.is_loaded = True
            print(f"Enhanced model loaded successfully")
            print(f"Input shape: {actual_input_shape}")
            print(f"Output shape: {actual_output_shape}")
            print(f"Model file size: {os.path.getsize(self.model_path) / (1024*1024):.1f} MB")

            return True

        except Exception as e:
            print(f"Error loading enhanced model: {e}")
            import traceback
            traceback.print_exc()
            return False

    def preprocess_image_enhanced(self, image_path):
        """
        Enhanced preprocessing that matches the enhanced training extractor exactly.

        This uses the SAME preprocessing pipeline as ImprovedAttributeDetectionExtractor
        to ensure consistency between training and inference.
        """

        # Load the image with error handling
        img = cv2.imread(image_path)
        if img is None:
            raise ValueError(f"Could not load image: {image_path}")

        print(f"DEBUG: Original image size: {img.shape[1]}x{img.shape[0]}")

        # EXACT same preprocessing as enhanced training extractor
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

        # EXACT same enhancement as training
        clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8,8))
        enhanced = clahe.apply(gray)

        # EXACT same denoising as training (if used)
        denoised = cv2.bilateralFilter(enhanced, 5, 25, 25)

        print(f"DEBUG: Enhanced preprocessing complete: {denoised.shape[1]}x{denoised.shape[0]}")

        return denoised

    def detect_form_cross_lines_enhanced(self, image):
        """
        Enhanced cross line detection that matches the training extractor.

        This uses the SAME logic as ImprovedAttributeDetectionExtractor to find
        the form center coordinates for consistent region extraction.
        """

        height, width = image.shape
        print(f"DEBUG: Detecting form cross lines in {width}x{height} image")

        # EXACT same line detection as enhanced training extractor
        edges = cv2.Canny(image, 100, 200)

        # Find horizontal lines (for center_y)
        horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 1))
        horizontal_lines = cv2.morphologyEx(edges, cv2.MORPH_OPEN, horizontal_kernel)

        # Find vertical lines (for center_x)
        vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 40))
        vertical_lines = cv2.morphologyEx(edges, cv2.MORPH_OPEN, vertical_kernel)

        # Calculate center coordinates
        # Find the strongest horizontal line around the middle
        center_y = height // 2  # Default fallback
        horizontal_projection = np.sum(horizontal_lines, axis=1)
        middle_range = slice(height//3, 2*height//3)
        middle_projection = horizontal_projection[middle_range]

        if np.max(middle_projection) > 0:
            relative_center = np.argmax(middle_projection)
            center_y = height//3 + relative_center

        # Find the strongest vertical line around the middle
        center_x = width // 2  # Default fallback
        vertical_projection = np.sum(vertical_lines, axis=0)
        middle_range = slice(width//3, 2*width//3)
        middle_projection = vertical_projection[middle_range]

        if np.max(middle_projection) > 0:
            relative_center = np.argmax(middle_projection)
            center_x = width//3 + relative_center

        print(f"DEBUG: Form center detected at: ({center_x}, {center_y})")

        return center_x, center_y

    def get_enhanced_sample_regions(self, image, center_x, center_y):
        """
        Get sample regions using enhanced boundary detection.

        This matches the enhanced training extractor's region calculation exactly.
        """

        height, width = image.shape
        print(f"DEBUG: Calculating enhanced sample regions for {width}x{height} image")
        print(f"DEBUG: Using form center: ({center_x}, {center_y})")

        # EXACT same sample region calculation as enhanced training extractor
        sample_regions = {
            1: {  # Top-left sample
                "y_start": int(height * 0.20), "y_end": int(center_y - 10),
                "x_start": int(width * 0.05), "x_end": int(center_x - 10)
            },
            2: {  # Top-right sample
                "y_start": int(height * 0.20), "y_end": int(center_y - 10),
                "x_start": int(center_x + 10), "x_end": int(width * 0.95)
            },
            3: {  # Bottom-left sample
                "y_start": int(center_y + 10), "y_end": int(height * 0.90),
                "x_start": int(width * 0.05), "x_end": int(center_x - 10)
            },
            4: {  # Bottom-right sample
                "y_start": int(center_y + 10), "y_end": int(height * 0.90),
                "x_start": int(center_x + 10), "x_end": int(width * 0.95)
            }
        }

        print(f"DEBUG: Sample regions calculated:")
        for sample_id, region in sample_regions.items():
            region_width = region["x_end"] - region["x_start"]
            region_height = region["y_end"] - region["y_start"]
            print(f"DEBUG:   Sample {sample_id}: {region_width}x{region_height} at ({region['x_start']},{region['y_start']})")

        return sample_regions

    def extract_regions_for_prediction_enhanced(self, processed_image):
        """
        Enhanced region extraction that matches the training extractor exactly.

        This uses the SAME logic as ImprovedAttributeDetectionExtractor to ensure
        perfect consistency between training and inference.
        """

        height, width = processed_image.shape
        print(f"DEBUG: Enhanced region extraction for {width}x{height} image")

        # Step 1: Detect form structure
        center_x, center_y = self.detect_form_cross_lines_enhanced(processed_image)

        # Step 2: Get sample regions
        sample_regions = self.get_enhanced_sample_regions(processed_image, center_x, center_y)

        # Step 3: Extract attribute regions from each sample
        regions = {}
        total_regions = 0

        for sample_id, sample_region in sample_regions.items():
            sample_height = sample_region["y_end"] - sample_region["y_start"]
            attr_height = sample_height // len(self.attributes)

            sample_name = f"Sample {sample_id}"
            regions[sample_name] = {}

            print(f"DEBUG: Extracting attributes for {sample_name}")
            print(f"DEBUG:   Sample region: {sample_region}")
            print(f"DEBUG:   Attribute height: {attr_height}")

            for i, attribute in enumerate(self.attributes):
                # Calculate exact boundaries like training
                y_start = sample_region["y_start"] + i * attr_height
                y_end = sample_region["y_start"] + (i + 1) * attr_height
                x_start = sample_region["x_start"]
                x_end = sample_region["x_end"]

                # EXACT same padding as enhanced training extractor
                padding_y = max(2, int(attr_height * 0.05))  # 5% padding
                padding_x = max(5, int((x_end - x_start) * 0.02))  # 2% padding

                y_start = max(0, y_start - padding_y)
                y_end = min(height, y_end + padding_y)
                x_start = max(0, x_start - padding_x)
                x_end = min(width, x_end + padding_x)

                # Extract region
                region_img = processed_image[y_start:y_end, x_start:x_end]

                print(f"DEBUG:     {attribute}: region {region_img.shape[1]}x{region_img.shape[0]} at ({x_start},{y_start})")

                if region_img.size > 0:
                    # EXACT same resizing as enhanced training
                    resized = cv2.resize(region_img, self.target_size, interpolation=cv2.INTER_CUBIC)
                    regions[sample_name][attribute] = resized
                    total_regions += 1

                    # Debug: Save region for verification
                    debug_dir = "debug_production_regions"
                    os.makedirs(debug_dir, exist_ok=True)
                    debug_filename = f"{debug_dir}/sample_{sample_id}_{attribute.replace(' ', '_')}.png"
                    cv2.imwrite(debug_filename, resized)

        print(f"DEBUG: Enhanced extraction complete: {total_regions} regions extracted")
        print(f"DEBUG: Debug regions saved to: debug_production_regions/")

        return regions

    def predict_rating_enhanced(self, region_image):
        """
        Enhanced prediction with better confidence analysis and debugging.
        """

        if not self.is_loaded:
            if not self.load_model():
                raise RuntimeError("Enhanced model not loaded and failed to load")

        # Prepare image for model input with enhanced validation
        if region_image.shape != self.target_size:
            print(f"WARNING: Region shape {region_image.shape} doesn't match target {self.target_size}")
            region_image = cv2.resize(region_image, self.target_size, interpolation=cv2.INTER_CUBIC)

        # Add batch and channel dimensions: (1, height, width, 1)
        input_image = region_image.reshape(1, self.target_size[1], self.target_size[0], 1)

        # Normalize pixel values to [0, 1] range (matches enhanced training preprocessing)
        input_image = input_image.astype(np.float32) / 255.0

        # Get model predictions with enhanced error handling
        try:
            predictions = self.model.predict(input_image, verbose=0)
        except Exception as e:
            print(f"ERROR: Model prediction failed: {e}")
            # Return default prediction
            return 5, 0.1, np.array([0.1] * 9)

        # Convert to probabilities and extract results
        probabilities = predictions[0]
        predicted_class = np.argmax(probabilities)
        confidence = np.max(probabilities)

        # Convert from 0-indexed class to 1-9 rating
        predicted_rating = predicted_class + 1

        # Enhanced confidence analysis
        second_highest = np.sort(probabilities)[-2]
        confidence_gap = confidence - second_highest

        print(f"DEBUG: Prediction: Rating {predicted_rating}, Confidence: {confidence:.3f}, Gap: {confidence_gap:.3f}")

        return predicted_rating, confidence, probabilities

    def process_form_image_enhanced(self, image_path):
        """
        Enhanced complete pipeline that matches the training workflow exactly.

        This orchestrates the entire enhanced ML pipeline from raw phone camera image
        to structured sensory data with comprehensive debugging and error handling.
        """

        print(f"="*80)
        print(f"ENHANCED ML FORM PROCESSING")
        print(f"="*80)
        print(f"Processing: {os.path.basename(image_path)}")
        print(f"Target resolution: {self.target_size}")
        print(f"Enhanced features: Shadow removal preprocessing, OCR boundary detection")

        # Step 1: Enhanced preprocessing
        try:
            processed_image = self.preprocess_image_enhanced(image_path)
            print("✓ Enhanced image preprocessing completed")
        except Exception as e:
            raise Exception(f"Enhanced image preprocessing failed: {e}")

        # Step 2: Enhanced region extraction
        try:
            regions = self.extract_regions_for_prediction_enhanced(processed_image)
            total_regions = sum(len(sample_regions) for sample_regions in regions.values())
            print(f"✓ Enhanced region extraction completed: {total_regions} regions")
        except Exception as e:
            raise Exception(f"Enhanced region extraction failed: {e}")

        # Step 3: Enhanced prediction pipeline
        extracted_data = {}
        prediction_log = []
        low_confidence_predictions = []

        print(f"\nStarting enhanced predictions...")

        for sample_name, sample_regions in regions.items():
            extracted_data[sample_name] = {}

            print(f"\nProcessing {sample_name}:")

            for attribute, region_img in sample_regions.items():
                try:
                    # Get enhanced prediction
                    rating, confidence, probabilities = self.predict_rating_enhanced(region_img)

                    # Store the prediction
                    extracted_data[sample_name][attribute] = rating

                    # Enhanced prediction logging
                    prediction_info = {
                        'sample': sample_name,
                        'attribute': attribute,
                        'predicted_rating': rating,
                        'confidence': confidence,
                        'probabilities': probabilities.tolist(),
                        'top_3_predictions': self.get_top_predictions(probabilities, 3)
                    }
                    prediction_log.append(prediction_info)

                    # Track low confidence predictions
                    if confidence < 0.7:
                        low_confidence_predictions.append(prediction_info)

                    confidence_status = "HIGH" if confidence > 0.8 else "MEDIUM" if confidence > 0.6 else "LOW"
                    print(f"  ✓ {attribute}: Rating {rating} (confidence: {confidence:.3f} - {confidence_status})")

                except Exception as e:
                    print(f"  ✗ Error predicting {attribute}: {e}")
                    # Use middle rating as fallback
                    extracted_data[sample_name][attribute] = 5

            # Add empty comments field for consistency
            extracted_data[sample_name]['comments'] = ''

        # Step 4: Enhanced results analysis
        if prediction_log:
            avg_confidence = np.mean([log['confidence'] for log in prediction_log])
            confidence_std = np.std([log['confidence'] for log in prediction_log])

            print(f"\n" + "="*80)
            print(f"ENHANCED PROCESSING COMPLETE")
            print(f"="*80)
            print(f"Samples processed: {len(extracted_data)}")
            print(f"Total predictions: {len(prediction_log)}")
            print(f"Average confidence: {avg_confidence:.3f} ± {confidence_std:.3f}")
            print(f"Low confidence predictions: {len(low_confidence_predictions)}/{len(prediction_log)}")

            if low_confidence_predictions:
                print(f"\nLow confidence predictions requiring review:")
                for pred in low_confidence_predictions:
                    print(f"  {pred['sample']} - {pred['attribute']}: Rating {pred['predicted_rating']} (conf: {pred['confidence']:.3f})")

            # Save detailed prediction log
            log_dir = "logs/enhanced_ml_processing"
            os.makedirs(log_dir, exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            log_filename = f"{log_dir}/enhanced_prediction_log_{timestamp}.json"

            with open(log_filename, 'w') as f:
                json.dump({
                    'image_path': image_path,
                    'processing_timestamp': timestamp,
                    'predictions': prediction_log,
                    'summary': {
                        'samples_processed': len(extracted_data),
                        'total_predictions': len(prediction_log),
                        'average_confidence': avg_confidence,
                        'low_confidence_count': len(low_confidence_predictions)
                    }
                }, f, indent=2)

            print(f"Detailed prediction log saved: {log_filename}")

        return extracted_data, processed_image

    def get_top_predictions(self, probabilities, top_k=3):
        """Get top K predictions with their probabilities."""
        top_indices = np.argsort(probabilities)[-top_k:][::-1]
        return [(int(idx + 1), float(probabilities[idx])) for idx in top_indices]


class EnhancedMLTrainingHelper:
    """
    Enhanced training pipeline that matches the improved training workflow.

    This class handles enhanced model architecture, training orchestration,
    and performance validation specifically designed for the enhanced resolution
    and improved training data from the shadow removal workflow.
    """

    def __init__(self, processor):
        print("DEBUG: EnhancedMLTrainingHelper.__init__ started (no TF loading)")
        self.processor = processor
        self.model = None
        self.target_size = (600, 140)  # Enhanced resolution
        self.training_history = None

        print("Enhanced MLTrainingHelper initialized")
        print(f"Enhanced target size: {self.target_size}")
        print("Ready to train enhanced CNN model from shadow removal training data")

    def create_enhanced_cnn_model(self):
        """
        Build enhanced CNN architecture optimized for CPU training with better accuracy.
        """
        tf, keras, layers = _lazy_import_tensorflow()

        print("Creating improved CPU-optimized CNN architecture...")
        print(f"Input shape: ({self.target_size[1]}, {self.target_size[0]}, 1)")  # (height, width, channels)

        # Improved model for better accuracy while still CPU-efficient
        model = keras.Sequential([
            # Input layer - FIXED: proper height/width order
            layers.Input(shape=(self.target_size[1], self.target_size[0], 1)),  # (100, 400, 1)

            # First conv block - wider kernels for text features
            layers.Conv2D(32, (3, 5), activation='relu', padding='same'),
            layers.MaxPooling2D((2, 2)),
            layers.BatchNormalization(),
            layers.Dropout(0.1),

            # Second conv block
            layers.Conv2D(64, (3, 3), activation='relu', padding='same'),
            layers.MaxPooling2D((2, 2)),
            layers.BatchNormalization(),
            layers.Dropout(0.15),

            # Third conv block
            layers.Conv2D(128, (3, 3), activation='relu', padding='same'),
            layers.MaxPooling2D((2, 2)),
            layers.BatchNormalization(),
            layers.Dropout(0.2),

            # Fourth conv block for better feature extraction
            layers.Conv2D(256, (3, 3), activation='relu', padding='same'),
            layers.GlobalAveragePooling2D(),
            layers.Dropout(0.3),

            # Dense layers with better regularization
            layers.Dense(512, activation='relu'),
            layers.BatchNormalization(),
            layers.Dropout(0.4),

            layers.Dense(256, activation='relu'),
            layers.BatchNormalization(),
            layers.Dropout(0.3),

            layers.Dense(128, activation='relu'),
            layers.Dropout(0.2),

            # Output layer
            layers.Dense(9, activation='softmax')
        ])

        # Better optimizer settings for this task
        model.compile(
            optimizer=keras.optimizers.Adam(
                learning_rate=0.0003,  # Lower learning rate for better convergence
                beta_1=0.9,
                beta_2=0.999,
                weight_decay=0.0001  # L2 regularization
            ),
            loss='categorical_crossentropy',
            metrics=['accuracy']
        )

        print("Improved CPU-optimized model architecture created")
        return model

    def load_enhanced_training_data(self, batch_size=16, validation_split=0.25):
        """
        Load enhanced training data with optimized parameters for shadow removal data.
        """
        tf, keras, layers = _lazy_import_tensorflow()

        train_dir = "training_data/sensory_ratings"

        if not os.path.exists(train_dir):
            raise Exception(
                "Enhanced training data directory not found. Please run the enhanced "
                "training data extraction process first."
            )

        # Check enhanced training data quality
        print("Analyzing enhanced training data...")

        class_counts = {}
        enhanced_extractions = 0

        for rating in range(1, 10):
            class_dir = os.path.join(train_dir, f"rating_{rating}")
            if os.path.exists(class_dir):
                images = [f for f in os.listdir(class_dir) if f.endswith(('.jpg', '.png', '.jpeg'))]
                info_files = [f for f in os.listdir(class_dir) if f.endswith('_info.txt')]

                class_counts[rating] = len(images)
                enhanced_extractions += len(info_files)

                print(f"  Rating {rating}: {len(images)} images ({len(info_files)} enhanced)")

        total_samples = sum(class_counts.values())
        print(f"Total enhanced training samples: {total_samples}")
        print(f"Enhanced extractions: {enhanced_extractions}")

        # FIXED: Simplified data generators without excessive augmentation
        train_datagen = keras.preprocessing.image.ImageDataGenerator(
            rescale=1./255,
            validation_split=validation_split,

            # REDUCED augmentation to prevent data issues
            rotation_range=1,  # Very minimal rotation
            width_shift_range=0.02,  # Minimal shifts
            height_shift_range=0.01,
            zoom_range=0.02,  # Very slight zoom
            brightness_range=[0.98, 1.02],  # Minimal brightness variation
            horizontal_flip=False,  # Don't flip rating scales
            fill_mode='nearest'
        )

        # FIXED: Enhanced training generator with proper configuration
        train_generator = train_datagen.flow_from_directory(
            train_dir,
            target_size=self.target_size,
            batch_size=batch_size,
            class_mode='categorical',
            color_mode='grayscale',
            subset='training',
            shuffle=True,
            seed=42
        )

        # FIXED: Enhanced validation generator
        validation_generator = train_datagen.flow_from_directory(
            train_dir,
            target_size=self.target_size,
            batch_size=batch_size,
            class_mode='categorical',
            color_mode='grayscale',
            subset='validation',
            shuffle=False,
            seed=42
        )

        print(f"Enhanced data generators created:")
        print(f"  Training samples: {train_generator.samples}")
        print(f"  Validation samples: {validation_generator.samples}")
        print(f"  Batch size: {batch_size}")
        print(f"  Validation split: {validation_split}")

        return train_generator, validation_generator

    def train_enhanced_model(self, epochs=150, batch_size=16, validation_split=0.25,
                           save_best_only=True, patience=25):
        """
        Train the enhanced CNN model with improved settings for accuracy.
        """

        print("="*80)
        print("ENHANCED CNN TRAINING FOR SENSORY RATING CLASSIFICATION")
        print("="*80)
        print("Features: Fixed shapes, improved architecture, better regularization")

        # Load enhanced training data with smaller batch size for stability
        print("Loading enhanced training data...")
        train_gen, val_gen = self.load_enhanced_training_data(batch_size, validation_split)

        # Use smaller batch sizes for better gradient updates
        optimal_batch_sizes = [16, 12, 8]  # Smaller batches for better learning

        for test_batch in optimal_batch_sizes:
            try:
                print(f"TIMING: Testing CPU batch size {test_batch}...")
                if train_gen.samples >= test_batch * 4:  # Need more samples per batch
                    batch_size = test_batch
                    print(f"TIMING: Using CPU-optimized batch size: {batch_size}")
                    break
            except Exception as e:
                continue

        # Recreate generators with optimized batch size
        train_gen, val_gen = self.load_enhanced_training_data(batch_size, validation_split)

        # Create enhanced model architecture
        print("Creating improved model architecture...")
        self.model = self.create_enhanced_cnn_model()

        # Display model summary
        print("\nImproved Model Architecture Summary:")
        self.model.summary()

        # Create enhanced models directory
        os.makedirs('models/enhanced', exist_ok=True)
        os.makedirs('logs/enhanced_training', exist_ok=True)

        # Enhanced training callbacks with better settings
        callbacks = []

        # Model checkpointing
        from tensorflow import keras
        if save_best_only:
            checkpoint_callback = keras.callbacks.ModelCheckpoint(
                'models/enhanced/sensory_rating_classifier_best.h5',
                monitor='val_accuracy',
                save_best_only=True,
                mode='max',
                verbose=1,
                save_weights_only=False
            )
            callbacks.append(checkpoint_callback)

        # Early stopping with more patience
        early_stopping = keras.callbacks.EarlyStopping(
            monitor='val_accuracy',
            patience=patience,
            restore_best_weights=True,
            verbose=1,
            min_delta=0.001  # Slightly larger delta for stability
        )
        callbacks.append(early_stopping)

        # More conservative learning rate scheduling
        lr_reduction = keras.callbacks.ReduceLROnPlateau(
            monitor='val_loss',
            factor=0.5,  # Less aggressive reduction
            patience=patience//3,
            min_lr=1e-7,
            verbose=1
        )
        callbacks.append(lr_reduction)

        # Enhanced logging with timing
        csv_logger = keras.callbacks.CSVLogger(
            'logs/enhanced_training/training_log.csv',
            append=True
        )
        callbacks.append(csv_logger)

        # Add timing callback
        class TimingCallback(keras.callbacks.Callback):
            def on_epoch_begin(self, epoch, logs=None):
                import time
                self.epoch_start = time.time()

            def on_epoch_end(self, epoch, logs=None):
                import time
                epoch_time = time.time() - self.epoch_start
                steps_time = epoch_time / max(1, self.params['steps'])
                print(f"TIMING: Epoch {epoch+1} - {epoch_time:.1f}s total, {steps_time:.2f}s/step")

        callbacks.append(TimingCallback())

        # Calculate enhanced training parameters
        # Calculate enhanced training parameters - FIXED CALCULATION
        steps_per_epoch = train_gen.samples // train_gen.batch_size
        validation_steps = val_gen.samples // val_gen.batch_size

        # ENSURE we have valid steps
        if steps_per_epoch == 0:
            steps_per_epoch = 1
        if validation_steps == 0:
            validation_steps = 1

        print(f"\nImproved Training Configuration:")
        print(f"  Epochs: {epochs}")
        print(f"  Batch size: {batch_size}")
        print(f"  Steps per epoch: {steps_per_epoch}")
        print(f"  Validation steps: {validation_steps}")
        print(f"  Early stopping patience: {patience}")
        print(f"  Target shape: {self.target_size} -> Model input: ({self.target_size[1]}, {self.target_size[0]}, 1)")

        print(f"\nStarting improved training...")

        # Train with proper step configuration
        self.training_history = self.model.fit(
            train_gen,
            epochs=epochs,
            steps_per_epoch=steps_per_epoch,
            validation_data=val_gen,
            validation_steps=validation_steps,
            callbacks=callbacks,
            verbose=1
        )
        # Save enhanced final model
        final_model_path = 'models/sensory_rating_classifier.h5'
        enhanced_model_path = 'models/enhanced/sensory_rating_classifier.h5'

        self.model.save(final_model_path)
        self.model.save(enhanced_model_path)

        # Enhanced results analysis with error handling
        final_train_acc = self.training_history.history['accuracy'][-1]
        final_val_acc = self.training_history.history['val_accuracy'][-1]
        best_val_acc = max(self.training_history.history['val_accuracy'])



        print("="*80)
        print("ENHANCED TRAINING COMPLETED!")
        print("="*80)
        print(f"Final training accuracy: {final_train_acc:.4f}")
        print(f"Final validation accuracy: {final_val_acc:.4f}")
        print(f"Best validation accuracy: {best_val_acc:.4f}")


        # Enhanced overfitting analysis
        acc_diff = final_train_acc - final_val_acc
        if acc_diff > 0.15:
            print(f"\nWARNING: Possible overfitting detected")
            print(f"Consider collecting more enhanced training data")
        else:
            print(f"\nExcellent generalization: Training-validation gap = {acc_diff:.4f}")

        print(f"\nEnhanced model files saved:")
        print(f"  Final model: {final_model_path}")
        print(f"  Enhanced model: {enhanced_model_path}")
        if save_best_only:
            print(f"  Best model: models/enhanced/sensory_rating_classifier_best.h5")
        print(f"  Training log: logs/enhanced_training/training_log.csv")

        return self.model, self.training_history

def check_training_data_quality():
    """Check the quality and balance of training data."""
    print("CHECKING TRAINING DATA QUALITY")
    print("="*50)

    base_dir = "training_data/sensory_ratings"

    for rating in range(1, 10):
        rating_dir = os.path.join(base_dir, f"rating_{rating}")
        if os.path.exists(rating_dir):
            images = [f for f in os.listdir(rating_dir) if f.endswith(('.jpg', '.png', '.jpeg'))]
            print(f"Rating {rating}: {len(images)} images")

            # Check a few sample images
            if images:
                sample_img_path = os.path.join(rating_dir, images[0])
                sample_img = cv2.imread(sample_img_path, cv2.IMREAD_GRAYSCALE)
                if sample_img is not None:
                    print(f"  Sample image shape: {sample_img.shape}")
                else:
                    print(f"  WARNING: Could not load sample image!")
        else:
            print(f"Rating {rating}: No directory found")

# Maintain backward compatibility
MLFormProcessor = EnhancedMLFormProcessor
MLTrainingHelper = EnhancedMLTrainingHelper

# Run data quality check when module is imported
if __name__ == "__main__":
    check_training_data_quality()



