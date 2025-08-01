# report_generator.py
import os
import traceback
import shutil
import pandas as pd
import math
import matplotlib.pyplot as plt
from PIL import Image
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from datetime import datetime
import processing
from utils import get_save_path, plotting_sheet_test, get_plot_sheet_names, debug_print, show_success_message
from resource_utils import get_resource_path
from tkinter import messagebox  # For showing info/errors

class ReportGenerator:
    def __init__(self, gui):
        """
        Initialize the ReportGenerator with a reference to the main Tk window.
        This is used for scheduling GUI callbacks.
        """
        self.gui = gui
        self.root = gui.root

    def generate_full_report(self, filtered_sheets: dict, plot_options: list) -> None:
        """
        Generate a full report (Excel and PowerPoint) for all sheets.
        This function assumes that progress reporting is handled externally.
    
        Args:
            filtered_sheets (dict): A dict mapping sheet names to sheet info.
            plot_options (list): The list of plot options.
        """
        save_path = get_save_path(".xlsx")
        if not save_path:
            raise ValueError("Save cancelled")

        ppt_save_path = save_path.replace('.xlsx', '.pptx')
        images_to_delete = []
        total_sheets = len(filtered_sheets)
        processed_count = 0

        try:
            with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
                for sheet_name, sheet_info in filtered_sheets.items():
                    try:
                        if not isinstance(sheet_info, dict) or "data" not in sheet_info:
                            debug_print(f"DEBUG: Skipping sheet '{sheet_name}': No valid 'data' key found.")
                            continue

                        data = sheet_info["data"]  #  Extract the actual DataFrame
                        debug_print(f"DEBUG: Processing sheet: {sheet_name}")
                        debug_print(f"DEBUG: Sheet data shape: {data.shape}")
                    
                        # Special handling for User Test Simulation
                        if sheet_name == "User Test Simulation":
                            debug_print(f"DEBUG: Processing User Test Simulation with 8-column format")
                            debug_print(f"DEBUG: Data columns: {data.columns.tolist()}")
                            debug_print(f"DEBUG: Data preview:\n{data.head()}")
                    
                        is_plotting = plotting_sheet_test(sheet_name, data)
                        debug_print(f"DEBUG: Sheet {sheet_name} is_plotting: {is_plotting}")
                    
                        process_function = processing.get_processing_function(sheet_name)
                        debug_print(f"DEBUG: Using processing function: {process_function.__name__}")
                    
                        if is_plotting:
                            processed_data, _, full_sample_data = process_function(data)
                            debug_print(f"DEBUG: Processed data shape: {processed_data.shape}")
                            debug_print(f"DEBUG: Full sample data shape: {full_sample_data.shape}")
                            valid_plot_options = processing.get_valid_plot_options(plot_options, full_sample_data)
                            debug_print(f"DEBUG: Valid plot options: {valid_plot_options}")
                        else:
                            data = data.astype(str).replace([pd.NA], '')
                            processed_data, _, full_sample_data = process_function(data)
                            valid_plot_options = []

                        if processed_data.empty or full_sample_data.empty:
                            debug_print(f"DEBUG: Skipping sheet '{sheet_name}' due to empty processed data.")
                            continue

                        self.write_excel_report(writer, sheet_name, processed_data, full_sample_data,
                                                  valid_plot_options, images_to_delete)
                        processed_count += 1
                    
                    except Exception as e:
                        debug_print(f"DEBUG: Error processing sheet '{sheet_name}': {e}")
                        import traceback
                        traceback.print_exc()
                        continue

            try:
                def update_ppt_prog():
                    def callback(percent):
                        total_percent = 50 + (percent * 0.5)  # PPT is second half
                        self.gui.progress_dialog.update_progress_bar(total_percent)
                        self.gui.root.update_idletasks()
                    return callback

                self.write_powerpoint_report(ppt_save_path, images_to_delete, plot_options, progress_callback=update_ppt_prog())
        
                # Final progress update
                self.gui.progress_dialog.update_progress_bar(100)
                self.gui.root.update_idletasks()

            except Exception as e:
                debug_print(f"DEBUG: Error writing PowerPoint report: {e}")
                raise

        except Exception as e:
            # Delete partial files
            for path in [save_path, ppt_save_path]:
                if path and os.path.exists(path):
                    os.remove(path)
            raise
        finally:
            self.cleanup_images(images_to_delete)

    def generate_test_report(self, selected_sheet: str, sheets: dict, plot_options: list) -> None:
        """
        Generate an Excel and PowerPoint report for only the specified sheet.
    
        Args:
            selected_sheet (str): The name of the sheet for the test report.
            sheets (dict): A dict mapping sheet names to raw sheet data.
            plot_options (list): List of plot options.
        """
        save_path = get_save_path(".xlsx")
        if not save_path:
            return

        images_to_delete = []
        try:
            if not selected_sheet:
                messagebox.showerror("Error", "No sheet is selected for the test report.")
                return

            sheet_info = sheets.get(selected_sheet, {})
            data = sheet_info.get("data", pd.DataFrame())  # Ensure it's a DataFrame
        
            debug_print(f"DEBUG: Test report for {selected_sheet}")
            debug_print(f"DEBUG: Data shape: {data.shape}")

            if data is None or data.empty:
                messagebox.showwarning("Warning", f"Sheet '{selected_sheet}' is empty.")
                return

            # Special handling for User Test Simulation
            if selected_sheet == "User Test Simulation":
                debug_print(f"DEBUG: Test report processing User Test Simulation with 8-column format")
                debug_print(f"DEBUG: Data columns: {data.columns.tolist()}")

            process_function = processing.get_processing_function(selected_sheet)
            debug_print(f"DEBUG: Using processing function: {process_function.__name__}")
        
            processed_data, _, full_sample_data = process_function(data)
            debug_print(f"DEBUG: Test report processed data shape: {processed_data.shape}")
            debug_print(f"DEBUG: Test report full sample data shape: {full_sample_data.shape}")
        
            if processed_data.empty or full_sample_data.empty:
                messagebox.showwarning("Warning", f"Sheet '{selected_sheet}' did not yield valid processed data.")
                return

            with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
                valid_plot_options = processing.get_valid_plot_options(plot_options, full_sample_data)
                debug_print(f"DEBUG: Test report valid plot options: {valid_plot_options}")
                self.write_excel_report(writer, selected_sheet, processed_data, full_sample_data,
                                          valid_plot_options, images_to_delete)

            ppt_save_path = save_path.replace('.xlsx', '.pptx')
            self.write_powerpoint_report_for_test(ppt_save_path, images_to_delete, selected_sheet,
                                                  processed_data, full_sample_data, plot_options)

            self.cleanup_images(images_to_delete)
            show_success_message("Success", f"Test report saved successfully to:\nExcel: {save_path}\nPowerPoint: {ppt_save_path}", self.gui.root)
        except Exception as e:
            debug_print(f"DEBUG: Test report generation error: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"An error occurred while generating the test report: {e}")

    def add_plots_to_slide(self, slide, sheet_name: str, full_sample_data: pd.DataFrame, valid_plot_options: list, images_to_delete: list) -> None:
        plot_top = Inches(1.21)
        left_column_x = Inches(8.43)
        right_column_x = Inches(10.84)
        numeric_data = full_sample_data.apply(pd.to_numeric, errors='coerce')
        if numeric_data.isna().all(axis=0).all():
            debug_print(f"No numeric data available for plotting in sheet '{sheet_name}'.")
            return
        sample_names = None
        if hasattr(self, 'header_data') and self.header_data and 'samples' in self.header_data:
            sample_names = [sample['id'] for sample in self.header_data['samples']]
            debug_print(f"DEBUG: Extracted sample names from header_data: {sample_names}")

        # Determine if this is User Test Simulation
        is_user_test_simulation = sheet_name in ["User Test Simulation", "User Simulation Test"]
        num_columns_per_sample = 8 if is_user_test_simulation else 12
    
        for i, plot_option in enumerate(valid_plot_options):
            plot_image_path = f"{sheet_name}_{plot_option}_plot.png"
            try:
                # FIXED: Correct argument order
                fig, sample_names_returned = processing.plot_all_samples(numeric_data, plot_option, num_columns_per_sample, sample_names)
            
                if is_user_test_simulation and hasattr(fig, 'is_split_plot') and fig.is_split_plot:
                    # For User Test Simulation split plots, save with adjusted size and spacing
                    plt.savefig(plot_image_path, dpi=150, bbox_inches='tight')
                    debug_print(f"DEBUG: Saved User Test Simulation split plot: {plot_image_path}")
                
                    # Adjust positioning for split plots (they're wider)
                    plot_x = left_column_x if i % 2 == 0 else right_column_x - Inches(1.0)  # Shift left for wider plots
                    if i % 2 != 0:
                        plot_top += Inches(2.2)  # More vertical spacing for taller plots
                    
                    # Add with larger dimensions for split plots
                    slide.shapes.add_picture(plot_image_path, plot_x, plot_top, Inches(3.5), Inches(2.0))
                else:
                    # Standard single plots
                    plt.savefig(plot_image_path, dpi=150)
                    debug_print(f"DEBUG: Saved standard plot: {plot_image_path}")
                
                    plot_x = left_column_x if i % 2 == 0 else right_column_x
                    if i % 2 != 0:
                        plot_top += Inches(1.83)
                    slide.shapes.add_picture(plot_image_path, plot_x, plot_top, Inches(2.29), Inches(1.72))
                
                plt.close()
                if plot_image_path not in images_to_delete:
                    images_to_delete.append(plot_image_path)
            except Exception as e:
                print(f"Error generating plot '{plot_option}' for sheet '{sheet_name}': {e}")
                import traceback
                traceback.print_exc()

    def write_excel_report(self, writer, sheet_name: str, processed_data, full_sample_data, valid_plot_options=[], images_to_delete=None) -> None:
        try:
            processed_data.astype(str).replace([pd.NA], '')
            processed_data.to_excel(writer, sheet_name=sheet_name, index=False)
            plot_sheet_names = get_plot_sheet_names()
            is_plotting = sheet_name in plot_sheet_names
            if is_plotting:
                self.add_plots_to_excel(writer, sheet_name, full_sample_data, images_to_delete, valid_plot_options)
        except Exception as e:
            print(f"Error writing Excel report for sheet '{sheet_name}': {e}")  
            traceback.print_exc()

    def write_powerpoint_report_for_test(self, ppt_save_path: str, images_to_delete: list, sheet_name: str, processed_data, full_sample_data, plot_options: list) -> None:
        try:
            prs = Presentation()
            prs.slide_width = Inches(13.33)
            prs.slide_height = Inches(7.5)

            current_file = self.gui.current_file
            image_paths = self.gui.sheet_images.get(current_file, {}).get(sheet_name, [])

            # Create main content slide
            main_slide = prs.slides.add_slide(prs.slide_layouts[6])
        
            # Add background and logo
            background_path = get_resource_path("resources/ccell_background.png")
            main_slide.shapes.add_picture(background_path, Inches(0), Inches(0),
                                          width=prs.slide_width, height=prs.slide_height)
            logo_path = get_resource_path("resources/ccell_logo_full.png")
            main_slide.shapes.add_picture(logo_path, Inches(11.21), Inches(0.43),
                                          width=Inches(1.57), height=Inches(0.53))

            # Add title
            title_shape = main_slide.shapes.add_textbox(Inches(0.45), Inches(-0.04), 
                                                       Inches(10.72), Inches(0.64))
            text_frame = title_shape.text_frame

            text_frame.clear()

            p = text_frame.add_paragraph()
            p.text = sheet_name
            p.font.name = "Montserrat"
            p.font.size = Pt(32)
            p.font.bold = True

            # Add table/plots
            plot_sheet_names = get_plot_sheet_names()
            is_plotting = sheet_name in plot_sheet_names

            # For test reports, ALWAYS use processed data for tables
            table_width = Inches(8.07) if is_plotting else Inches(13.03)
            self.add_table_to_slide(main_slide, 
                                   processed_data,  # <-- ALWAYS use processed_data for test reports
                                   table_width, 
                                   is_plotting)

            # Only add plots if it's a plotting sheet AND we have valid plot options
            if is_plotting:
                valid_plot_options = processing.get_valid_plot_options(plot_options, full_sample_data)
                if valid_plot_options:
                    self.add_plots_to_slide(main_slide, sheet_name, full_sample_data,
                                           valid_plot_options, images_to_delete)

            # Create image slide if needed
            if image_paths:
                img_slide = prs.slides.add_slide(prs.slide_layouts[6])
                self.setup_image_slide(prs, img_slide, sheet_name)
                self.add_images_to_slide(img_slide, image_paths)

            processing.clean_presentation_tables(prs)
            prs.save(ppt_save_path)
            debug_print(f"PowerPoint test report saved to {ppt_save_path}")
        
        except Exception as e:
            print(f"Error generating test PowerPoint: {e}")
            if os.path.exists(ppt_save_path):
                os.remove(ppt_save_path)
            raise
            traceback.print_exc()

    def write_powerpoint_report(self, ppt_save_path: str, images_to_delete: list, plot_options: list, progress_callback = None) -> None:
        try:
            processed_slides = 0
            prs = Presentation()
            prs.slide_width = Inches(13.33)
            prs.slide_height = Inches(7.5)

            total_slides = 1  # Cover slide
            for sheet_name in self.gui.filtered_sheets.keys():
                total_slides += 1  # Main slide
                if sheet_name in self.gui.sheet_images and self.gui.sheet_images[self.gui.current_file][sheet_name]:
                    total_slides += 1  # Image slide


            cover_slide = prs.slides.add_slide(prs.slide_layouts[6])
            bg_path = get_resource_path("resources/Cover_Page_Logo.jpg")
            cover_slide.shapes.add_picture(bg_path, left=Inches(0), top=Inches(0),
                                             width=prs.slide_width, height=prs.slide_height)
            textbox_title = cover_slide.shapes.add_textbox(
                left=Inches((prs.slide_width.inches - 12) / 2),
                top=Inches(2.35),
                width=Inches(12),
                height=Inches(0.88)
            )
            text_frame = textbox_title.text_frame
            text_frame.margin_top = 0
            text_frame.margin_bottom = 0
            for para in list(text_frame.paragraphs):
                text_frame._element.remove(para._element)
            p = text_frame.add_paragraph()
            if self.gui.file_path:
                p.text = f"{os.path.splitext(os.path.basename(self.gui.file_path))[0]} Standard Test Report"
            else:
                p.text = "Standard Test Report"
            p.alignment = PP_ALIGN.CENTER
            p.space_before = Pt(0)
            p.space_after = Pt(0)
            text_frame.word_wrap = True
            run = p.runs[0] if p.runs else p.add_run()
            run.font.name = "Montserrat"
            run.font.size = Pt(46)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            cover_slide.shapes._spTree.remove(textbox_title._element)
            cover_slide.shapes._spTree.append(textbox_title._element)
            textbox_sub = cover_slide.shapes.add_textbox(
                left=Inches(5.73),
                top=Inches(4.05),
                width=Inches(1.87),
                height=Inches(0.37)
            )
            sub_frame = textbox_sub.text_frame
            sub_frame.margin_top = 0
            sub_frame.margin_bottom = 0
            for para in list(sub_frame.paragraphs):
                sub_frame._element.remove(para._element)
            p2 = sub_frame.add_paragraph()
            p2.text = datetime.today().strftime("%d %B %Y")
            p2.alignment = PP_ALIGN.CENTER
            p2.space_before = Pt(0)
            p2.space_after = Pt(0)
            sub_frame.word_wrap = False
            run2 = p2.runs[0] if p2.runs else p2.add_run()
            run2.font.name = "Montserrat"
            run2.font.size = Pt(16)
            run2.font.bold = True
            run2.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            cover_slide.shapes._spTree.remove(textbox_sub._element)
            cover_slide.shapes._spTree.append(textbox_sub._element)
            logo_path = get_resource_path("resources/ccell_logo_full_white.png")
            cover_slide.shapes.add_picture(logo_path, left=Inches(5.7), top=Inches(6.12),
                                             width=Inches(1.93), height=Inches(0.66))
            
            # start setting up main slides

            background_path = get_resource_path("resources/ccell_background.png")
            logo_path = get_resource_path("resources/ccell_logo_full.png")
            plot_sheet_names = get_plot_sheet_names()

            processed_slides += 1
            self._update_ppt_progress(processed_slides, total_slides)

            for sheet_name in self.gui.filtered_sheets.keys():
                try:
                    data = self.gui.sheets.get(sheet_name)
                    if data is None or data.empty:
                        # Try to get data from filtered_sheets instead
                        sheet_info = self.gui.filtered_sheets.get(sheet_name)
                        if sheet_info and "data" in sheet_info:
                            data = sheet_info["data"]
                        else:
                            debug_print(f"Skipping sheet '{sheet_name}': No data available.")
                            continue

                    process_function = processing.get_processing_function(sheet_name)
                    processed_data, _, full_sample_data = process_function(data)

                    if processed_data.empty:
                        debug_print(f"Skipping sheet '{sheet_name}': Processed data is empty.")
                        continue
                    try:
                        slide_layout = prs.slide_layouts[6]
                        slide = prs.slides.add_slide(slide_layout)
                    except IndexError:
                        raise ValueError(f"Slide layout 5 not found in the PowerPoint template.")
                    slide.shapes.add_picture(background_path, left=Inches(0), top=Inches(0),
                                             width=Inches(13.33), height=Inches(7.5))
                    slide.shapes.add_picture(logo_path, left=Inches(11.21), top=Inches(0.43),
                                             width=Inches(1.57), height=Inches(0.53))
                    
                    title_shape = slide.shapes.add_textbox(Inches(0.45), Inches(-0.04), Inches(10.72), Inches(0.64)) # hard coded
                    text_frame = title_shape.text_frame
                    text_frame.clear()
                    p = text_frame.add_paragraph()
                    p.text = sheet_name
                    p.alignment = PP_ALIGN.LEFT
                    run = p.runs[0]
                    run.font.name = "Montserrat"
                    run.font.size = Pt(32)
                    run.font.bold = True

                    is_plotting = sheet_name in plot_sheet_names
                    if not is_plotting:
                        if processed_data.empty:
                            debug_print(f"Skipping non-plotting sheet '{sheet_name}' due to empty data.")
                            continue
                        table_width = Inches(13.03)
                        self.add_table_to_slide(slide, full_sample_data, table_width, is_plotting)
                    else:
                        table_width = Inches(8.07)
                        self.add_table_to_slide(slide, processed_data, table_width, is_plotting)
                        valid_plot_options = processing.get_valid_plot_options(plot_options, full_sample_data)
                        if valid_plot_options:
                            self.add_plots_to_slide(slide, sheet_name, full_sample_data, valid_plot_options, images_to_delete)
                        else:
                            debug_print(f"No valid plot options for sheet '{sheet_name}'. Skipping plots.")
                    if slide.shapes.title:
                        title_shape = slide.shapes.title
                        spTree = slide.shapes._spTree
                        spTree.remove(title_shape._element)
                        spTree.append(title_shape._element)

                    
                    processed_slides += 1

                    if progress_callback:
                        progress_callback(processed_slides, total_slides)
                    
                    self._update_ppt_progress(processed_slides, total_slides)

                    # Add image slide if images exist

                    if self.gui.current_file in self.gui.sheet_images and sheet_name in self.gui.sheet_images.get(self.gui.current_file, {}):
                        debug_print("Images Exist! Adding a slide...")
                        current_file = self.gui.current_file
                        image_paths = self.gui.sheet_images.get(current_file, {}).get(sheet_name, [])
                        valid_image_paths = [path for path in image_paths if os.path.exists(path)]
                        img_slide = prs.slides.add_slide(prs.slide_layouts[6])
                        self.setup_image_slide(prs, img_slide, sheet_name)
                        self.add_images_to_slide(img_slide, valid_image_paths)

                        # Update progress after image slide
                        processed_slides += 1
                        self._update_ppt_progress(processed_slides, total_slides)


                except Exception as sheet_error:
                    debug_print(f"Error processing sheet '{sheet_name}': {sheet_error}")
                    processed_slides += 1
                    traceback.print_exc()
                    continue
            processing.clean_presentation_tables(prs)
            prs.save(ppt_save_path)
            debug_print(f"PowerPoint report saved successfully at {ppt_save_path}.")
        except Exception as e:
            print(f"Error writing PowerPoint report: {e}")
            traceback.print_exc()
              # Skip this sheet and continue with others
        finally:
            for image_path in set(images_to_delete):
                if os.path.exists(image_path):
                    try:
                        os.remove(image_path)
                    except OSError as cleanup_error:
                        print(f"Error deleting image {image_path}: {cleanup_error}")

    def _update_ppt_progress(self, processed: int, total: int):
        """Update progress for PowerPoint phase (40% of total)"""
        base_progress = 60  # Excel phase already completed
        ppt_progress = int((processed / total) * 40)  # 40% allocated to PPT
        self.gui.progress_dialog.update_progress_bar(base_progress + ppt_progress)
        self.gui.root.update_idletasks()

    def add_plots_to_excel(self, writer, sheet_name: str, full_sample_data, images_to_delete: list, valid_plot_options: list) -> None:
        try:
            worksheet = writer.sheets[sheet_name]
            numeric_data = full_sample_data.apply(pd.to_numeric, errors='coerce')
            sample_names = None
            if hasattr(self, 'header_data') and self.header_data and 'samples' in self.header_data:
                sample_names = [sample['id'] for sample in self.header_data['samples']]
                debug_print(f"DEBUG: Extracted sample names from header_data: {sample_names}")
            if numeric_data.isna().all().all():
                return
            
            # Determine if this is User Test Simulation
            is_user_test_simulation = sheet_name in ["User Test Simulation", "User Simulation Test"]
            num_columns_per_sample = int(8 if is_user_test_simulation else 12)
        
            for i, plot_option in enumerate(valid_plot_options):
                plot_image_path = f"{sheet_name}_{plot_option}_plot.png"
                try:
                    # Correct argument order
                    fig, sample_names_returned = processing.plot_all_samples(numeric_data, plot_option, num_columns_per_sample, sample_names)
                
                    if is_user_test_simulation and hasattr(fig, 'is_split_plot') and fig.is_split_plot:
                        # For User Test Simulation split plots, save with higher DPI and better format
                        plt.savefig(plot_image_path, dpi=200, bbox_inches='tight')
                        debug_print(f"DEBUG: Saved User Test Simulation split plot for Excel: {plot_image_path}")
                    
                        # Adjust Excel positioning for wider split plots
                        col_offset = 10 + (i % 2) * 15  # More spacing for wider plots
                        row_offset = 2 + (i // 2) * 25  # More vertical spacing
                    else:
                        # Standard plots
                        plt.savefig(plot_image_path, dpi=300)
                        debug_print(f"DEBUG: Saved standard plot for Excel: {plot_image_path}")
                    
                        col_offset = 10 + (i % 2) * 10
                        row_offset = 2 + (i // 2) * 20
                
                    plt.close()
                    images_to_delete.append(plot_image_path)
                    worksheet.insert_image(row_offset, col_offset, plot_image_path)
                
                except Exception as e:
                    print(f"Error generating plot '{plot_option}' for sheet '{sheet_name}': {e}")
                    import traceback
                    traceback.print_exc()
        except Exception as e:
            print(f"Error adding plots to Excel for sheet '{sheet_name}': {e}")
            import traceback
            traceback.print_exc()

    def add_logo_to_slide(self, slide) -> None:
        logo_path = get_resource_path('resources/ccell_logo_full.png')
        left = Inches(7.75)
        top = Inches(0.43)
        height = Inches(0.53)
        width = Inches(1.57)
        if os.path.exists(logo_path):
            slide.shapes.add_picture(logo_path, left, top, width, height)
        else:
            print(f"Logo image not found at {logo_path}")

    def cleanup_images(self, images_to_delete: list) -> None:
        for image_path in set(images_to_delete):
            try:
                os.remove(image_path)
            except FileNotFoundError:
                # The file was already removed or never created; that's fine.
                pass
            except OSError as e:
                print(f"Error deleting file {image_path}: {e}")

    def add_table_to_slide(self, slide, processed_data, table_width, is_plotting) -> bool:
        processed_data = processed_data.fillna(' ')
        processed_data = processed_data.astype(str)
        rows, cols = processed_data.shape
        if cols > 20 and rows > 30 and not is_plotting:
            debug_print(f"Skipping table creation for slide. Number of columns ({cols}) exceeds 20, number of rows ({rows}) exceeds 30, and it is not a plotting sheet.")
            return
        table_left = Inches(0.15)
        table_top = Inches(1.19)
        max_table_height = Inches(5.97)
        table_shape = slide.shapes.add_table(rows+1, cols, table_left, table_top, table_width,
                                              Inches(0.4 * (rows + 1))).table
        header_font = {'name': 'Arial', 'size': Pt(10), 'bold': True}
        cell_font = {'name': 'Arial', 'size': Pt(8), 'bold': False}
        for col_idx, col_name in enumerate(processed_data.columns):
            cell = table_shape.cell(0, col_idx)
            cell.text = str(col_name)
            for paragraph in cell.text_frame.paragraphs:
                paragraph.font.name = header_font['name']
                paragraph.font.size = header_font['size']
                paragraph.font.bold = header_font['bold']
                paragraph.alignment = PP_ALIGN.CENTER
        for row_idx in range(rows):
            for col_idx in range(cols):
                cell = table_shape.cell(row_idx + 1, col_idx)
                cell.text = processed_data.iat[row_idx, col_idx]
                for paragraph in cell.text_frame.paragraphs:
                    paragraph.font.name = cell_font['name']
                    paragraph.font.size = cell_font['size']
                    paragraph.font.bold = cell_font['bold']
                    paragraph.alignment = PP_ALIGN.CENTER
        return True

    def add_images_to_slide(self, slide, image_paths: list) -> None:
        """Adds images to the slide in a grid layout."""
        start_left = Inches(0.05)
        start_top = Inches(1.2)
        available_width = Inches(13.18)  # 13.33 - 0.15
        available_height = Inches(6.3)   # 7.5 - 1.2

        num_images = len(image_paths)
        if num_images == 0:
            return

        # Calculate grid dimensions to best fit available space
        cols = math.ceil(math.sqrt(num_images * (available_width / available_height)))
        rows = math.ceil(num_images / cols)

        # Adjust to prevent excessive columns/rows
        cols = min(cols, 4)  # Max 4 columns for better layout
        rows = math.ceil(num_images / cols)

        # Calculate cell dimensions with margins
        horizontal_margin = Inches(0.2)
        vertical_margin = Inches(0.2)
    
        cell_width = (available_width - (cols - 1) * horizontal_margin) / cols
        cell_height = (available_height - (rows - 1) * vertical_margin) / rows

        # Position each image in grid
        for i, img_path in enumerate(image_paths):
            # Get original image dimensions
            with Image.open(img_path) as img:
                img_width, img_height = img.size
                aspect_ratio = img_width / img_height

            # Calculate display dimensions
            max_width = cell_width
            max_height = cell_height
        
            # Determine which dimension to constrain
            if (cell_width / cell_height) > aspect_ratio:
                # Constrain by height
                display_height = min(cell_height, max_height)
                display_width = display_height * aspect_ratio
            else:
                # Constrain by width
                display_width = min(cell_width, max_width)
                display_height = display_width / aspect_ratio

            # Center the image in the cell
            row = i // cols
            col = i % cols
            x_offset = (cell_width - display_width) / 2
            y_offset = (cell_height - display_height) / 2
        
            left = start_left + col * (cell_width + horizontal_margin) + x_offset
            top = start_top + row * (cell_height + vertical_margin) + y_offset

            slide.shapes.add_picture(img_path, left=left, top=top, width=display_width, height = display_height)

    def setup_image_slide(self, prs, slide, sheet_name):
        """Sets up the background, logo, and title for an image slide."""
        # Add background
        background_path = get_resource_path("resources/ccell_background.png")
        slide.shapes.add_picture(background_path, Inches(0), Inches(0), 
                                 width=prs.slide_width, height=prs.slide_height)
    
        # Add logo
        logo_path = get_resource_path("resources/ccell_logo_full.png")
        slide.shapes.add_picture(logo_path, Inches(11.21), Inches(0.43), 
                                 width=Inches(1.57), height=Inches(0.53))
    
        # Add title
        title_shape = slide.shapes.add_textbox(Inches(0.45), Inches(-0.04), Inches(10.72), Inches(0.64))
        text_frame = title_shape.text_frame
        text_frame.clear()
        p = text_frame.add_paragraph()
        p.text = sheet_name
        p.alignment = PP_ALIGN.LEFT
        run = p.runs[0]
        run.font.name = "Montserrat"
        run.font.size = Pt(32)
        run.font.bold = True