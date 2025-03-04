# report_generator.py
import os
import shutil
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from datetime import datetime
import processing
from utils import get_resource_path, get_save_path, plotting_sheet_test

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
                            print(f"Skipping sheet '{sheet_name}': No valid 'data' key found.")
                            continue

                        data = sheet_info["data"]  #  Extract the actual DataFrame
                        print(f"Processing sheet: {sheet_name}")
                        
                        is_plotting = plotting_sheet_test(sheet_name, data)
                        process_function = processing.get_processing_function(sheet_name)
                        if is_plotting:
                            processed_data, _, full_sample_data = process_function(data)
                            valid_plot_options = processing.get_valid_plot_options(plot_options, full_sample_data)
                        else:
                            data = data.astype(str).replace([pd.NA], '')
                            processed_data, _, full_sample_data = process_function(data)
                            valid_plot_options = []

                        if processed_data.empty or full_sample_data.empty:
                            print(f"Skipping sheet '{sheet_name}' due to empty processed data.")
                            continue

                        self.write_excel_report(writer, sheet_name, processed_data, full_sample_data,
                                                valid_plot_options, images_to_delete)

                        if is_plotting and valid_plot_options:
                            numeric_data = full_sample_data.apply(pd.to_numeric, errors='coerce')
                            for plot_option in valid_plot_options:
                                plot_image_path = f"{sheet_name}_{plot_option}_plot.png"
                                try:
                                    processing.plot_all_samples(numeric_data, 12, plot_option)
                                    plt.savefig(plot_image_path)
                                    plt.close()
                                    if plot_image_path not in images_to_delete:
                                        images_to_delete.append(plot_image_path)
                                except Exception as plot_error:
                                    print(f"Failed to generate plot {plot_option} for {sheet_name}: {plot_error}")

                    except Exception as e:
                        print(f"Error processing sheet '{sheet_name}': {e}")
                        continue

                    # Phase 1 - excel - 60%
                    processed_count += 1
                    progress = int((processed_count / total_sheets) * 60)
                    self.gui.progress_dialog.update_progress_bar(progress)
                    self.gui.root.update_idletasks()
                    
            
            # PowerPoint phase (remaining 40% of progress)
            try:
                total_slides = len(filtered_sheets)  # Adjust this based on actual slide count
                current_slide = 0
            
                def update_ppt_progress():
                    nonlocal current_slide
                    base_progress = 60  # Excel phase complete
                    ppt_progress = int((current_slide / total_slides) * 40)
                    self.gui.progress_dialog.update_progress_bar(base_progress + ppt_progress)
                    self.gui.root.update_idletasks()

                self.write_powerpoint_report(ppt_save_path, images_to_delete, plot_options, progress_callback=update_ppt_progress)
            
                # Final progress update
                self.gui.progress_dialog.update_progress_bar(100)
                self.gui.root.update_idletasks()

            except Exception as e:
                print(f"Error writing PowerPoint report: {e}")
                raise

        except Exception as e:
            # Delete partial files
            for path in [save_path, ppt_save_path]:
                if path and os.path.exists(path):
                    os.remove(path)
            raise
        finally:
            self.cleanup_images(images_to_delete)
            # Only leave this if you need to ensure cleanup - remove progress update here
            # self.gui.root.update_idletasks()
            

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

            if data is None or data.empty:
                messagebox.showwarning("Warning", f"Sheet '{selected_sheet}' is empty.")
                return

            process_function = processing.get_processing_function(selected_sheet)
            processed_data, _, full_sample_data = process_function(data)
            if processed_data.empty or full_sample_data.empty:
                messagebox.showwarning("Warning", f"Sheet '{selected_sheet}' did not yield valid processed data.")
                return

            with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
                valid_plot_options = processing.get_valid_plot_options(plot_options, full_sample_data)
                self.write_excel_report(writer, selected_sheet, processed_data, full_sample_data,
                                          valid_plot_options, images_to_delete)

            ppt_save_path = save_path.replace('.xlsx', '.pptx')
            self.write_powerpoint_report_for_test(ppt_save_path, images_to_delete, selected_sheet,
                                                  processed_data, full_sample_data, plot_options)

            self.cleanup_images(images_to_delete)
            messagebox.showinfo("Success", f"Test report saved successfully to:\nExcel: {save_path}\nPowerPoint: {ppt_save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while generating the test report: {e}")

    def add_plots_to_slide(self, slide, sheet_name: str, full_sample_data: pd.DataFrame, valid_plot_options: list, images_to_delete: list) -> None:
        plot_top = Inches(1.21)
        left_column_x = Inches(8.43)
        right_column_x = Inches(10.84)
        numeric_data = full_sample_data.apply(pd.to_numeric, errors='coerce')
        if numeric_data.isna().all(axis=0).all():
            print(f"No numeric data available for plotting in sheet '{sheet_name}'.")
            return

        for i, plot_option in enumerate(valid_plot_options):
            plot_image_path = f"{sheet_name}_{plot_option}_plot.png"
            try:
                processing.plot_all_samples(numeric_data, 12, plot_option)
                plt.savefig(plot_image_path)
                plt.close()
                if plot_image_path not in images_to_delete:
                    images_to_delete.append(plot_image_path)
                plot_x = left_column_x if i % 2 == 0 else right_column_x
                if i % 2 != 0:
                    plot_top += Inches(1.83)
                slide.shapes.add_picture(plot_image_path, plot_x, plot_top, Inches(2.29), Inches(1.72))
            except Exception as e:
                print(f"Error generating plot '{plot_option}' for sheet '{sheet_name}': {e}")

    def write_excel_report(self, writer, sheet_name: str, processed_data, full_sample_data, valid_plot_options=[], images_to_delete=None) -> None:
        try:
            processed_data.astype(str).replace([pd.NA], '')
            processed_data.to_excel(writer, sheet_name=sheet_name, index=False)
            plot_sheet_names = processing.get_plot_sheet_names()
            is_plotting = sheet_name in plot_sheet_names
            if is_plotting:
                self.add_plots_to_excel(writer, sheet_name, full_sample_data, images_to_delete, valid_plot_options)
        except Exception as e:
            print(f"Error writing Excel report for sheet '{sheet_name}': {e}")

    def write_powerpoint_report_for_test(self, ppt_save_path: str, images_to_delete: list, sheet_name: str, processed_data, full_sample_data, plot_options: list) -> None:
        try:
            prs = Presentation()
            prs.slide_width = Inches(13.33)
            prs.slide_height = Inches(7.5)
            try:
                slide_layout = prs.slide_layouts[5]
                slide = prs.slides.add_slide(slide_layout)
            except IndexError:
                raise ValueError("Slide layout 5 not found in the PowerPoint template.")

            background_path = get_resource_path("resources/ccell_background.png")
            slide.shapes.add_picture(background_path, left=Inches(0), top=Inches(0),
                                      width=prs.slide_width, height=prs.slide_height)
            logo_path = get_resource_path("resources/ccell_logo_full.png")
            slide.shapes.add_picture(logo_path, left=Inches(11.21), top=Inches(0.43),
                                      width=Inches(1.57), height=Inches(0.53))
            if slide.shapes.title:
                title_shape = slide.shapes.title
            else:
                title_shape = slide.shapes.add_textbox(Inches(0.45), Inches(0.45), Inches(10.72), Inches(0.64))
            text_frame = title_shape.text_frame
            text_frame.margin_top = 0
            text_frame.margin_bottom = 0
            for para in list(text_frame.paragraphs):
                text_frame._element.remove(para._element)
            p = text_frame.add_paragraph()
            p.text = sheet_name
            p.alignment = PP_ALIGN.LEFT
            p.space_before = Pt(0)
            p.space_after = Pt(0)
            text_frame.word_wrap = True
            run = p.runs[0] if p.runs else p.add_run()
            run.font.name = "Montserrat"
            run.font.size = Pt(32)
            run.font.bold = True
            spTree = slide.shapes._spTree
            spTree.remove(title_shape._element)
            spTree.append(title_shape._element)
            plot_sheet_names = processing.get_plot_sheet_names()
            is_plotting = sheet_name in plot_sheet_names
            if is_plotting:
                table_width = Inches(8.07)
                self.add_table_to_slide(slide, processed_data, table_width, is_plotting)
                valid_plot_options = processing.get_valid_plot_options(plot_options, full_sample_data)
                if valid_plot_options:
                    self.add_plots_to_slide(slide, sheet_name, full_sample_data, valid_plot_options, images_to_delete)
                else:
                    print(f"No valid plot options for sheet '{sheet_name}'. Skipping plots.")
            else:
                table_width = Inches(13.03)
                self.add_table_to_slide(slide, full_sample_data, table_width, is_plotting)
            processing.clean_presentation_tables(prs)
            prs.save(ppt_save_path)
            print(f"PowerPoint test report saved successfully at {ppt_save_path}.")
        except Exception as e:
            print(f"Error writing PowerPoint test report: {e}")

    def write_powerpoint_report(self, ppt_save_path: str, images_to_delete: list, plot_options: list, progress_callback = None) -> None:
        try:
            prs = Presentation()
            prs.slide_width = Inches(13.33)
            prs.slide_height = Inches(7.5)
            processed_slides = 0
            total_slides = len(self.gui.filtered_sheets)
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
            background_path = get_resource_path("resources/ccell_background.png")
            logo_path = get_resource_path("resources/ccell_logo_full.png")
            plot_sheet_names = processing.get_plot_sheet_names()

            processed_slides += 1
            self._update_ppt_progress(processed_slides, total_slides)

            for sheet_name in self.gui.filtered_sheets.keys():
                try:
                    data = self.gui.sheets.get(sheet_name)
                    if data is None or data.empty:
                        print(f"Skipping sheet '{sheet_name}': No data available.")
                        continue

                    process_function = processing.get_processing_function(sheet_name)
                    processed_data, _, full_sample_data = process_function(data)

                    if processed_data.empty:
                        print(f"Skipping sheet '{sheet_name}': Processed data is empty.")
                        continue
                    try:
                        slide_layout = prs.slide_layouts[5]
                        slide = prs.slides.add_slide(slide_layout)
                    except IndexError:
                        raise ValueError(f"Slide layout 5 not found in the PowerPoint template.")
                    slide.shapes.add_picture(background_path, left=Inches(0), top=Inches(0),
                                             width=Inches(13.33), height=Inches(7.5))
                    slide.shapes.add_picture(logo_path, left=Inches(11.21), top=Inches(0.43),
                                             width=Inches(1.57), height=Inches(0.53))
                    if slide.shapes.title:
                        title_shape = slide.shapes.title
                        title_shape.left = Inches(0.45)
                        title_shape.top = Inches(0.45)
                        title_shape.width = Inches(10.72)
                        title_shape.height = Inches(0.64)
                        text_frame = title_shape.text_frame
                        text_frame.margin_top = 0
                        text_frame.margin_bottom = 0
                        for para in list(text_frame.paragraphs):
                            text_frame._element.remove(para._element)
                        p = text_frame.add_paragraph()
                        p.text = sheet_name
                        p.alignment = PP_ALIGN.LEFT
                        p.space_before = Pt(0)
                        p.space_after = Pt(0)
                        run = p.runs[0] if p.runs else p.add_run()
                        run.font.name = "Montserrat"
                        run.font.size = Pt(32)
                        run.font.bold = True
                    else:
                        print(f"Warning: No title placeholder found for sheet '{sheet_name}'.")
                    is_plotting = sheet_name in plot_sheet_names
                    if not is_plotting:
                        if processed_data.empty:
                            print(f"Skipping non-plotting sheet '{sheet_name}' due to empty data.")
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
                            print(f"No valid plot options for sheet '{sheet_name}'. Skipping plots.")
                    if slide.shapes.title:
                        title_shape = slide.shapes.title
                        spTree = slide.shapes._spTree
                        spTree.remove(title_shape._element)
                        spTree.append(title_shape._element)
                    processed_slides += 1
                    
                    self._update_ppt_progress(processed_slides, total_slides)
                except Exception as sheet_error:
                    print(f"Error processing sheet '{sheet_name}': {sheet_error}")
                    processed_slides += 1
            processing.clean_presentation_tables(prs)
            prs.save(ppt_save_path)
            print(f"PowerPoint report saved successfully at {ppt_save_path}.")
        except Exception as e:
            print(f"Error writing PowerPoint report: {e}")
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
            if numeric_data.isna().all().all():
                return
            for i, plot_option in enumerate(valid_plot_options):
                plot_image_path = f"{sheet_name}_{plot_option}_plot.png"
                try:
                    processing.plot_all_samples(numeric_data, 12, plot_option)
                    plt.savefig(plot_image_path, dpi=300)
                    plt.close()
                    images_to_delete.append(plot_image_path)
                    col_offset = 10 + (i % 2) * 10
                    row_offset = 2 + (i // 2) * 20
                    worksheet.insert_image(row_offset, col_offset, plot_image_path)
                except Exception as e:
                    print(f"Error generating plot '{plot_option}' for sheet '{sheet_name}': {e}")
        except Exception as e:
            print(f"Error adding plots to Excel for sheet '{sheet_name}': {e}")

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
            print(f"Skipping table creation for slide. Number of columns ({cols}) exceeds 20, number of rows ({rows}) exceeds 30, and it is not a plotting sheet.")
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

