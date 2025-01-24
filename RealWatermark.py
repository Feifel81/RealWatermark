import os
import time
import fitz  # PyMuPDF
import pytesseract
from pathlib import Path
from PySide6.QtWidgets import (QApplication, QMainWindow, QFileDialog, QVBoxLayout, QPushButton, QLineEdit,
                               QLabel, QProgressBar, QSpinBox, QColorDialog, QComboBox, QWidget, QMessageBox, QCheckBox)
from PySide6.QtCore import QThread, Signal, Qt  # Added Qt here
from PySide6.QtGui import QColor, QPixmap, QPainter
from pdf2image import convert_from_path
from PIL import Image, ImageDraw, ImageFont, ImageColor
import subprocess


class PDFProcessor(QThread):
    progress = Signal(int)
    completed = Signal()
    error = Signal(str)

    def __init__(self, input_dirs, output_dir, watermark_text=None, watermark_image=None, transparency=50,
                 position="center", font_color="#FFFFFF", ocr_enabled=False, ocr_language="eng", compress_enabled=False, dpi=150):
        super().__init__()
        self.input_dirs = input_dirs
        self.output_dir = output_dir
        self.watermark_text = watermark_text
        self.watermark_image = watermark_image
        self.transparency = transparency
        self.position = position
        self.font_color = font_color
        self.ocr_enabled = ocr_enabled
        self.ocr_language = ocr_language
        self.compress_enabled = compress_enabled
        self.dpi = dpi
        self.is_running = True
        self.is_paused = False

    def run(self):
        try:
            pdf_files = []
            for directory in self.input_dirs:
                for root, _, files in os.walk(directory):
                    for file in files:
                        if file.endswith(".pdf"):
                            pdf_files.append(os.path.join(root, file))

            total_files = len(pdf_files)
            processed_files = 0

            for pdf_file in pdf_files:
                if not self.is_running:
                    break

                # Pause functionality
                while self.is_paused:
                    time.sleep(0.1)

                relative_path = os.path.relpath(pdf_file, self.input_dirs[0])
                output_pdf_path = os.path.join(self.output_dir, relative_path)

                os.makedirs(os.path.dirname(output_pdf_path), exist_ok=True)

                # Convert PDF to images with specified DPI
                images = convert_from_path(pdf_file, dpi=self.dpi)
                processed_images = []

                for index, image in enumerate(images):
                    if not self.is_running:
                        break

                    # Pause functionality
                    while self.is_paused:
                        time.sleep(0.1)

                    # Apply watermark
                    if self.watermark_text or self.watermark_image:
                        image = self.add_watermark(image)

                    # Convert to RGB to standardize color modes
                    image = image.convert("RGB")
                    processed_images.append(image)

                # Save the processed images as PDF
                if processed_images:
                    temp_pdf_path = output_pdf_path.replace(".pdf", "_temp.pdf")
                    processed_images[0].save(temp_pdf_path, save_all=True, append_images=processed_images[1:], quality=85)

                    # Apply OCR if enabled
                    if self.ocr_enabled:
                        self.apply_ocr(temp_pdf_path, output_pdf_path)
                        if os.path.exists(temp_pdf_path):  # Ensure the file exists before deleting
                            os.remove(temp_pdf_path)  # Delete the temporary file after OCR
                    else:
                        os.replace(temp_pdf_path, output_pdf_path)

                    # Compress the PDF if enabled
                    if self.compress_enabled:
                        self.compress_pdf(output_pdf_path)

                processed_files += 1
                self.progress.emit(int((processed_files / total_files) * 100))

            self.completed.emit()
        except Exception as e:
            self.error.emit(str(e))

    def add_watermark(self, image):
        """
        Adds a text and/or image watermark to the image.
        """
        # Convert the image to RGBA if necessary
        image = image.convert("RGBA")
        base_layer = Image.new("RGBA", image.size, (255, 255, 255, 0))  # Transparent layer
        draw = ImageDraw.Draw(base_layer)

        # Add text watermark
        if self.watermark_text:
            page_width, page_height = image.size

            # Calculate font size based on the position
            if self.position == "diagonal":
                # For diagonal text, calculate font size to fit the diagonal of the page
                diagonal_length = (page_width ** 2 + page_height ** 2) ** 0.5
                font_size = int(diagonal_length / len(self.watermark_text))  # Adjust based on text length

                # Ensure the font size is not too small or too large
                font_size = max(50, min(font_size, 500))  # Set a reasonable range for font size

                # Load the font
                font = ImageFont.truetype("arial.ttf", font_size)

                # Calculate text size
                text_bbox = draw.textbbox((0, 0), self.watermark_text, font=font)
                text_width = text_bbox[2] - text_bbox[0]
                text_height = text_bbox[3] - text_bbox[1]

                # Create a new layer for the rotated text
                watermark_layer = Image.new("RGBA", (page_width * 2, page_height * 2), (255, 255, 255, 0))
                draw_layer = ImageDraw.Draw(watermark_layer)

                # Calculate the starting position (bottom-left corner)
                start_x = (watermark_layer.width - text_width) // 2
                start_y = (watermark_layer.height - text_height) // 2

                # Draw the text on the watermark layer
                draw_layer.text((start_x, start_y), self.watermark_text, font=font,
                                fill=(*ImageColor.getrgb(self.font_color), int(255 * (self.transparency / 100))))

                # Rotate the watermark layer by 45 degrees
                watermark_layer = watermark_layer.rotate(45, expand=True)

                # Calculate the offset to align the rotated text from bottom-left to top-right
                offset_x = (watermark_layer.width - page_width) // 2
                offset_y = (watermark_layer.height - page_height) // 2

                # Paste the rotated text onto the base layer
                base_layer.paste(watermark_layer, (-offset_x, -offset_y), watermark_layer)
            else:
                # For other positions, calculate font size to fit the width of the page
                font_size = int(page_width / len(self.watermark_text))  # Adjust based on text length

                # Ensure the font size is not too small or too large
                font_size = max(10, min(font_size, 200))  # Set a reasonable range for font size

                # Load the font
                font = ImageFont.truetype("arial.ttf", font_size)

                # Calculate text size
                text_bbox = draw.textbbox((0, 0), self.watermark_text, font=font)
                text_width = text_bbox[2] - text_bbox[0]
                text_height = text_bbox[3] - text_bbox[1]

                # Calculate position based on the selected alignment
                position = self.calculate_position(image.size, (text_width, text_height))
                draw.text(position, self.watermark_text, font=font,
                          fill=(*ImageColor.getrgb(self.font_color), int(255 * (self.transparency / 100))))

        # Add image watermark
        if self.watermark_image:
            watermark = Image.open(self.watermark_image).convert("RGBA")

            # Scale the image to fit the page while maintaining aspect ratio
            page_width, page_height = image.size
            watermark_width, watermark_height = watermark.size

            # Calculate the scaling factor to fit the image within the page
            scale_factor = min(page_width / watermark_width, page_height / watermark_height)
            new_width = int(watermark_width * scale_factor)
            new_height = int(watermark_height * scale_factor)

            # Resize the watermark image
            watermark = watermark.resize((new_width, new_height), Image.Resampling.LANCZOS)

            # Apply transparency
            alpha = watermark.split()[3]
            alpha = alpha.point(lambda p: p * (self.transparency / 100))
            watermark.putalpha(alpha)

            # Calculate position to center the image on the page
            position = (
                (page_width - new_width) // 2,
                (page_height - new_height) // 2
            )

            # Paste the watermark onto the base layer
            base_layer.paste(watermark, position, watermark)

        # Combine the layers and convert back to RGB
        return Image.alpha_composite(image, base_layer).convert("RGB")

    def calculate_position(self, image_size, watermark_size):
        """
        Calculates the position for the watermark based on the specified setting.
        """
        image_width, image_height = image_size
        watermark_width, watermark_height = watermark_size

        if self.position == "center":
            return ((image_width - watermark_width) // 2, (image_height - watermark_height) // 2)
        elif self.position == "top-left":
            return (10, 10)
        elif self.position == "top-right":
            return (image_width - watermark_width - 10, 10)
        elif self.position == "bottom-left":
            return (10, image_height - watermark_height - 10)
        elif self.position == "bottom-right":
            return (image_width - watermark_width - 10, image_height - watermark_height - 10)
        return (0, 0)

    def apply_ocr(self, input_pdf_path, output_pdf_path):
        """
        Applies OCR to the PDF using ocrmypdf.
        """
        try:
            # Run ocrmypdf command
            command = ["ocrmypdf", "-l", self.ocr_language, input_pdf_path, output_pdf_path]
            subprocess.run(command, check=True)
        except subprocess.CalledProcessError as e:
            raise Exception(f"OCR failed: {e}")

    def compress_pdf(self, pdf_path):
        """
        Compresses the PDF file.
        """
        # Create a temporary file for compressed output
        temp_compress_path = pdf_path.replace(".pdf", "_compressed.pdf")

        doc = fitz.open(pdf_path)
        doc.save(temp_compress_path, deflate=True, garbage=4)
        doc.close()

        # Replace the original file with the compressed file
        os.replace(temp_compress_path, pdf_path)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("REAL PDF WATERMARK")
        self.setGeometry(300, 100, 800, 600)
        layout = QVBoxLayout()

        # Add Logo
        self.logo_label = QLabel(self)
        logo_pixmap = QPixmap("Logo.jpeg")  # Ensure the logo file is in the same directory
        self.logo_label.setPixmap(logo_pixmap.scaled(800, 600, Qt.KeepAspectRatio, Qt.SmoothTransformation))  # Resize the logo
        self.logo_label.setAlignment(Qt.AlignCenter)
        self.logo_label.setStyleSheet("background-color: rgba(255, 255, 255, 0.5);")  # Set transparency
        layout.addWidget(self.logo_label)

        # Input and output folder selection
        self.input_dirs_btn = QPushButton("Select Folders")
        self.input_dirs_btn.clicked.connect(self.select_input_dirs)
        layout.addWidget(self.input_dirs_btn)

        self.selected_dirs_label = QLabel("No folders selected")
        layout.addWidget(self.selected_dirs_label)

        self.output_dir_btn = QPushButton("Select Output Folder")
        self.output_dir_btn.clicked.connect(self.select_output_dir)
        layout.addWidget(self.output_dir_btn)

        self.selected_output_dir_label = QLabel("No output folder selected")
        layout.addWidget(self.selected_output_dir_label)

        # Watermark text and image
        self.watermark_text_input = QLineEdit()
        self.watermark_text_input.setPlaceholderText("Enter watermark text (optional)")
        layout.addWidget(self.watermark_text_input)

        self.watermark_image_btn = QPushButton("Upload Watermark Image")
        self.watermark_image_btn.clicked.connect(self.select_watermark_image)
        layout.addWidget(self.watermark_image_btn)

        self.selected_watermark_image_label = QLabel("No image selected")
        layout.addWidget(self.selected_watermark_image_label)

        # Transparency and position
        self.transparency_spinner = QSpinBox()
        self.transparency_spinner.setValue(50)
        layout.addWidget(QLabel("Transparency (%)"))
        layout.addWidget(self.transparency_spinner)

        self.position_combo = QComboBox()
        self.position_combo.addItems(["Center", "Top Left", "Top Right", "Bottom Left", "Bottom Right", "Diagonal"])
        layout.addWidget(QLabel("Watermark Position"))
        layout.addWidget(self.position_combo)

        # Font color
        self.color_picker_btn = QPushButton("Select Color")
        self.color_picker_btn.clicked.connect(self.select_font_color)
        layout.addWidget(self.color_picker_btn)

        # OCR options
        self.ocr_checkbox = QCheckBox("Enable OCR")
        layout.addWidget(self.ocr_checkbox)

        self.ocr_language_combo = QComboBox()
        self.ocr_language_combo.addItems(["eng", "deu", "fra", "spa", "chi_sim"])  # Add more languages as needed
        layout.addWidget(QLabel("OCR Language"))
        layout.addWidget(self.ocr_language_combo)

        # Compression options
        self.compress_checkbox = QCheckBox("Compress PDF")
        layout.addWidget(self.compress_checkbox)

        # DPI settings
        self.dpi_combo = QComboBox()
        self.dpi_combo.addItems(["75", "100", "150", "200", "250", "300"])
        self.dpi_combo.setCurrentText("150")
        layout.addWidget(QLabel("DPI Setting"))
        layout.addWidget(self.dpi_combo)

        # Progress bar
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)

        # Start, Abort, and Pause buttons
        self.start_btn = QPushButton("Start")
        self.start_btn.clicked.connect(self.start_processing)
        layout.addWidget(self.start_btn)

        self.abort_btn = QPushButton("Abort")
        self.abort_btn.clicked.connect(self.abort_processing)
        self.abort_btn.setEnabled(False)
        layout.addWidget(self.abort_btn)

        self.pause_btn = QPushButton("Pause")
        self.pause_btn.clicked.connect(self.toggle_pause)
        self.pause_btn.setEnabled(False)
        layout.addWidget(self.pause_btn)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        self.input_dirs = []
        self.output_dir = ""
        self.watermark_image = None
        self.font_color = "#FFFFFF"
        self.processor = None
        self.is_paused = False

    def select_input_dirs(self):
        dialog = QFileDialog(self, "Select Folders")
        dialog.setFileMode(QFileDialog.Directory)
        dialog.setOption(QFileDialog.ReadOnly, True)
        dialog.setOption(QFileDialog.ShowDirsOnly, True)
        if dialog.exec():
            dirs = dialog.selectedFiles()
            if dirs:
                self.input_dirs = dirs
                self.selected_dirs_label.setText("Selected Folders: " + ", ".join(self.input_dirs))

    def select_output_dir(self):
        output_dir = QFileDialog.getExistingDirectory(self, "Select Output Folder", "")
        if output_dir:
            self.output_dir = output_dir
            self.selected_output_dir_label.setText("Output Folder: " + self.output_dir)

    def select_watermark_image(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Select Watermark Image", "", "Images (*.png *.jpg *.jpeg)")
        if file_name:
            self.watermark_image = file_name
            self.selected_watermark_image_label.setText(f"Selected Image: {os.path.basename(file_name)}")

    def select_font_color(self):
        color = QColorDialog.getColor()
        if color.isValid():
            self.font_color = color.name()

    def start_processing(self):
        if not self.input_dirs or not self.output_dir:
            QMessageBox.warning(self, "Error", "Please select input and output folders.")
            return

        # Disable Start button and enable Abort/Pause buttons
        self.start_btn.setEnabled(False)
        self.abort_btn.setEnabled(True)
        self.pause_btn.setEnabled(True)

        self.processor = PDFProcessor(
            input_dirs=self.input_dirs,
            output_dir=self.output_dir,
            watermark_text=self.watermark_text_input.text(),
            watermark_image=self.watermark_image,
            transparency=self.transparency_spinner.value(),
            position=self.position_combo.currentText().lower().replace(" ", "-"),
            font_color=self.font_color,
            ocr_enabled=self.ocr_checkbox.isChecked(),
            ocr_language=self.ocr_language_combo.currentText(),
            compress_enabled=self.compress_checkbox.isChecked(),
            dpi=int(self.dpi_combo.currentText())
        )
        self.processor.progress.connect(self.update_progress)
        self.processor.completed.connect(self.processing_complete)
        self.processor.error.connect(self.show_error)
        self.processor.start()

    def abort_processing(self):
        """Abort the processing."""
        if self.processor:
            self.processor.stop()
            self.abort_btn.setEnabled(False)
            self.pause_btn.setEnabled(False)
            self.start_btn.setEnabled(True)
            QMessageBox.information(self, "Aborted", "Processing has been aborted.")

    def toggle_pause(self):
        """Toggle between pause and resume."""
        if self.processor:
            if self.is_paused:
                self.processor.resume()
                self.pause_btn.setText("Pause")
                self.is_paused = False
            else:
                self.processor.pause()
                self.pause_btn.setText("Resume")
                self.is_paused = True

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def processing_complete(self):
        QMessageBox.information(self, "Complete", "Processing completed!")
        # Re-enable Start button and disable Abort/Pause buttons
        self.start_btn.setEnabled(True)
        self.abort_btn.setEnabled(False)
        self.pause_btn.setEnabled(False)

    def show_error(self, message):
        QMessageBox.critical(self, "Error", message)
        # Re-enable Start button and disable Abort/Pause buttons
        self.start_btn.setEnabled(True)
        self.abort_btn.setEnabled(False)
        self.pause_btn.setEnabled(False)


app = QApplication([])
window = MainWindow()
window.show()
app.exec()
