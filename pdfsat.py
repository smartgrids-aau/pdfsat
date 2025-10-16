#!/usr/bin/env python3
"""
PDF Show and Tell - Dual Screen Presentation Tool
A professional presentation tool for displaying PDF slides with presenter view

Wilfried Elmenreich
Version 0.1, October 2025
Released under the WTFPL (Do What The Fuck You Want To Public License)
"""


import sys
import os
from datetime import datetime, timedelta
from pathlib import Path
import configparser

# Check for required dependencies
try:
    import fitz  # PyMuPDF
except ImportError:
    print("\n" + "="*60)
    print("ERROR: PyMuPDF (fitz) is not installed.")
    print("\nYou need to install PyMuPDF, PyQt6, and screeninfo.")
    print("Please enter these commands:\n")
    commands = "pip install PyMuPDF PyQt6 screeninfo"
    print(f"  {commands}\n")
    print("="*60)
    
    # Try to copy to clipboard
    try:
        import pyperclip
        pyperclip.copy(commands)
        print("The commands have been copied to your clipboard.\n")
    except:
        pass
    
    sys.exit(1)

try:
    from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                                  QHBoxLayout, QLabel, QPushButton, QFileDialog,
                                  QTextEdit, QFrame, QMessageBox)
    from PyQt6.QtCore import Qt, QTimer, QSize, QPoint, pyqtSignal, QPointF
    from PyQt6.QtGui import QPixmap, QImage, QKeySequence, QShortcut, QScreen, QPainter, QColor, QPen, QRadialGradient, QMouseEvent
except ImportError:
    print("\n" + "="*60)
    print("ERROR: PyQt6 is not installed.")
    print("\nYou need to install PyMuPDF, PyQt6, and screeninfo.")
    print("Please enter these commands:\n")
    commands = "pip install PyMuPDF PyQt6 screeninfo"
    print(f"  {commands}\n")
    print("="*60)
    
    # Try to copy to clipboard
    try:
        import pyperclip
        pyperclip.copy(commands)
        print("The commands have been copied to your clipboard.\n")
    except:
        pass
    
    sys.exit(1)

try:
    from screeninfo import get_monitors
except ImportError:
    print("\n" + "="*60)
    print("ERROR: screeninfo is not installed.")
    print("\nYou need to install PyMuPDF, PyQt6, and screeninfo.")
    print("Please enter these commands:\n")
    commands = "pip install PyMuPDF PyQt6 screeninfo"
    print(f"  {commands}\n")
    print("="*60)
    
    # Try to copy to clipboard
    try:
        import pyperclip
        pyperclip.copy(commands)
        print("The commands have been copied to your clipboard.\n")
    except:
        pass
    
    sys.exit(1)


class Config:
    """Handle configuration persistence"""
    def __init__(self):
        self.config_file = Path.home() / '.pdfsat.ini'
        self.config = configparser.ConfigParser()
        self.load()
    
    def load(self):
        if self.config_file.exists():
            self.config.read(self.config_file)
    
    def save(self):
        with open(self.config_file, 'w') as f:
            self.config.write(f)
    
    def get(self, section, key, fallback=None):
        return self.config.get(section, key, fallback=fallback)
    
    def set(self, section, key, value):
        if not self.config.has_section(section):
            self.config.add_section(section)
        self.config.set(section, key, str(value))


class SlideCache:
    """Cache rendered slides for performance"""
    def __init__(self, pdf_path, dpi=150):
        self.doc = fitz.open(pdf_path)
        self.dpi = dpi
        self.cache = {}
        self.total_slides = len(self.doc)
    
    def get_slide(self, page_num):
        """Get slide as QPixmap, cache if not already cached"""
        if page_num not in self.cache:
            page = self.doc[page_num]
            mat = fitz.Matrix(self.dpi / 72, self.dpi / 72)
            pix = page.get_pixmap(matrix=mat)
            
            # Convert to QImage
            img_data = pix.samples
            qimg = QImage(img_data, pix.width, pix.height, 
                         pix.stride, QImage.Format.Format_RGB888)
            self.cache[page_num] = QPixmap.fromImage(qimg)
        
        return self.cache[page_num]
    
    def close(self):
        self.doc.close()
        self.cache.clear()


class NotesLoader:
    """Load and parse notes from text file"""
    def __init__(self, notes_path):
        self.notes = {}
        if notes_path and Path(notes_path).exists():
            self._parse_notes(notes_path)
    
    def _parse_notes(self, notes_path):
        with open(notes_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Split by slide markers
        parts = content.split('---')
        current_slide = 0
        
        for part in parts:
            part = part.strip()
            if not part:
                continue
            
            # Check for explicit slide number marker --12--
            import re
            match = re.match(r'^--(\d+)--\s*', part)
            if match:
                current_slide = int(match.group(1)) - 1  # 0-indexed
                part = part[match.end():].strip()
            
            self.notes[current_slide] = part
            current_slide += 1
    
    def get_notes(self, slide_num):
        return self.notes.get(slide_num, "")


class FullScreenWindow(QMainWindow):
    """Fullscreen window for audience display"""
    def __init__(self, presenter_window=None):
        super().__init__()
        self.presenter_window = presenter_window
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setStyleSheet("background-color: black;")
        
        # Central widget
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # Slide display
        self.slide_label = QLabel()
        self.slide_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.slide_label.setStyleSheet("background-color: black;")
        layout.addWidget(self.slide_label)
        
        self.is_blanked = False
        self.current_pixmap = None
        self.pointer_pos = None  # Position for laser pointer
    
    def show_slide(self, pixmap):
        """Display a slide"""
        self.current_pixmap = pixmap
        if self.is_blanked:
            self.slide_label.clear()
            return
        
        self._update_display()
    
    def _update_display(self):
        """Update the display with current slide and pointer"""
        if not self.current_pixmap or self.is_blanked:
            return
        
        # Scale to fit screen while maintaining aspect ratio
        scaled = self.current_pixmap.scaled(self.size(), Qt.AspectRatioMode.KeepAspectRatio, 
                                           Qt.TransformationMode.SmoothTransformation)
        
        # If pointer is active, draw it on the pixmap
        if self.pointer_pos:
            # Create a copy to draw on
            display_pixmap = QPixmap(scaled.size())
            display_pixmap.fill(Qt.GlobalColor.black)
            
            painter = QPainter(display_pixmap)
            painter.drawPixmap(0, 0, scaled)
            
            # Draw laser pointer effect
            self._draw_laser_pointer(painter, self.pointer_pos)
            
            painter.end()
            self.slide_label.setPixmap(display_pixmap)
        else:
            self.slide_label.setPixmap(scaled)
    
    def _draw_laser_pointer(self, painter, pos):
        """Draw a laser pointer at the given position"""
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # Create radial gradient for glow effect
        gradient = QRadialGradient(QPointF(pos.x(), pos.y()), 30)
        gradient.setColorAt(0, QColor(255, 0, 0, 200))
        gradient.setColorAt(0.3, QColor(255, 50, 0, 150))
        gradient.setColorAt(0.6, QColor(255, 100, 0, 80))
        gradient.setColorAt(1, QColor(255, 150, 0, 0))
        
        # Draw outer glow (sun rays effect)
        painter.setBrush(gradient)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawEllipse(QPointF(pos.x(), pos.y()), 30, 30)
        
        # Draw bright center dot
        center_gradient = QRadialGradient(QPointF(pos.x(), pos.y()), 15)
        center_gradient.setColorAt(0, QColor(255, 255, 255, 255))
        center_gradient.setColorAt(0.5, QColor(255, 0, 0, 255))
        center_gradient.setColorAt(1, QColor(200, 0, 0, 200))
        
        painter.setBrush(center_gradient)
        painter.drawEllipse(QPointF(pos.x(), pos.y()), 15, 15)
    
    def set_pointer_position(self, pos):
        """Set laser pointer position (None to hide)"""
        self.pointer_pos = pos
        self._update_display()
    
    def blank(self):
        """Blank the screen"""
        self.is_blanked = True
        self.slide_label.clear()
    
    def unblank(self):
        """Unblank the screen"""
        self.is_blanked = False
        self._update_display()
    
    def keyPressEvent(self, event):
        """Forward key events to presenter window"""
        if self.presenter_window:
            # Forward the event to the presenter window
            if event.key() == Qt.Key.Key_Right or event.key() == Qt.Key.Key_Space:
                self.presenter_window.next_slide()
            elif event.key() == Qt.Key.Key_Left:
                self.presenter_window.prev_slide()
            elif event.key() == Qt.Key.Key_Escape:
                self.presenter_window.stop_presenting()
            elif event.key() == Qt.Key.Key_B:
                self.presenter_window.toggle_blank()
        else:
            super().keyPressEvent(event)


class PresenterWindow(QMainWindow):
    """Main presenter window with controls"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Show and Tell - Presenter View")
        self.resize(1200, 800)
        
        self.config = Config()
        self.slide_cache = None
        self.notes_loader = None
        self.fullscreen_window = None
        self.current_slide = 0
        self.is_presenting = False
        self.is_blanked = False
        self.start_time = None
        
        # Timer for clock
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000)
        
        self._init_ui()
        self._load_last_session()
        
        # Move to primary screen after initialization
        QTimer.singleShot(100, self._move_to_primary_screen)
    
    def _move_to_primary_screen(self):
        """Move window to primary screen"""
        screens = QApplication.screens()
        if screens:
            primary_screen = screens[0]
            geometry = primary_screen.availableGeometry()
            # Center the window on primary screen
            self.move(geometry.center() - self.rect().center())
    
    def _init_ui(self):
        """Initialize the user interface"""
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        
        # Top control bar with file name and time/duration
        top_layout = QHBoxLayout()
        
        self.file_label = QLabel("No PDF loaded")
        self.file_label.setStyleSheet("font-weight: bold;")
        top_layout.addWidget(self.file_label)
        
        top_layout.addStretch()
        
        # Time and duration with larger font
        time_frame = QFrame()
        time_layout = QHBoxLayout(time_frame)
        time_layout.setContentsMargins(10, 5, 10, 5)
        
        self.time_label = QLabel("Time: --:--:--")
        self.time_label.setStyleSheet("font-size: 24pt; font-weight: bold; color: #0066cc;")
        time_layout.addWidget(self.time_label)
        
        time_layout.addSpacing(20)
        
        self.duration_label = QLabel("Duration: --:--")
        self.duration_label.setStyleSheet("font-size: 24pt; font-weight: bold; color: #cc0000;")
        time_layout.addWidget(self.duration_label)
        
        top_layout.addWidget(time_frame)
        
        main_layout.addLayout(top_layout)
        
        # File control bar
        control_layout = QHBoxLayout()
        
        self.open_btn = QPushButton("Open PDF")
        self.open_btn.clicked.connect(self.open_pdf)
        control_layout.addWidget(self.open_btn)
        
        self.notes_btn = QPushButton("Load Notes")
        self.notes_btn.clicked.connect(self.load_notes)
        control_layout.addWidget(self.notes_btn)
        
        control_layout.addStretch()
        
        main_layout.addLayout(control_layout)
        
        # Presentation controls
        pres_control_layout = QHBoxLayout()
        
        self.start_btn = QPushButton("Start from Beginning [F5]")
        self.start_btn.clicked.connect(self.start_from_beginning)
        self.start_btn.setEnabled(False)
        pres_control_layout.addWidget(self.start_btn)
        
        self.start_current_btn = QPushButton("Start from Current [Shift+F5]")
        self.start_current_btn.clicked.connect(self.start_from_current)
        self.start_current_btn.setEnabled(False)
        pres_control_layout.addWidget(self.start_current_btn)
        
        self.stop_btn = QPushButton("Stop Presenting [Esc]")
        self.stop_btn.clicked.connect(self.stop_presenting)
        self.stop_btn.setEnabled(False)
        pres_control_layout.addWidget(self.stop_btn)
        
        self.blank_btn = QPushButton("(Un)blank Screen [B]")
        self.blank_btn.clicked.connect(self.toggle_blank)
        self.blank_btn.setEnabled(False)
        pres_control_layout.addWidget(self.blank_btn)
        
        pres_control_layout.addStretch()
        
        main_layout.addLayout(pres_control_layout)
        
        # Slide display area
        slides_layout = QHBoxLayout()
        
        # Current slide (left, larger)
        current_frame = QFrame()
        current_frame.setFrameStyle(QFrame.Shape.Box)
        current_layout = QVBoxLayout(current_frame)
        current_label_text = QLabel("Current Slide")
        current_label_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        current_layout.addWidget(current_label_text)
        
        self.current_slide_label = QLabel()
        self.current_slide_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.current_slide_label.setMinimumSize(600, 450)
        self.current_slide_label.setStyleSheet("background-color: white;")
        self.current_slide_label.setMouseTracking(True)
        self.current_slide_label.mouseMoveEvent = self._handle_mouse_move
        
        # Install event filter to track mouse leaving
        self.current_slide_label.installEventFilter(self)
        
        current_layout.addWidget(self.current_slide_label)
        
        slides_layout.addWidget(current_frame, 2)
        
        # Next slide (right, smaller)
        next_frame = QFrame()
        next_frame.setFrameStyle(QFrame.Shape.Box)
        next_layout = QVBoxLayout(next_frame)
        next_label_text = QLabel("Next Slide")
        next_label_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        next_layout.addWidget(next_label_text)
        
        self.next_slide_label = QLabel()
        self.next_slide_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.next_slide_label.setMinimumSize(300, 225)
        self.next_slide_label.setStyleSheet("background-color: white;")
        next_layout.addWidget(self.next_slide_label)
        next_layout.addStretch()
        
        slides_layout.addWidget(next_frame, 1)
        
        main_layout.addLayout(slides_layout, 2)
        
        # Navigation controls
        nav_layout = QHBoxLayout()
        
        self.prev_btn = QPushButton("Previous [←]")
        self.prev_btn.clicked.connect(self.prev_slide)
        self.prev_btn.setEnabled(False)
        nav_layout.addWidget(self.prev_btn)
        
        self.slide_info_label = QLabel("Slide 0 / 0")
        self.slide_info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        nav_layout.addWidget(self.slide_info_label)
        
        self.next_btn = QPushButton("Next [→ / Space]")
        self.next_btn.clicked.connect(self.next_slide)
        self.next_btn.setEnabled(False)
        nav_layout.addWidget(self.next_btn)
        
        main_layout.addLayout(nav_layout)
        
        # Notes area (below)
        notes_frame = QFrame()
        notes_frame.setFrameStyle(QFrame.Shape.Box)
        notes_layout = QVBoxLayout(notes_frame)
        
        notes_header = QLabel("Speaker Notes")
        notes_header.setStyleSheet("font-weight: bold;")
        notes_layout.addWidget(notes_header)
        
        self.notes_text = QTextEdit()
        self.notes_text.setReadOnly(True)
        self.notes_text.setMaximumHeight(150)
        notes_layout.addWidget(self.notes_text)
        
        main_layout.addWidget(notes_frame)
        
        # Keyboard shortcuts
        self._setup_shortcuts()
    
    def eventFilter(self, obj, event):
        """Event filter to detect mouse leaving the slide label"""
        if obj == self.current_slide_label and event.type() == event.Type.Leave:
            if self.is_presenting and self.fullscreen_window:
                self.fullscreen_window.set_pointer_position(None)
        return super().eventFilter(obj, event)
    
    def _handle_mouse_move(self, event):
        """Handle mouse movement over current slide"""
        if not self.is_presenting or not self.fullscreen_window:
            return
        
        # Get mouse position relative to the label
        pos = event.pos()
        
        # Get the pixmap and its position within the label
        pixmap = self.current_slide_label.pixmap()
        if not pixmap:
            return
        
        # Calculate the actual position of the scaled pixmap within the label
        label_size = self.current_slide_label.size()
        pixmap_size = pixmap.size()
        
        # Calculate offsets (centering)
        x_offset = (label_size.width() - pixmap_size.width()) / 2
        y_offset = (label_size.height() - pixmap_size.height()) / 2
        
        # Check if mouse is within the pixmap area
        if (pos.x() < x_offset or pos.x() > x_offset + pixmap_size.width() or
            pos.y() < y_offset or pos.y() > y_offset + pixmap_size.height()):
            self.fullscreen_window.set_pointer_position(None)
            return
        
        # Calculate relative position within the pixmap (0-1 range)
        rel_x = (pos.x() - x_offset) / pixmap_size.width()
        rel_y = (pos.y() - y_offset) / pixmap_size.height()
        
        # Get the fullscreen window's slide label size
        fullscreen_pixmap = self.fullscreen_window.slide_label.pixmap()
        if fullscreen_pixmap:
            # Calculate position on fullscreen display
            fs_x = rel_x * fullscreen_pixmap.width()
            fs_y = rel_y * fullscreen_pixmap.height()
            
            self.fullscreen_window.set_pointer_position(QPoint(int(fs_x), int(fs_y)))
    
    def _setup_shortcuts(self):
        """Setup keyboard shortcuts"""
        # Navigation
        QShortcut(QKeySequence(Qt.Key.Key_Right), self, self.next_slide)
        QShortcut(QKeySequence(Qt.Key.Key_Left), self, self.prev_slide)
        QShortcut(QKeySequence(Qt.Key.Key_Space), self, self.next_slide)
        
        # Presentation control
        QShortcut(QKeySequence(Qt.Key.Key_F5), self, self.start_from_beginning)
        QShortcut(QKeySequence(Qt.Modifier.SHIFT | Qt.Key.Key_F5), self, self.start_from_current)
        QShortcut(QKeySequence(Qt.Key.Key_Escape), self, self.stop_presenting)
        QShortcut(QKeySequence(Qt.Key.Key_B), self, self.toggle_blank)
    
    def keyPressEvent(self, event):
        """Handle key press events"""
        # Navigation keys
        if event.key() == Qt.Key.Key_Right or event.key() == Qt.Key.Key_Space:
            self.next_slide()
        elif event.key() == Qt.Key.Key_Left:
            self.prev_slide()
        elif event.key() == Qt.Key.Key_Escape:
            self.stop_presenting()
        elif event.key() == Qt.Key.Key_B:
            self.toggle_blank()
        elif event.key() == Qt.Key.Key_F5:
            if event.modifiers() & Qt.KeyboardModifier.ShiftModifier:
                self.start_from_current()
            else:
                self.start_from_beginning()
        else:
            super().keyPressEvent(event)
    
    def open_pdf(self):
        """Open a PDF file"""
        last_dir = self.config.get('Session', 'last_directory', fallback=str(Path.home()))
        
        filename, _ = QFileDialog.getOpenFileName(
            self, "Open PDF", last_dir, "PDF Files (*.pdf)"
        )
        
        if filename:
            self._load_pdf(filename)
    
    def _load_pdf(self, pdf_path):
        """Load PDF and initialize cache"""
        try:
            # Close existing cache
            if self.slide_cache:
                self.slide_cache.close()
            
            # Load new PDF
            self.slide_cache = SlideCache(pdf_path)
            
            # Save directory
            self.config.set('Session', 'last_directory', str(Path(pdf_path).parent))
            self.config.set('Session', 'last_file', pdf_path)
            
            # Try to load notes automatically
            notes_path = Path(pdf_path).with_name(
                Path(pdf_path).stem + '_notes.txt'
            )
            if notes_path.exists():
                self.notes_loader = NotesLoader(str(notes_path))
            else:
                self.notes_loader = NotesLoader(None)
            
            # Reset to first slide
            last_slide = int(self.config.get('Session', 'last_slide', fallback='0'))
            self.current_slide = min(last_slide, self.slide_cache.total_slides - 1)
            
            # Update UI
            self.file_label.setText(Path(pdf_path).name)
            self.start_btn.setEnabled(True)
            self.start_current_btn.setEnabled(True)
            self.prev_btn.setEnabled(True)
            self.next_btn.setEnabled(True)
            
            self.update_slides()
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load PDF: {str(e)}")
    
    def load_notes(self):
        """Load a custom notes file"""
        if not self.slide_cache:
            QMessageBox.warning(self, "Warning", "Please load a PDF first")
            return
        
        last_dir = self.config.get('Session', 'last_directory', fallback=str(Path.home()))
        
        filename, _ = QFileDialog.getOpenFileName(
            self, "Open Notes File", last_dir, "Text Files (*.txt)"
        )
        
        if filename:
            self.notes_loader = NotesLoader(filename)
            self.update_slides()
    
    def update_slides(self):
        """Update slide display"""
        if not self.slide_cache:
            return
        
        # Current slide
        current_pixmap = self.slide_cache.get_slide(self.current_slide)
        scaled_current = current_pixmap.scaled(
            self.current_slide_label.size(), 
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation
        )
        self.current_slide_label.setPixmap(scaled_current)
        
        # Next slide
        if self.current_slide < self.slide_cache.total_slides - 1:
            next_pixmap = self.slide_cache.get_slide(self.current_slide + 1)
            scaled_next = next_pixmap.scaled(
                self.next_slide_label.size(),
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation
            )
            self.next_slide_label.setPixmap(scaled_next)
        else:
            self.next_slide_label.clear()
            self.next_slide_label.setText("End of presentation")
        
        # Update slide info
        self.slide_info_label.setText(
            f"Slide {self.current_slide + 1} / {self.slide_cache.total_slides}"
        )
        
        # Update notes
        if self.notes_loader:
            notes = self.notes_loader.get_notes(self.current_slide)
            self.notes_text.setPlainText(notes)
        else:
            self.notes_text.setPlainText("")
        
        # Update fullscreen if presenting
        if self.is_presenting and self.fullscreen_window:
            slide_pixmap = self.slide_cache.get_slide(self.current_slide)
            self.fullscreen_window.show_slide(slide_pixmap)
        
        # Save current position
        self.config.set('Session', 'last_slide', str(self.current_slide))
        self.config.save()
    
    def next_slide(self):
        """Go to next slide"""
        if self.slide_cache and self.current_slide < self.slide_cache.total_slides - 1:
            self.current_slide += 1
            self.update_slides()
    
    def prev_slide(self):
        """Go to previous slide"""
        if self.slide_cache and self.current_slide > 0:
            self.current_slide -= 1
            self.update_slides()
    
    def start_from_beginning(self):
        """Start presentation from first slide"""
        if not self.slide_cache:
            return
        self.current_slide = 0
        self.update_slides()
        self._start_presentation()
    
    def start_from_current(self):
        """Start presentation from current slide"""
        if not self.slide_cache:
            return
        self._start_presentation()
    
    def _start_presentation(self):
        """Initialize and show fullscreen window"""
        # Move presenter window to primary screen first
        screens = QApplication.screens()
        if screens:
            primary_screen = screens[0]
            geometry = primary_screen.availableGeometry()
            self.move(geometry.center() - self.rect().center())
        
        # Find secondary monitor for presentation
        if len(screens) < 2:
            QMessageBox.warning(
                self, "Warning", 
                "No secondary display detected. Presentation will open on primary screen."
            )
            target_screen = screens[0]
        else:
            target_screen = screens[1]
        
        # Create fullscreen window
        if not self.fullscreen_window:
            self.fullscreen_window = FullScreenWindow(self)
        
        # Position on target screen
        geometry = target_screen.geometry()
        self.fullscreen_window.setGeometry(geometry)
        self.fullscreen_window.showFullScreen()
        
        # Show current slide
        slide_pixmap = self.slide_cache.get_slide(self.current_slide)
        self.fullscreen_window.show_slide(slide_pixmap)
        
        # Update state
        self.is_presenting = True
        self.is_blanked = False
        self.start_time = datetime.now()
        
        # Update UI
        self.stop_btn.setEnabled(True)
        self.blank_btn.setEnabled(True)
        self.start_btn.setEnabled(False)
        self.start_current_btn.setEnabled(False)
        
        # Keep focus on presenter window
        self.activateWindow()
        self.raise_()
    
    def stop_presenting(self):
        """Stop presentation and close fullscreen window"""
        if self.fullscreen_window:
            self.fullscreen_window.close()
        
        self.is_presenting = False
        self.is_blanked = False
        self.start_time = None
        
        # Update UI
        self.stop_btn.setEnabled(False)
        self.blank_btn.setEnabled(False)
        self.start_btn.setEnabled(True)
        self.start_current_btn.setEnabled(True)
        self.duration_label.setText("Duration: --:--")
    
    def toggle_blank(self):
        """Toggle screen blanking"""
        if not self.is_presenting or not self.fullscreen_window:
            return
        
        self.is_blanked = not self.is_blanked
        
        if self.is_blanked:
            self.fullscreen_window.blank()
        else:
            self.fullscreen_window.unblank()
            slide_pixmap = self.slide_cache.get_slide(self.current_slide)
            self.fullscreen_window.show_slide(slide_pixmap)
    
    def update_time(self):
        """Update time display"""
        # Current time
        current_time = datetime.now().strftime("%H:%M:%S")
        self.time_label.setText(f"Time: {current_time}")
        
        # Duration
        if self.start_time:
            duration = datetime.now() - self.start_time
            minutes = int(duration.total_seconds() // 60)
            seconds = int(duration.total_seconds() % 60)
            self.duration_label.setText(f"Duration: {minutes:02d}:{seconds:02d}")
    
    def _load_last_session(self):
        """Load last session from config"""
        last_file = self.config.get('Session', 'last_file', fallback=None)
        if last_file and Path(last_file).exists():
            self._load_pdf(last_file)
    
    def closeEvent(self, event):
        """Handle window close"""
        if self.fullscreen_window:
            self.fullscreen_window.close()
        if self.slide_cache:
            self.slide_cache.close()
        event.accept()


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("PDF Show and Tell")
    
    window = PresenterWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
            