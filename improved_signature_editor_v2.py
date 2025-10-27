#!/usr/bin/env python3
"""
Professional Email Signature Editor with Draggable/Resizable Images
Similar to Outlook's signature editor with live preview
No visible HTML code - only rendered signature
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, simpledialog, filedialog, colorchooser
from tkinter import font as tkfont
import base64
from PIL import Image, ImageTk, ImageDraw
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass
from datetime import datetime
import html
import re


@dataclass
class SignatureElement:
    """Represents an element in the signature"""
    element_type: str  # 'text', 'image', 'separator'
    content: str
    font_family: str = "Arial"
    font_size: int = 12
    font_weight: str = "normal"  # 'normal', 'bold'
    font_style: str = "normal"  # 'normal', 'italic'
    color: str = "#000000"
    image_data: Optional[str] = None
    image_width: Optional[int] = None
    image_height: Optional[int] = None


@dataclass
class DraggableImage:
    """Represents a draggable/resizable image"""
    canvas_id: int
    text_index: str  # Position in text widget
    image_data: str  # Base64 data
    width: int
    height: int
    original_width: int
    original_height: int
    photo_image: tk.PhotoImage


class ImprovedSignatureEditor(tk.Toplevel):
    """
    Professional signature editor with live preview and draggable images.

    Features:
    - Rich text editing with live preview
    - Formatting toolbar (bold, italic, font, size, color)
    - Draggable and resizable images/logos
    - No visible HTML code
    - Auto-save functionality
    - Professional layout similar to Outlook
    - Template management
    - Multiple image support
    - Advanced HTML generation
    """

    # Class constants for better maintainability
    WINDOW_WIDTH = 1200
    WINDOW_HEIGHT = 700
    PREVIEW_UPDATE_DELAY = 500  # milliseconds

    FONT_FAMILIES = ["Arial", "Calibri", "Times New Roman", "Courier New", "Verdana", "Georgia"]
    FONT_SIZES = ["8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "36"]

    def __init__(self, parent, database_manager, log_callback, on_save_callback=None):
        """
        Initialize the signature editor.

        Args:
            parent: Parent window
            database_manager: Database manager instance
            log_callback: Logging function
            on_save_callback: Function to call after saving with HTML content
        """
        super().__init__(parent)

        self.db = database_manager
        self.log = log_callback
        self.on_save_callback = on_save_callback  # Store the callback
        self.update_timer = None
        self.images_data = []  # Store multiple images
        self.templates = {}  # Store signature templates
        self.preview_images = []  # Keep references to preview images
        self.draggable_images = []  # Store draggable image objects
        
        # Image dragging/resizing state
        self.selected_image = None
        self.drag_start_x = 0
        self.drag_start_y = 0
        self.resize_mode = None  # 'se', 'sw', 'ne', 'nw', 'e', 'w', 'n', 's', or None

        self._setup_window()
        self._create_ui()
        self._load_existing_signature()
        self._load_templates()

    def _setup_window(self) -> None:
        """Configure window properties"""
        self.title("Email Signature Editor")
        self.geometry(f"{self.WINDOW_WIDTH}x{self.WINDOW_HEIGHT}")
        self.configure(bg="#f0f0f0")

        # Make window modal
        self.transient(self.master)
        self.grab_set()

        # Center window on screen
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (self.WINDOW_WIDTH // 2)
        y = (self.winfo_screenheight() // 2) - (self.WINDOW_HEIGHT // 2)
        self.geometry(f"+{x}+{y}")

    def _create_ui(self) -> None:
        """Create the user interface"""
        # Main container with padding
        main_container = ttk.Frame(self, padding=15)
        main_container.pack(fill=tk.BOTH, expand=True)

        # Header
        self._create_header(main_container)

        # Toolbar
        self._create_toolbar(main_container)

        # Split view: Editor on left, Preview on right
        self._create_split_view(main_container)

        # Bottom buttons
        self._create_bottom_buttons(main_container)

    def _create_header(self, parent: ttk.Frame) -> None:
        """Create header section"""
        header_frame = ttk.Frame(parent)
        header_frame.pack(fill=tk.X, pady=(0, 10))

        title_label = ttk.Label(
            header_frame,
            text="✉️ Email Signature Editor",
            font=("Arial", 16, "bold")
        )
        title_label.pack(side=tk.LEFT)

        subtitle_label = ttk.Label(
            header_frame,
            text="Create a professional signature with live preview - Drag & resize images!",
            font=("Arial", 9),
            foreground="gray"
        )
        subtitle_label.pack(side=tk.LEFT, padx=(10, 0))

    def _create_toolbar(self, parent: ttk.Frame) -> None:
        """Create formatting toolbar"""
        toolbar_frame = ttk.LabelFrame(parent, text="Formatting Tools", padding=10)
        toolbar_frame.pack(fill=tk.X, pady=(0, 10))

        # Row 1: Text formatting
        row1 = ttk.Frame(toolbar_frame)
        row1.pack(fill=tk.X, pady=(0, 5))

        # Font family
        ttk.Label(row1, text="Font:").pack(side=tk.LEFT, padx=(0, 5))
        self.font_var = tk.StringVar(value="Arial")
        font_combo = ttk.Combobox(
            row1,
            textvariable=self.font_var,
            values=self.FONT_FAMILIES,
            width=15,
            state="readonly"
        )
        font_combo.pack(side=tk.LEFT, padx=(0, 10))
        font_combo.bind("<<ComboboxSelected>>", lambda e: self._on_format_change())

        # Font size
        ttk.Label(row1, text="Size:").pack(side=tk.LEFT, padx=(0, 5))
        self.size_var = tk.StringVar(value="12")
        size_combo = ttk.Combobox(
            row1,
            textvariable=self.size_var,
            values=self.FONT_SIZES,
            width=5,
            state="readonly"
        )
        size_combo.pack(side=tk.LEFT, padx=(0, 10))
        size_combo.bind("<<ComboboxSelected>>", lambda e: self._on_format_change())

        # Bold button
        self.bold_btn = ttk.Button(
            row1,
            text="B",
            width=3,
            command=self._toggle_bold
        )
        self.bold_btn.pack(side=tk.LEFT, padx=2)

        # Italic button
        self.italic_btn = ttk.Button(
            row1,
            text="I",
            width=3,
            command=self._toggle_italic
        )
        self.italic_btn.pack(side=tk.LEFT, padx=2)

        # Underline button
        self.underline_btn = ttk.Button(
            row1,
            text="U",
            width=3,
            command=self._toggle_underline
        )
        self.underline_btn.pack(side=tk.LEFT, padx=2)

        # Color button
        self.color_var = tk.StringVar(value="#000000")
        color_btn = ttk.Button(
            row1,
            text="🎨 Color",
            command=self._choose_color
        )
        color_btn.pack(side=tk.LEFT, padx=(10, 2))

        # Alignment buttons
        ttk.Separator(row1, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)
        ttk.Button(row1, text="⬅️", width=3, command=lambda: self._set_alignment("left")).pack(side=tk.LEFT, padx=2)
        ttk.Button(row1, text="↔️", width=3, command=lambda: self._set_alignment("center")).pack(side=tk.LEFT, padx=2)
        ttk.Button(row1, text="➡️", width=3, command=lambda: self._set_alignment("right")).pack(side=tk.LEFT, padx=2)

        # Row 2: Content tools
        row2 = ttk.Frame(toolbar_frame)
        row2.pack(fill=tk.X)

        # Template management
        ttk.Button(
            row2,
            text="📋 Templates",
            command=self._manage_templates
        ).pack(side=tk.LEFT, padx=2)

        # Quick templates
        ttk.Button(
            row2,
            text="📝 Add Contact Info",
            command=self._add_contact_template
        ).pack(side=tk.LEFT, padx=2)

        # Image insertion
        ttk.Button(
            row2,
            text="🖼️ Insert Logo/Image",
            command=self._insert_image
        ).pack(side=tk.LEFT, padx=2)

        # Social media links
        ttk.Button(
            row2,
            text="🔗 Add Social Links",
            command=self._add_social_links
        ).pack(side=tk.LEFT, padx=2)

        # Separator
        ttk.Separator(row2, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        # Clear button
        ttk.Button(
            row2,
            text="🗑️ Clear All",
            command=self._clear_signature
        ).pack(side=tk.LEFT, padx=2)

    def _create_split_view(self, parent: ttk.Frame) -> None:
        """Create split view with editor and preview"""
        split_frame = ttk.Frame(parent)
        split_frame.pack(fill=tk.BOTH, expand=True)

        # Left side: Rich text editor
        left_frame = ttk.LabelFrame(split_frame, text="✏️ Edit Signature", padding=10)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        # Create text editor with formatting support
        self.editor = tk.Text(
            left_frame,
            wrap=tk.WORD,
            font=("Arial", 12),
            undo=True,
            maxundo=-1,
            bg="white",
            relief=tk.FLAT,
            borderwidth=1
        )
        self.editor.pack(fill=tk.BOTH, expand=True)

        # Bind events for live update
        self.editor.bind("<KeyRelease>", lambda e: self._schedule_preview_update())
        self.editor.bind("<<Modified>>", lambda e: self._schedule_preview_update())

        # Configure text tags for formatting
        self._configure_text_tags()

        # Right side: Live preview with canvas for draggable images
        right_frame = ttk.LabelFrame(split_frame, text="👁️ Live Preview (Drag images to reposition)", padding=10)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))

        # Preview canvas for draggable images
        preview_canvas_frame = ttk.Frame(right_frame)
        preview_canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        self.preview_canvas = tk.Canvas(
            preview_canvas_frame,
            bg="white",
            relief=tk.SUNKEN,
            borderwidth=2
        )
        preview_scrollbar = ttk.Scrollbar(preview_canvas_frame, orient=tk.VERTICAL, command=self.preview_canvas.yview)
        self.preview_canvas.configure(yscrollcommand=preview_scrollbar.set)
        
        preview_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.preview_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Bind mouse events for dragging and resizing
        self.preview_canvas.bind("<Button-1>", self._on_canvas_click)
        self.preview_canvas.bind("<B1-Motion>", self._on_canvas_drag)
        self.preview_canvas.bind("<ButtonRelease-1>", self._on_canvas_release)
        self.preview_canvas.bind("<Motion>", self._on_canvas_motion)

        # Add instructions
        instructions = ttk.Label(
            parent,
            text="💡 Tip: Type your signature on the left. In preview, click & drag images to move them, drag corners to resize!",
            foreground="blue",
            font=("Arial", 9)
        )
        instructions.pack(pady=(10, 0))

    def _create_bottom_buttons(self, parent: ttk.Frame) -> None:
        """Create bottom action buttons"""
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=(15, 0))

        # Save button (primary action)
        save_btn = ttk.Button(
            button_frame,
            text="💾 Save Signature",
            command=self._save_signature,
            width=20
        )
        save_btn.pack(side=tk.LEFT, padx=5)

        # Export HTML button
        export_btn = ttk.Button(
            button_frame,
            text="📤 Export HTML",
            command=self._export_html,
            width=20
        )
        export_btn.pack(side=tk.LEFT, padx=5)

        # Test email button
        test_btn = ttk.Button(
            button_frame,
            text="📧 Send Test Email",
            command=self._send_test_email,
            width=20
        )
        test_btn.pack(side=tk.LEFT, padx=5)

        # Close button
        close_btn = ttk.Button(
            button_frame,
            text="❌ Close",
            command=self.destroy,
            width=15
        )
        close_btn.pack(side=tk.RIGHT, padx=5)

    def _configure_text_tags(self) -> None:
        """Configure text tags for the editor"""
        # Bold tag
        bold_font = tkfont.Font(family="Arial", size=12, weight="bold")
        self.editor.tag_configure("bold", font=bold_font)

        # Italic tag
        italic_font = tkfont.Font(family="Arial", size=12, slant="italic")
        self.editor.tag_configure("italic", font=italic_font)

        # Underline tag
        self.editor.tag_configure("underline", underline=True)

        # Bold-italic tag
        bold_italic_font = tkfont.Font(family="Arial", size=12, weight="bold", slant="italic")
        self.editor.tag_configure("bold_italic", font=bold_italic_font)

        # Alignment tags
        self.editor.tag_configure("left", justify=tk.LEFT)
        self.editor.tag_configure("center", justify=tk.CENTER)
        self.editor.tag_configure("right", justify=tk.RIGHT)

    def _toggle_bold(self) -> None:
        """Toggle bold formatting on selected text"""
        try:
            # Get selection
            sel_start = self.editor.index(tk.SEL_FIRST)
            sel_end = self.editor.index(tk.SEL_LAST)

            # Check if already bold
            current_tags = self.editor.tag_names(sel_start)

            if "bold" in current_tags:
                self.editor.tag_remove("bold", sel_start, sel_end)
            else:
                self.editor.tag_add("bold", sel_start, sel_end)

            self._schedule_preview_update()

        except tk.TclError:
            messagebox.showinfo(
                "No Selection",
                "Please select text first before applying formatting.",
                parent=self
            )

    def _toggle_italic(self) -> None:
        """Toggle italic formatting on selected text"""
        try:
            sel_start = self.editor.index(tk.SEL_FIRST)
            sel_end = self.editor.index(tk.SEL_LAST)

            current_tags = self.editor.tag_names(sel_start)

            if "italic" in current_tags:
                self.editor.tag_remove("italic", sel_start, sel_end)
            else:
                self.editor.tag_add("italic", sel_start, sel_end)

            self._schedule_preview_update()

        except tk.TclError:
            messagebox.showinfo(
                "No Selection",
                "Please select text first before applying formatting.",
                parent=self
            )

    def _toggle_underline(self) -> None:
        """Toggle underline formatting on selected text"""
        try:
            sel_start = self.editor.index(tk.SEL_FIRST)
            sel_end = self.editor.index(tk.SEL_LAST)

            current_tags = self.editor.tag_names(sel_start)

            if "underline" in current_tags:
                self.editor.tag_remove("underline", sel_start, sel_end)
            else:
                self.editor.tag_add("underline", sel_start, sel_end)

            self._schedule_preview_update()

        except tk.TclError:
            messagebox.showinfo(
                "No Selection",
                "Please select text first before applying formatting.",
                parent=self
            )

    def _choose_color(self) -> None:
        """Choose text color"""
        try:
            sel_start = self.editor.index(tk.SEL_FIRST)
            sel_end = self.editor.index(tk.SEL_LAST)

            # Open color chooser
            color = colorchooser.askcolor(
                initialcolor=self.color_var.get(),
                parent=self
            )

            if color[1]:  # color[1] is hex code
                self.color_var.set(color[1])
                # Create or update color tag
                tag_name = f"color_{color[1].replace('#', '')}"
                self.editor.tag_configure(tag_name, foreground=color[1])
                self.editor.tag_add(tag_name, sel_start, sel_end)

                self._schedule_preview_update()

        except tk.TclError:
            messagebox.showinfo(
                "No Selection",
                "Please select text first before applying color.",
                parent=self
            )

    def _set_alignment(self, alignment: str) -> None:
        """Set text alignment for selected paragraph"""
        try:
            # Get current line
            current_index = self.editor.index(tk.INSERT)
            line_start = f"{current_index.split('.')[0]}.0"
            line_end = f"{current_index.split('.')[0]}.end"

            # Remove other alignment tags
            for align in ["left", "center", "right"]:
                self.editor.tag_remove(align, line_start, line_end)

            # Add new alignment
            self.editor.tag_add(alignment, line_start, line_end)
            self._schedule_preview_update()

        except Exception as e:
            self.log(f"ERROR: Alignment error: {e}")

    def _on_format_change(self) -> None:
        """Handle font/size changes"""
        try:
            sel_start = self.editor.index(tk.SEL_FIRST)
            sel_end = self.editor.index(tk.SEL_LAST)

            # Create dynamic tag
            font_family = self.font_var.get()
            font_size = int(self.size_var.get())

            tag_name = f"font_{font_family}_{font_size}"
            font_obj = tkfont.Font(family=font_family, size=font_size)
            self.editor.tag_configure(tag_name, font=font_obj)
            self.editor.tag_add(tag_name, sel_start, sel_end)

            self._schedule_preview_update()

        except tk.TclError:
            # No selection, apply to future typing
            pass

    def _add_contact_template(self) -> None:
        """Add contact information template"""
        # Create dialog for contact info
        dialog = tk.Toplevel(self)
        dialog.title("Add Contact Information")
        dialog.geometry("450x400")
        dialog.transient(self)
        dialog.grab_set()

        # Center dialog
        dialog.geometry(f"+{self.winfo_x() + 100}+{self.winfo_y() + 100}")

        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        # Input fields
        fields = {}
        labels = ["Full Name", "Job Title", "Company", "Department", "Phone", "Mobile", "Email", "Website"]

        for i, label in enumerate(labels):
            ttk.Label(frame, text=f"{label}:").grid(row=i, column=0, sticky=tk.W, pady=5)
            entry = ttk.Entry(frame, width=35)
            entry.grid(row=i, column=1, sticky=tk.EW, pady=5, padx=(10, 0))
            fields[label] = entry

        frame.columnconfigure(1, weight=1)

        def insert_contact():
            # Clear editor first if empty
            current_content = self.editor.get("1.0", tk.END).strip()

            # Get current position
            insert_pos = self.editor.index(tk.INSERT)

            # Build contact signature
            name = fields["Full Name"].get().strip()
            if name:
                self.editor.insert(insert_pos, name + "\n")
                # Apply bold formatting to name
                line_start = insert_pos
                line_end = f"{insert_pos}+{len(name)}c"
                self.editor.tag_add("bold", line_start, line_end)
                # Apply larger font
                tag_name = "font_Arial_14"
                self.editor.tag_configure(tag_name, font=tkfont.Font(family="Arial", size=14, weight="bold"))
                self.editor.tag_add(tag_name, line_start, line_end)
                insert_pos = f"{insert_pos}+{len(name)+1}c"

            title = fields["Job Title"].get().strip()
            if title:
                self.editor.insert(insert_pos, title + "\n")
                line_start = insert_pos
                line_end = f"{insert_pos}+{len(title)}c"
                self.editor.tag_add("italic", line_start, line_end)
                insert_pos = f"{insert_pos}+{len(title)+1}c"

            company = fields["Company"].get().strip()
            dept = fields["Department"].get().strip()
            if company:
                company_line = company + (f" - {dept}" if dept else "")
                self.editor.insert(insert_pos, company_line + "\n")
                insert_pos = f"{insert_pos}+{len(company_line)+1}c"

            # Contact details with icons
            contact_line = []
            if fields["Phone"].get().strip():
                contact_line.append(f"📞 {fields['Phone'].get().strip()}")
            if fields["Mobile"].get().strip():
                contact_line.append(f"📱 {fields['Mobile'].get().strip()}")
            if fields["Email"].get().strip():
                contact_line.append(f"✉️ {fields['Email'].get().strip()}")

            if contact_line:
                contact_text = " | ".join(contact_line) + "\n"
                self.editor.insert(insert_pos, contact_text)
                insert_pos = f"{insert_pos}+{len(contact_text)}c"

            website = fields["Website"].get().strip()
            if website:
                website_text = f"🌐 {website}\n"
                self.editor.insert(insert_pos, website_text)

            self._schedule_preview_update()
            dialog.destroy()

        # Buttons
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=len(labels), column=0, columnspan=2, pady=(20, 0))

        ttk.Button(btn_frame, text="Insert", command=insert_contact, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy, width=12).pack(side=tk.LEFT, padx=5)

    def _insert_image(self) -> None:
        """Insert image/logo into signature"""
        file_path = filedialog.askopenfilename(
            title="Select Logo/Image",
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.gif"),
                ("All files", "*.*")
            ],
            parent=self
        )

        if not file_path:
            return

        try:
            # Load image
            pil_image = Image.open(file_path)

            # Ask for size
            max_width = simpledialog.askinteger(
                "Image Width",
                "Enter maximum width (pixels):",
                initialvalue=150,
                minvalue=50,
                maxvalue=500,
                parent=self
            )

            if not max_width:
                return

            # Resize maintaining aspect ratio
            aspect_ratio = pil_image.height / pil_image.width
            new_height = int(max_width * aspect_ratio)
            pil_image = pil_image.resize((max_width, new_height), Image.Resampling.LANCZOS)

            # Convert to base64
            import io
            buffer = io.BytesIO()
            pil_image.save(buffer, format="PNG")
            img_data = base64.b64encode(buffer.getvalue()).decode()

            # Store image data
            image_info = {
                "data": img_data,
                "width": max_width,
                "height": new_height,
                "index": len(self.images_data)
            }
            self.images_data.append(image_info)

            # Show in editor as placeholder
            placeholder = f"\n[Image #{len(self.images_data)}: {max_width}x{new_height}px]\n"
            self.editor.insert(tk.INSERT, placeholder)

            self._schedule_preview_update()

            messagebox.showinfo(
                "Image Added",
                f"Image added successfully ({max_width}x{new_height}px)\n\nImage reference: #{len(self.images_data)}\n\nIn preview, you can drag the image to reposition it and drag corners to resize!",
                parent=self
            )

        except Exception as e:
            self.log(f"ERROR: Failed to insert image: {e}")
            messagebox.showerror(
                "Error",
                f"Failed to load image:\n{str(e)}",
                parent=self
            )

    def _add_social_links(self) -> None:
        """Add social media links"""
        # Create dialog for social links
        dialog = tk.Toplevel(self)
        dialog.title("Add Social Media Links")
        dialog.geometry("450x350")
        dialog.transient(self)
        dialog.grab_set()

        # Center dialog
        dialog.geometry(f"+{self.winfo_x() + 100}+{self.winfo_y() + 100}")

        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        # Social media fields
        fields = {}
        social_platforms = [
            ("LinkedIn", "🔗"),
            ("Twitter/X", "🐦"),
            ("Facebook", "📘"),
            ("Instagram", "📷"),
            ("GitHub", "💻"),
            ("Website", "🌐")
        ]

        for i, (platform, icon) in enumerate(social_platforms):
            ttk.Label(frame, text=f"{icon} {platform}:").grid(row=i, column=0, sticky=tk.W, pady=5)
            entry = ttk.Entry(frame, width=35)
            entry.grid(row=i, column=1, sticky=tk.EW, pady=5, padx=(10, 0))
            fields[platform] = entry

        frame.columnconfigure(1, weight=1)

        def insert_social():
            insert_pos = self.editor.index(tk.INSERT)

            # Add separator if content exists
            current_content = self.editor.get("1.0", tk.END).strip()
            if current_content:
                self.editor.insert(insert_pos, "\n───────────────────\n")
                insert_pos = f"{insert_pos}+2l"

            social_lines = []
            for platform, icon in social_platforms:
                url = fields[platform].get().strip()
                if url:
                    if not url.startswith("http"):
                        url = "https://" + url
                    social_lines.append(f"{icon} {platform}: {url}")

            if social_lines:
                social_text = "\n".join(social_lines) + "\n"
                self.editor.insert(insert_pos, social_text)

            self._schedule_preview_update()
            dialog.destroy()

        # Buttons
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=len(social_platforms), column=0, columnspan=2, pady=(20, 0))

        ttk.Button(btn_frame, text="Insert", command=insert_social, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy, width=12).pack(side=tk.LEFT, padx=5)

    def _manage_templates(self) -> None:
        """Manage signature templates"""
        dialog = tk.Toplevel(self)
        dialog.title("Signature Templates")
        dialog.geometry("500x400")
        dialog.transient(self)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Save and load signature templates", font=("Arial", 11, "bold")).pack(pady=(0, 10))

        # Template list
        list_frame = ttk.LabelFrame(frame, text="Saved Templates", padding=10)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Listbox with scrollbar
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        template_list = tk.Listbox(list_frame, yscrollcommand=scrollbar.set)
        template_list.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=template_list.yview)

        # Load existing templates
        for template_name in self.templates.keys():
            template_list.insert(tk.END, template_name)

        # Buttons
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X)

        def save_template():
            name = simpledialog.askstring("Save Template", "Enter template name:", parent=dialog)
            if name:
                # Get current signature content and formatting
                content = self._get_editor_content_with_formatting()
                self.templates[name] = content
                template_list.insert(tk.END, name)
                self._save_templates_to_db()
                messagebox.showinfo("Success", f"Template '{name}' saved!", parent=dialog)

        def load_template():
            selection = template_list.curselection()
            if selection:
                template_name = template_list.get(selection[0])
                if messagebox.askyesno("Load Template",
                    f"Load template '{template_name}'?\nThis will replace current content.", parent=dialog):
                    self._load_template_content(self.templates[template_name])
                    dialog.destroy()

        def delete_template():
            selection = template_list.curselection()
            if selection:
                template_name = template_list.get(selection[0])
                if messagebox.askyesno("Delete Template",
                    f"Delete template '{template_name}'?", parent=dialog):
                    del self.templates[template_name]
                    template_list.delete(selection[0])
                    self._save_templates_to_db()

        ttk.Button(btn_frame, text="💾 Save Current", command=save_template, width=15).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="📥 Load", command=load_template, width=15).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="🗑️ Delete", command=delete_template, width=15).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Close", command=dialog.destroy, width=15).pack(side=tk.RIGHT, padx=2)

    def _clear_signature(self) -> None:
        """Clear all signature content"""
        if messagebox.askyesno(
            "Confirm Clear",
            "Are you sure you want to clear the entire signature?",
            parent=self
        ):
            self.editor.delete("1.0", tk.END)
            self.images_data = []
            self.draggable_images = []
            self._update_preview()

    def _schedule_preview_update(self) -> None:
        """Schedule preview update with debouncing"""
        if self.update_timer:
            self.after_cancel(self.update_timer)
        self.update_timer = self.after(self.PREVIEW_UPDATE_DELAY, self._update_preview)

    def _update_preview(self) -> None:
        """Update the live preview with draggable images"""
        # Clear canvas
        self.preview_canvas.delete("all")
        self.draggable_images = []
        self.preview_images = []
        
        y_offset = 10
        x_offset = 10
        
        # Get all content
        content = self.editor.get("1.0", tk.END)
        lines = content.split("\n")
        
        for line_num, line in enumerate(lines):
            line_start_index = f"{line_num + 1}.0"
            line_end_index = f"{line_num + 1}.end"
            
            # Check for image placeholder
            image_match = re.match(r'\[Image #(\d+):', line.strip())
            if image_match:
                img_index = int(image_match.group(1)) - 1
                if 0 <= img_index < len(self.images_data):
                    img_data = self.images_data[img_index]
                    
                    # Create PhotoImage from base64
                    photo_image = tk.PhotoImage(data=img_data["data"])
                    self.preview_images.append(photo_image)
                    
                    # Create image on canvas
                    canvas_id = self.preview_canvas.create_image(
                        x_offset, y_offset, 
                        anchor=tk.NW, 
                        image=photo_image
                    )
                    
                    # Create DraggableImage object
                    draggable_img = DraggableImage(
                        canvas_id=canvas_id,
                        text_index=line_start_index,
                        image_data=img_data["data"],
                        width=img_data["width"],
                        height=img_data["height"],
                        original_width=img_data["width"],
                        original_height=img_data["height"],
                        photo_image=photo_image
                    )
                    self.draggable_images.append(draggable_img)
                    
                    # Draw resize handles (corners)
                    self._draw_resize_handles(canvas_id, x_offset, y_offset, img_data["width"], img_data["height"])
                    
                    y_offset += img_data["height"] + 10
                continue
            
            # Handle text lines
            if line.strip():
                # Get all dump data for styling
                all_data = self.editor.dump(line_start_index, line_end_index, "all")
                
                # Simple text rendering for now
                text_id = self.preview_canvas.create_text(
                    x_offset, y_offset,
                    anchor=tk.NW,
                    text=line,
                    font=("Arial", 12),
                    width=500
                )
                
                # Get text bounding box for next line position
                bbox = self.preview_canvas.bbox(text_id)
                if bbox:
                    y_offset = bbox[3] + 5
            else:
                y_offset += 10  # Empty line spacing
        
        # Configure scroll region
        self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all"))
        
        # Reset modified flag
        try:
            self.editor.edit_modified(False)
        except tk.TclError:
            pass

    def _draw_resize_handles(self, image_id, x, y, width, height):
        """Draw resize handles at corners of image"""
        handle_size = 8
        
        # Corner positions
        corners = [
            (x, y, "nw"),  # Top-left
            (x + width, y, "ne"),  # Top-right
            (x, y + height, "sw"),  # Bottom-left
            (x + width, y + height, "se"),  # Bottom-right
        ]
        
        for cx, cy, corner_type in corners:
            handle = self.preview_canvas.create_rectangle(
                cx - handle_size//2, cy - handle_size//2,
                cx + handle_size//2, cy + handle_size//2,
                fill="blue", outline="darkblue", width=2,
                tags=(f"handle_{image_id}", corner_type)
            )

    def _on_canvas_click(self, event):
        """Handle mouse click on canvas"""
        # Check if clicked on an image
        items = self.preview_canvas.find_overlapping(event.x, event.y, event.x, event.y)
        
        for item_id in items:
            # Check if it's an image
            for draggable_img in self.draggable_images:
                if draggable_img.canvas_id == item_id:
                    self.selected_image = draggable_img
                    self.drag_start_x = event.x
                    self.drag_start_y = event.y
                    self.resize_mode = None
                    return
                
            # Check if it's a resize handle
            tags = self.preview_canvas.gettags(item_id)
            for tag in tags:
                if tag in ["nw", "ne", "sw", "se"]:
                    # Find which image this handle belongs to
                    for draggable_img in self.draggable_images:
                        if f"handle_{draggable_img.canvas_id}" in tags:
                            self.selected_image = draggable_img
                            self.drag_start_x = event.x
                            self.drag_start_y = event.y
                            self.resize_mode = tag
                            return
        
        # Clicked on empty space
        self.selected_image = None
        self.resize_mode = None

    def _on_canvas_drag(self, event):
        """Handle mouse drag on canvas"""
        if not self.selected_image:
            return
        
        dx = event.x - self.drag_start_x
        dy = event.y - self.drag_start_y
        
        if self.resize_mode:
            # Resize image
            self._resize_image(dx, dy)
        else:
            # Move image
            self.preview_canvas.move(self.selected_image.canvas_id, dx, dy)
            # Move handles
            for item in self.preview_canvas.find_withtag(f"handle_{self.selected_image.canvas_id}"):
                self.preview_canvas.move(item, dx, dy)
        
        self.drag_start_x = event.x
        self.drag_start_y = event.y

    def _on_canvas_release(self, event):
        """Handle mouse release on canvas"""
        if self.selected_image and self.resize_mode:
            # Update the stored image dimensions
            coords = self.preview_canvas.coords(self.selected_image.canvas_id)
            bbox = self.preview_canvas.bbox(self.selected_image.canvas_id)
            if bbox:
                new_width = bbox[2] - bbox[0]
                new_height = bbox[3] - bbox[1]
                
                # Update the image data
                for img_data in self.images_data:
                    if img_data["data"] == self.selected_image.image_data:
                        img_data["width"] = int(new_width)
                        img_data["height"] = int(new_height)
                        break
                
                # Update draggable image dimensions
                self.selected_image.width = int(new_width)
                self.selected_image.height = int(new_height)
        
        self.resize_mode = None

    def _on_canvas_motion(self, event):
        """Handle mouse motion to change cursor"""
        items = self.preview_canvas.find_overlapping(event.x, event.y, event.x, event.y)
        
        cursor = ""
        for item_id in items:
            tags = self.preview_canvas.gettags(item_id)
            if "nw" in tags or "se" in tags:
                cursor = "size_nw_se"
                break
            elif "ne" in tags or "sw" in tags:
                cursor = "size_ne_sw"
                break
            elif any(tag.startswith("handle_") for tag in tags):
                cursor = "hand2"
                break
        
        if cursor:
            self.preview_canvas.configure(cursor=cursor)
        else:
            # Check if over an image
            for draggable_img in self.draggable_images:
                if draggable_img.canvas_id in items:
                    self.preview_canvas.configure(cursor="fleur")
                    return
            self.preview_canvas.configure(cursor="")

    def _resize_image(self, dx, dy):
        """Resize the selected image based on resize mode"""
        if not self.selected_image or not self.resize_mode:
            return
        
        # Get current image position and size
        coords = self.preview_canvas.coords(self.selected_image.canvas_id)
        x, y = coords[0], coords[1]
        
        current_width = self.selected_image.width
        current_height = self.selected_image.height
        
        # Calculate new dimensions based on corner being dragged
        if self.resize_mode == "se":  # Bottom-right corner
            new_width = max(50, current_width + dx)
            new_height = max(50, current_height + dy)
        elif self.resize_mode == "sw":  # Bottom-left corner
            new_width = max(50, current_width - dx)
            new_height = max(50, current_height + dy)
            if new_width != current_width:
                x += dx
        elif self.resize_mode == "ne":  # Top-right corner
            new_width = max(50, current_width + dx)
            new_height = max(50, current_height - dy)
            if new_height != current_height:
                y += dy
        elif self.resize_mode == "nw":  # Top-left corner
            new_width = max(50, current_width - dx)
            new_height = max(50, current_height - dy)
            if new_width != current_width:
                x += dx
            if new_height != current_height:
                y += dy
        else:
            return
        
        # Maintain aspect ratio
        aspect_ratio = self.selected_image.original_height / self.selected_image.original_width
        new_height = int(new_width * aspect_ratio)
        
        # Resize the image
        try:
            # Decode base64 and create PIL image
            import io
            img_bytes = base64.b64decode(self.selected_image.image_data)
            pil_image = Image.open(io.BytesIO(img_bytes))
            
            # Resize
            resized_image = pil_image.resize((int(new_width), int(new_height)), Image.Resampling.LANCZOS)
            
            # Convert back to PhotoImage
            photo_image = ImageTk.PhotoImage(resized_image)
            self.preview_images.append(photo_image)  # Keep reference
            
            # Update canvas image
            self.preview_canvas.itemconfig(self.selected_image.canvas_id, image=photo_image)
            self.preview_canvas.coords(self.selected_image.canvas_id, x, y)
            
            # Update stored dimensions
            self.selected_image.width = int(new_width)
            self.selected_image.height = int(new_height)
            self.selected_image.photo_image = photo_image
            
            # Redraw handles
            for item in self.preview_canvas.find_withtag(f"handle_{self.selected_image.canvas_id}"):
                self.preview_canvas.delete(item)
            self._draw_resize_handles(self.selected_image.canvas_id, x, y, new_width, new_height)
            
        except Exception as e:
            self.log(f"ERROR: Failed to resize image: {e}")

    def _generate_html(self) -> str:
        """Generate HTML from editor content with proper formatting"""
        html_parts = ['<div style="font-family: Arial, sans-serif; line-height: 1.6;">']

        # Get all content
        content = self.editor.get("1.0", tk.END)
        lines = content.split("\n")

        for line_num, line in enumerate(lines):
            # Use line_num + 1 to construct the index
            line_start_index = f"{line_num + 1}.0"
            line_end_index = f"{line_num + 1}.end"

            # Check for image placeholder
            image_match = re.match(r'\[Image #(\d+):', line.strip())
            if image_match:
                img_index = int(image_match.group(1)) - 1
                if 0 <= img_index < len(self.images_data):
                    img_data = self.images_data[img_index]
                    html_parts.append(
                        f'<img src="data:image/png;base64,{img_data["data"]}" '
                        f'width="{img_data["width"]}" height="{img_data["height"]}" '
                        f'style="display: block; margin: 5px 0;" />'
                    )
                continue

            # Handle text lines
            if not line.strip():
                # It's an empty line, treat as a break
                html_parts.append('<p style="margin: 3px 0;">&nbsp;</p>')
                continue

            # Get all dump data for the current line
            all_data = self.editor.dump(line_start_index, line_end_index, "all")

            if not all_data:
                continue

            # Check for paragraph-level alignment
            para_tags = self.editor.tag_names(line_start_index)
            para_style_dict = {}
            if "center" in para_tags:
                para_style_dict["text-align"] = "center"
            elif "right" in para_tags:
                para_style_dict["text-align"] = "right"
            else:
                para_style_dict["text-align"] = "left"

            line_html_parts = []

            for key, value, index in all_data:
                if key == "text":
                    # Get tags for this specific text segment
                    tags = self.editor.tag_names(index)
                    if tags is None:
                        tags = ()

                    style_dict = {}

                    for tag in tags:
                        if tag.startswith("font_"):
                            parts = tag.split("_")
                            if len(parts) >= 3:
                                font_family_name = " ".join(parts[1:-1])
                                font_size = parts[-1]
                                style_dict["font-family"] = f"'{font_family_name}'"
                                style_dict["font-size"] = f"{font_size}px"
                        elif tag.startswith("color_"):
                            style_dict["color"] = "#" + tag.replace("color_", "")
                        elif tag == "bold":
                            style_dict["font-weight"] = "bold"
                        elif tag == "italic":
                            style_dict["font-style"] = "italic"
                        elif tag == "underline":
                            style_dict["text-decoration"] = "underline"

                    style_str = "; ".join(f"{k}: {v}" for k, v in style_dict.items())

                    # Escape HTML and wrap in span if styled
                    escaped_text = html.escape(value).replace("\n", "")

                    if style_str:
                        line_html_parts.append(f'<span style="{style_str}">{escaped_text}</span>')
                    else:
                        line_html_parts.append(escaped_text)

            # Join all spans into a single paragraph
            para_style_str = "; ".join(f"{k}: {v}" for k, v in para_style_dict.items())
            html_parts.append(f'<p style="margin: 3px 0; {para_style_str}">{"".join(line_html_parts)}</p>')

        html_parts.append('</div>')
        return "\n".join(html_parts)

    def _get_editor_content_with_formatting(self) -> dict:
        """Get editor content with all formatting information"""
        content = {
            "text": self.editor.get("1.0", tk.END),
            "tags": {},
            "images": self.images_data.copy()
        }

        # Store tag information
        for tag_name in self.editor.tag_names():
            if tag_name not in ["sel", "current"]:
                ranges = self.editor.tag_ranges(tag_name)
                if ranges:
                    content["tags"][tag_name] = [(str(ranges[i]), str(ranges[i+1]))
                                                  for i in range(0, len(ranges), 2)]

        return content

    def _load_template_content(self, template_data: dict) -> None:
        """Load template content into editor"""
        # Clear current content
        self.editor.delete("1.0", tk.END)

        # Insert text
        self.editor.insert("1.0", template_data.get("text", ""))

        # Apply tags
        for tag_name, ranges in template_data.get("tags", {}).items():
            for start, end in ranges:
                try:
                    self.editor.tag_add(tag_name, start, end)
                except tk.TclError:
                    self.log(f"WARN: Could not apply tag '{tag_name}' from template.")

        # Restore images
        self.images_data = template_data.get("images", []).copy()

        self._schedule_preview_update()

    def _save_templates_to_db(self) -> None:
        """Save templates to database"""
        try:
            import json

            self.db.execute_query(
                """CREATE TABLE IF NOT EXISTS signature_templates (
                    id INTEGER PRIMARY KEY,
                    templates_json TEXT,
                    last_updated TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )""",
                ()
            )

            # Serialize the templates
            templates_json = json.dumps(self.templates)

            self.db.execute_query(
                """INSERT OR REPLACE INTO signature_templates (id, templates_json, last_updated)
                   VALUES (1, ?, CURRENT_TIMESTAMP)""",
                (templates_json,)
            )

            self.log("INFO: Templates saved successfully")

        except Exception as e:
            self.log(f"ERROR: Failed to save templates: {e}")

    def _load_templates(self) -> None:
        """Load templates from database"""
        try:
            import json

            # Ensure table exists first
            self.db.execute_query(
                """CREATE TABLE IF NOT EXISTS signature_templates (
                    id INTEGER PRIMARY KEY,
                    templates_json TEXT,
                    last_updated TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )""",
                ()
            )

            result = self.db.execute_query(
                "SELECT templates_json FROM signature_templates WHERE id = 1",
                (),
                fetchone=True
            )

            if result and isinstance(result, dict) and result.get('templates_json'):
                self.templates = json.loads(result['templates_json'])
                self.log(f"INFO: Loaded {len(self.templates)} templates")
            else:
                self.templates = {}

        except Exception as e:
            self.log(f"INFO: No templates found or error loading: {e}")
            self.templates = {}

    def _export_html(self) -> None:
        """Export signature as HTML file"""
        try:
            html_content = self._generate_html()

            file_path = filedialog.asksaveasfilename(
                title="Export Signature as HTML",
                defaultextension=".html",
                filetypes=[("HTML files", "*.html"), ("All files", "*.*")],
                parent=self
            )

            if file_path:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(html_content)

                messagebox.showinfo(
                    "Success",
                    f"Signature exported successfully to:\n{file_path}",
                    parent=self
                )

        except Exception as e:
            self.log(f"ERROR: Failed to export HTML: {e}")
            messagebox.showerror(
                "Error",
                f"Failed to export signature:\n{str(e)}",
                parent=self
            )

    def _save_signature(self) -> None:
        """Save signature to database"""
        try:
            html_content = self._generate_html()

            # Ensure table exists
            self.db.execute_query(
                """CREATE TABLE IF NOT EXISTS email_signatures (
                    id INTEGER PRIMARY KEY,
                    html_content TEXT,
                    last_updated TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )""",
                ()
            )

            # Save to database
            self.db.execute_query(
                """INSERT OR REPLACE INTO email_signatures (id, html_content, last_updated)
                   VALUES (1, ?, CURRENT_TIMESTAMP)""",
                (html_content,)
            )

            self.log("INFO: Email signature saved successfully")

            messagebox.showinfo(
                "Success",
                "Email signature saved successfully!\n\nYour signature will be used in all outgoing emails.",
                parent=self
            )

            # Call the callback to update the SettingsWindow display
            if self.on_save_callback:
                self.on_save_callback(html_content)

        except Exception as e:
            self.log(f"ERROR: Failed to save signature: {e}")
            messagebox.showerror(
                "Error",
                f"Failed to save signature:\n{str(e)}",
                parent=self
            )

    def _send_test_email(self) -> None:
        """Send a test email with the signature"""
        messagebox.showinfo(
            "Test Email",
            "Test email functionality will use the saved signature.\n\n"
            "Please save your signature first, then use the 'Test Email' button "
            "in the main Settings window.",
            parent=self
        )

    def _load_existing_signature(self) -> None:
        """Load existing signature from database"""
        try:
            # Ensure table exists first
            self.db.execute_query(
                """CREATE TABLE IF NOT EXISTS email_signatures (
                    id INTEGER PRIMARY KEY,
                    html_content TEXT,
                    last_updated TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )""",
                ()
            )

            result = self.db.execute_query(
                "SELECT html_content FROM email_signatures WHERE id = 1",
                (),
                fetchone=True
            )

            if result and isinstance(result, dict) and result.get('html_content'):
                # Load the HTML into the editor
                # This requires parsing the HTML back into the editor's format (complex)
                # For now, just log that it exists
                self.log("INFO: Existing signature found in database.")

        except Exception as e:
            self.log(f"INFO: No existing signature found or error loading: {e}")


# Export the class
__all__ = ['ImprovedSignatureEditor']
