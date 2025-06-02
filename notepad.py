import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font
import os
from PIL import Image, ImageTk, ImageGrab
import io
import re

class ModernNotepad:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Modern Notepad")
        self.root.geometry("800x600")
        self.root.minsize(600, 400)
        
        # Theme variables
        self.current_theme = "dark"  # dark, light, system
        self.themes = {
            "dark": {
                "bg": "#2b2b2b",
                "header_bg": "#1e1e1e",
                "text_bg": "#1e1e1e",
                "text_fg": "#ffffff",
                "select_bg": "#0078d4",
                "status_fg": "#cccccc"
            },
            "light": {
                "bg": "#f0f0f0",
                "header_bg": "#e0e0e0",
                "text_bg": "#ffffff",
                "text_fg": "#000000",
                "select_bg": "#0078d4",
                "status_fg": "#333333"
            }
        }
        # Initialize system theme
        self.themes["system"] = self.get_system_theme()
        
        # Configure style
        self.setup_styles()
        
        # Variables
        self.current_file = None
        self.is_always_on_top = False
        self.text_modified = False
        self.is_code_mode = False
        self.line_numbers_frame = None
        self.line_numbers_canvas = None
        self.current_language = "Python"
        
        # Supported programming languages
        self.languages = [
            "Python", "JavaScript", "HTML", "CSS", "Java", "C++", "C#", 
            "PHP", "Ruby", "Go", "Swift", "TypeScript", "SQL", "Rust", 
            "Kotlin", "Bash", "PowerShell", "XML", "JSON", "YAML"
        ]
        
        # Create UI
        self.create_header()
        self.create_text_area()
        self.create_status_bar()
        
        # Initialize formatting state tracking
        self.current_formatting = {
            'bold': False,
            'italic': False,
            'underline': False,
            'size': 12,
            'color': None,
            'highlight': None
        }
        
        # Bind events
        self.bind_events()
        
        # Set initial focus
        self.text_area.focus_set()
    
    def get_system_theme(self):
        """Detect system theme (Windows)"""
        try:
            import winreg
            registry = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
            key = winreg.OpenKey(registry, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize")
            value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            winreg.CloseKey(key)
            
            if value == 0:  # Dark theme
                return {
                    "bg": "#2b2b2b", "header_bg": "#1e1e1e", "text_bg": "#1e1e1e",
                    "text_fg": "#ffffff", "select_bg": "#0078d4", "status_fg": "#cccccc"
                }
            else:  # Light theme
                return {
                    "bg": "#f0f0f0", "header_bg": "#e0e0e0", "text_bg": "#ffffff",
                    "text_fg": "#000000", "select_bg": "#0078d4", "status_fg": "#333333"
                }
        except:
            # Default to dark theme if detection fails
            return {
                "bg": "#2b2b2b", "header_bg": "#1e1e1e", "text_bg": "#1e1e1e",
                "text_fg": "#ffffff", "select_bg": "#0078d4", "status_fg": "#cccccc"
            }
    
    def setup_styles(self):
        """Configure modern styling based on current theme"""
        theme = self.themes[self.current_theme]
        self.root.configure(bg=theme["bg"])
        
        # Configure ttk styles
        style = ttk.Style()
        style.theme_use('clam')
        
        # Header style
        style.configure('Header.TFrame', background=theme["header_bg"])
        
        # Determine button colors based on theme
        if self.current_theme == "light":
            button_bg = "#d0d0d0"
            button_active = "#c0c0c0"
            button_pressed = "#b0b0b0"
            text_color = "black"
        else:
            button_bg = "#404040"
            button_active = "#505050"
            button_pressed = "#606060"
            text_color = "white"
        
        style.configure('Header.TButton', 
                       background=button_bg,
                       foreground=text_color,
                       borderwidth=0,
                       focuscolor='none')
        style.map('Header.TButton',
                 background=[('active', button_active),
                           ('pressed', button_pressed)])
        
        # Header label style
        style.configure('Header.TLabel',
                      background=theme["header_bg"],
                      foreground=text_color)
        
        # Always on top button style
        style.configure('OnTop.TButton',
                       background='#0078d4',
                       foreground='white',
                       borderwidth=0,
                       focuscolor='none')
        style.map('OnTop.TButton',
                 background=[('active', '#106ebe'),
                           ('pressed', '#005a9e')])
        
        # Active formatting button style
        style.configure('Active.TButton',
                       background='#0078d4',
                       foreground='white',
                       borderwidth=0,
                       focuscolor='none')
        style.map('Active.TButton',
                 background=[('active', '#106ebe'),
                           ('pressed', '#005a9e')])
        
        # Status bar style
        style.configure('Status.TLabel',
                       background=theme["header_bg"],
                       foreground=theme["status_fg"],
                       padding=(10, 5))
        
        # Combobox style
        combo_bg = "#333333" if self.current_theme != "light" else "#ffffff"
        combo_field_bg = "#333333" if self.current_theme != "light" else "#f0f0f0"
        
        style.configure('TCombobox',
                      fieldbackground=combo_field_bg,
                      background=combo_bg,
                      foreground=text_color,
                      arrowcolor=text_color,
                      selectbackground='#0078d4',
                      selectforeground='white')
        style.map('TCombobox',
                fieldbackground=[('readonly', combo_field_bg)],
                selectbackground=[('readonly', '#0078d4')],
                selectforeground=[('readonly', 'white')])
    
    def create_header(self):
        """Create the header with menu bar"""
        self.header_frame = ttk.Frame(self.root, style='Header.TFrame')
        self.header_frame.pack(fill='x', padx=0, pady=0)
        
        # Create menu bar
        self.create_menu_bar()
        
        # Right side - Always on top button
        right_frame = ttk.Frame(self.header_frame, style='Header.TFrame')
        right_frame.pack(side='right', padx=10, pady=8)
        
        # Language selector (initially hidden, shown in code mode)
        self.language_frame = ttk.Frame(right_frame, style='Header.TFrame')
        self.language_frame.pack(side='left', padx=(0, 10), pady=0)
        
        self.language_label = ttk.Label(self.language_frame, text="Language:", 
                                      style='Header.TLabel', foreground='white')
        self.language_label.pack(side='left', padx=(0, 5))
        
        self.language_var = tk.StringVar(value="Python")
        self.language_dropdown = ttk.Combobox(self.language_frame, 
                                           textvariable=self.language_var,
                                           values=self.languages,
                                           width=12,
                                           state="readonly")
        self.language_dropdown.pack(side='left')
        self.language_dropdown.bind("<<ComboboxSelected>>", self.on_language_change)
        
        # Hide language selector initially (shown only in code mode)
        self.language_frame.pack_forget()
        
        # Text formatting toolbar (shown only in normal mode)
        self.formatting_frame = ttk.Frame(right_frame, style='Header.TFrame')
        self.formatting_frame.pack(side='left', padx=(0, 10), pady=0)
        
        # Font style buttons
        self.bold_button = ttk.Button(self.formatting_frame, text="B", width=3,
                                    command=self.toggle_bold, style='Header.TButton')
        self.bold_button.pack(side='left', padx=1)
        
        self.italic_button = ttk.Button(self.formatting_frame, text="I", width=3,
                                      command=self.toggle_italic, style='Header.TButton')
        self.italic_button.pack(side='left', padx=1)
        
        self.underline_button = ttk.Button(self.formatting_frame, text="U", width=3,
                                         command=self.toggle_underline, style='Header.TButton')
        self.underline_button.pack(side='left', padx=1)
        
        # Font size selector
        self.size_var = tk.StringVar(value="12")
        self.size_dropdown = ttk.Combobox(self.formatting_frame,
                                        textvariable=self.size_var,
                                        values=["8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "32"],
                                        width=4, state="readonly")
        self.size_dropdown.pack(side='left', padx=2)
        self.size_dropdown.bind("<<ComboboxSelected>>", self.change_font_size)
        
        # Text color button
        self.text_color_button = ttk.Button(self.formatting_frame, text="A", width=3,
                                          command=self.change_text_color, style='Header.TButton')
        self.text_color_button.pack(side='left', padx=1)
        
        # Highlight color button
        self.highlight_button = ttk.Button(self.formatting_frame, text="🖍", width=3,
                                         command=self.change_highlight_color, style='Header.TButton')
        self.highlight_button.pack(side='left', padx=1)
        
        # Show formatting toolbar initially (hidden in code mode)
        # Will be managed by mode switching functions
        
        # Always on top button
        self.on_top_button = ttk.Button(right_frame, text="📌 Stay on Top", 
                                       command=self.toggle_always_on_top,
                                       style='Header.TButton')
        self.on_top_button.pack(side='right')
        
        # Add a separator below header
        self.header_separator = ttk.Separator(self.root, orient='horizontal')
        self.header_separator.pack(fill='x', pady=(0, 1))
    
    def create_menu_bar(self):
        """Create menu bar with File menu"""
        menubar = tk.Menu(self.root, bg='#1e1e1e', fg='white', 
                         activebackground='#404040', activeforeground='white',
                         borderwidth=0)
        self.root.config(menu=menubar)
        
        # Get theme colors for menus
        theme = self.themes[self.current_theme]
        menu_bg = theme["header_bg"]
        menu_fg = "black" if self.current_theme == "light" else "white"
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0, bg=menu_bg, fg=menu_fg,
                           activebackground='#0078d4', activeforeground='white',
                           borderwidth=0)
        menubar.add_cascade(label="File", menu=file_menu)
        
        file_menu.add_command(label="New                    Ctrl+N", command=self.new_file)
        file_menu.add_command(label="Open                   Ctrl+O", command=self.open_file)
        file_menu.add_separator()
        file_menu.add_command(label="Save                   Ctrl+S", command=self.save_file)
        file_menu.add_command(label="Save As          Ctrl+Shift+S", command=self.save_as_file)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_closing)
        
        # Edit menu
        edit_menu = tk.Menu(menubar, tearoff=0, bg=menu_bg, fg=menu_fg,
                           activebackground='#0078d4', activeforeground='white',
                           borderwidth=0)
        menubar.add_cascade(label="Edit", menu=edit_menu)
        
        edit_menu.add_command(label="Copy                   Ctrl+C", command=self.copy_text)
        edit_menu.add_command(label="Paste                  Ctrl+V", command=self.paste_content)
        edit_menu.add_command(label="Cut                    Ctrl+X", command=self.cut_text)
        edit_menu.add_separator()
        edit_menu.add_command(label="Select All             Ctrl+A", command=self.select_all)
        
        # View menu
        theme = self.themes[self.current_theme]
        menu_bg = theme["header_bg"]
        menu_fg = "black" if self.current_theme == "light" else "white"
        
        view_menu = tk.Menu(menubar, tearoff=0, bg=menu_bg, fg=menu_fg,
                           activebackground='#0078d4', activeforeground='white',
                           borderwidth=0)
        menubar.add_cascade(label="View", menu=view_menu)
        
        view_menu.add_command(label="Normal Mode", command=self.switch_to_normal_mode)
        view_menu.add_command(label="Code Mode", command=self.switch_to_code_mode)
        view_menu.add_command(label="Spreadsheet Mode", command=self.switch_to_spreadsheet_mode)
        view_menu.add_separator()
        
        # Theme submenu
        theme_menu = tk.Menu(view_menu, tearoff=0, bg=menu_bg, fg=menu_fg,
                            activebackground='#0078d4', activeforeground='white',
                            borderwidth=0)
        view_menu.add_cascade(label="Theme", menu=theme_menu)
        
        theme_menu.add_command(label="Dark Theme", command=lambda: self.change_theme("dark"))
        theme_menu.add_command(label="Light Theme", command=lambda: self.change_theme("light"))
        theme_menu.add_command(label="System Theme", command=lambda: self.change_theme("system"))
    
    def create_text_area(self):
        """Create the main text editing area"""
        # Frame for text area and scrollbar
        theme = self.themes[self.current_theme]
        text_frame = tk.Frame(self.root, bg=theme["bg"])
        text_frame.pack(fill='both', expand=True, padx=10, pady=(0, 10))
        
        # Create main text container
        self.text_container = tk.Frame(text_frame, bg=theme["bg"])
        self.text_container.pack(fill='both', expand=True)
        
        # Text widget with theme-based styling
        theme = self.themes[self.current_theme]
        self.text_area = tk.Text(self.text_container,
                                bg=theme["text_bg"],
                                fg=theme["text_fg"],
                                insertbackground=theme["text_fg"],
                                selectbackground=theme["select_bg"],
                                selectforeground='#ffffff',
                                font=('Consolas', 11),
                                wrap='word',
                                undo=True,
                                maxundo=50,
                                borderwidth=0,
                                highlightthickness=0,
                                padx=15,
                                pady=15)
        
        # Create context menu for right-click
        self.create_context_menu()
        
        # Apply theme to text area
        self.apply_theme_to_text_area()
        
        # Scrollbar
        self.scrollbar = ttk.Scrollbar(self.text_container, orient='vertical', command=self.text_area.yview)
        self.text_area.configure(yscrollcommand=self.scrollbar.set)
        
        # Pack text area and scrollbar
        self.text_area.pack(side='left', fill='both', expand=True)
        self.scrollbar.pack(side='right', fill='y')
        
        # Configure syntax highlighting tags
        self.setup_syntax_highlighting()
    
    def create_status_bar(self):
        """Create status bar"""
        self.status_bar = ttk.Label(self.root, text="Ready", style='Status.TLabel')
        self.status_bar.pack(fill='x', side='bottom')
    
    def bind_events(self):
        """Bind keyboard shortcuts and events"""
        # File operations shortcuts
        self.root.bind('<Control-n>', lambda e: self.new_file())
        self.root.bind('<Control-o>', lambda e: self.open_file())
        self.root.bind('<Control-s>', lambda e: self.save_file())
        self.root.bind('<Control-Shift-S>', lambda e: self.save_as_file())
        
        # Edit shortcuts
        self.root.bind('<Control-c>', lambda e: self.copy_text())
        self.root.bind('<Control-v>', lambda e: self.paste_content())
        self.root.bind('<Control-x>', lambda e: self.cut_text())
        self.root.bind('<Control-a>', lambda e: self.select_all())
        
        # Text modification tracking
        self.text_area.bind('<KeyPress>', self.on_text_change)
        self.text_area.bind('<Button-1>', self.update_status)
        self.text_area.bind('<KeyRelease>', self.update_status)
        
        # Formatting inheritance for new text
        self.text_area.bind('<KeyPress>', self.on_key_press_format, add='+')
        self.text_area.bind('<Button-1>', self.update_current_formatting, add='+')
        self.text_area.bind('<KeyRelease>', self.update_current_formatting, add='+')
        
        # Text replacement when typing over selection
        self.text_area.bind('<Key>', self.handle_text_replacement, add='+')
        self.text_area.bind('<BackSpace>', self.handle_backspace, add='+')
        self.text_area.bind('<Delete>', self.handle_delete, add='+')
        
        # Code mode specific bindings
        self.text_area.bind('<KeyRelease>', self.on_key_release)
        self.text_area.bind('<Button-1>', self.on_click)
        
        # Right-click context menu
        self.text_area.bind('<Button-3>', self.show_context_menu)
        
        # Window close event
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def toggle_always_on_top(self):
        """Toggle the always on top functionality"""
        self.is_always_on_top = not self.is_always_on_top
        self.root.attributes('-topmost', self.is_always_on_top)
        
        if self.is_always_on_top:
            self.on_top_button.configure(text="📌 On Top", style='OnTop.TButton')
            self.status_bar.configure(text="Window is now always on top")
        else:
            self.on_top_button.configure(text="📌 Stay on Top", style='Header.TButton')
            self.status_bar.configure(text="Window is no longer always on top")
    
    def new_file(self):
        """Create a new file"""
        if self.check_unsaved_changes():
            self.text_area.delete(1.0, tk.END)
            self.current_file = None
            self.text_modified = False
            self.update_title()
            self.status_bar.configure(text="New file created")
    
    def open_file(self):
        """Open an existing file"""
        if self.check_unsaved_changes():
            file_path = filedialog.askopenfilename(
                title="Open File",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
            )
            if file_path:
                try:
                    with open(file_path, 'r', encoding='utf-8') as file:
                        content = file.read()
                        self.text_area.delete(1.0, tk.END)
                        self.text_area.insert(1.0, content)
                        self.current_file = file_path
                        self.text_modified = False
                        self.update_title()
                        self.status_bar.configure(text=f"Opened: {os.path.basename(file_path)}")
                except Exception as e:
                    messagebox.showerror("Error", f"Could not open file: {str(e)}")
    
    def save_file(self):
        """Save the current file"""
        if self.current_file:
            try:
                content = self.text_area.get(1.0, tk.END + '-1c')
                with open(self.current_file, 'w', encoding='utf-8') as file:
                    file.write(content)
                self.text_modified = False
                self.update_title()
                self.status_bar.configure(text=f"Saved: {os.path.basename(self.current_file)}")
            except Exception as e:
                messagebox.showerror("Error", f"Could not save file: {str(e)}")
        else:
            self.save_as_file()
    
    def save_as_file(self):
        """Save the file with a new name"""
        # Check if we're in spreadsheet mode
        is_spreadsheet_mode = hasattr(self, 'spreadsheet_frame') and self.spreadsheet_frame.winfo_ismapped()
        
        if is_spreadsheet_mode:
            # Spreadsheet mode - save as Excel file
            default_ext = ".xlsx"
            filetypes = [("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")]
            
            file_path = filedialog.asksaveasfilename(
                title="Save Spreadsheet As",
                defaultextension=default_ext,
                filetypes=filetypes
            )
            
            if file_path:
                try:
                    self.save_spreadsheet_data(file_path)
                    self.current_file = file_path
                    self.text_modified = False
                    self.update_title()
                    self.status_bar.configure(text=f"Saved as: {os.path.basename(file_path)}")
                except Exception as e:
                    messagebox.showerror("Error", f"Could not save spreadsheet: {str(e)}")
        else:
            # Text/Code mode - save as text file
            # Determine default extension based on selected language
            default_ext = ".txt"
            filetypes = [("Text files", "*.txt")]
            
            if self.is_code_mode:
                # Map languages to their file extensions
                lang_extensions = {
                    "Python": ".py",
                    "JavaScript": ".js",
                    "HTML": ".html",
                    "CSS": ".css",
                    "Java": ".java",
                    "C++": ".cpp",
                    "C#": ".cs",
                    "PHP": ".php",
                    "Ruby": ".rb",
                    "Go": ".go",
                    "Swift": ".swift",
                    "TypeScript": ".ts",
                    "SQL": ".sql",
                    "Rust": ".rs",
                    "Kotlin": ".kt",
                    "Bash": ".sh",
                    "PowerShell": ".ps1",
                    "XML": ".xml",
                    "JSON": ".json",
                    "YAML": ".yml"
                }
                
                # Set default extension based on current language
                if self.current_language in lang_extensions:
                    default_ext = lang_extensions[self.current_language]
                    filetypes.insert(0, (f"{self.current_language} files", f"*{default_ext}"))
            
            # Add all files option
            filetypes.append(("All files", "*.*"))
            
            file_path = filedialog.asksaveasfilename(
                title="Save As",
                defaultextension=default_ext,
                filetypes=filetypes
            )
            
            if file_path:
                try:
                    content = self.text_area.get(1.0, tk.END + '-1c')
                    with open(file_path, 'w', encoding='utf-8') as file:
                        file.write(content)
                    self.current_file = file_path
                    self.text_modified = False
                    self.update_title()
                    self.status_bar.configure(text=f"Saved as: {os.path.basename(file_path)}")
                except Exception as e:
                    messagebox.showerror("Error", f"Could not save file: {str(e)}")
    
    def save_spreadsheet_data(self, file_path):
        """Save spreadsheet data to Excel or CSV format"""
        import csv
        
        # Determine file format based on extension
        file_ext = os.path.splitext(file_path)[1].lower()
        
        if file_ext == '.csv':
            # Save as CSV
            with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                
                # Write data row by row
                for row in range(1, self.max_rows):
                    row_data = []
                    for col in range(1, self.max_cols):
                        cell_id = f"{chr(64 + col)}{row}"
                        if cell_id in self.cells:
                            cell_value = self.cells[cell_id].get()
                            row_data.append(cell_value)
                        else:
                            row_data.append("")
                    
                    # Only write rows that have some data
                    if any(cell.strip() for cell in row_data):
                        writer.writerow(row_data)
        
        elif file_ext == '.xlsx':
            # Save as Excel - try to use openpyxl if available
            try:
                import openpyxl
                from openpyxl import Workbook
                
                wb = Workbook()
                ws = wb.active
                ws.title = "Sheet1"
                
                # Write column headers
                for col in range(1, self.max_cols):
                    col_letter = chr(64 + col)
                    ws.cell(row=1, column=col, value=col_letter)
                
                # Write data
                for row in range(1, self.max_rows):
                    for col in range(1, self.max_cols):
                        cell_id = f"{chr(64 + col)}{row}"
                        if cell_id in self.cells:
                            cell_value = self.cells[cell_id].get()
                            if cell_value.strip():  # Only write non-empty cells
                                ws.cell(row=row+1, column=col, value=cell_value)
                
                wb.save(file_path)
                
            except ImportError:
                # Fallback to CSV if openpyxl is not available
                messagebox.showwarning("Excel Support", 
                    "Excel format requires 'openpyxl' library. Saving as CSV instead.\n"
                    "To install: pip install openpyxl")
                csv_path = os.path.splitext(file_path)[0] + '.csv'
                self.save_spreadsheet_data(csv_path)
                return
        
        else:
            # Unsupported format, save as CSV
            csv_path = os.path.splitext(file_path)[0] + '.csv'
            self.save_spreadsheet_data(csv_path)
    
    def on_text_change(self, event=None):
        """Handle text changes"""
        if not self.text_modified:
            self.text_modified = True
            self.update_title()
    
    def update_status(self, event=None):
        """Update status bar with cursor position"""
        cursor_pos = self.text_area.index(tk.INSERT)
        line, col = cursor_pos.split('.')
        total_lines = self.text_area.index(tk.END + '-1c').split('.')[0]
        self.status_bar.configure(text=f"Line {line}, Column {int(col)+1} | Total Lines: {total_lines}")
    
    def update_title(self):
        """Update window title"""
        title = "Modern Notepad"
        if self.current_file:
            filename = os.path.basename(self.current_file)
            title = f"{filename} - Modern Notepad"
        if self.text_modified:
            title = f"*{title}"
        self.root.title(title)
    
    def check_unsaved_changes(self):
        """Check for unsaved changes and prompt user"""
        if self.text_modified:
            response = messagebox.askyesnocancel(
                "Unsaved Changes",
                "You have unsaved changes. Do you want to save before continuing?"
            )
            if response is True:  # Yes, save
                self.save_file()
                return True
            elif response is False:  # No, don't save
                return True
            else:  # Cancel
                return False
        return True
    
    def create_context_menu(self):
        """Create right-click context menu"""
        theme = self.themes[self.current_theme]
        menu_bg = theme["header_bg"]
        menu_fg = "black" if self.current_theme == "light" else "white"
        
        self.context_menu = tk.Menu(self.root, tearoff=0, bg=menu_bg, fg=menu_fg,
                                   activebackground='#0078d4', activeforeground='white',
                                   borderwidth=0)
        self.context_menu.add_command(label="Copy", command=self.copy_text)
        self.context_menu.add_command(label="Paste", command=self.paste_content)
        self.context_menu.add_command(label="Cut", command=self.cut_text)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Select All", command=self.select_all)
    
    def show_context_menu(self, event):
        """Show context menu on right-click"""
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()
    
    def copy_text(self):
        """Copy selected text to clipboard"""
        try:
            selected_text = self.text_area.selection_get()
            self.root.clipboard_clear()
            self.root.clipboard_append(selected_text)
            self.status_bar.configure(text="Text copied to clipboard")
        except tk.TclError:
            self.status_bar.configure(text="No text selected")
    
    def cut_text(self):
        """Cut selected text to clipboard"""
        try:
            selected_text = self.text_area.selection_get()
            self.root.clipboard_clear()
            self.root.clipboard_append(selected_text)
            self.text_area.delete(tk.SEL_FIRST, tk.SEL_LAST)
            self.on_text_change()
            self.status_bar.configure(text="Text cut to clipboard")
        except tk.TclError:
            self.status_bar.configure(text="No text selected")
    
    def paste_content(self):
        """Paste content from clipboard (text or image)"""
        try:
            # Try to get image from clipboard first
            try:
                image = ImageGrab.grabclipboard()
                if image:
                    self.paste_image(image)
                    return
            except:
                pass
            
            # If no image, try to get text
            clipboard_text = self.root.clipboard_get()
            if clipboard_text:
                # Insert at cursor position
                cursor_pos = self.text_area.index(tk.INSERT)
                self.text_area.insert(cursor_pos, clipboard_text)
                self.on_text_change()
                self.status_bar.configure(text="Text pasted from clipboard")
        except tk.TclError:
            self.status_bar.configure(text="Nothing to paste")
    
    def paste_image(self, image):
        """Paste image into text area"""
        try:
            # Resize image if too large
            max_width, max_height = 400, 300
            if image.width > max_width or image.height > max_height:
                image.thumbnail((max_width, max_height), Image.Resampling.LANCZOS)
            
            # Convert to PhotoImage
            photo = ImageTk.PhotoImage(image)
            
            # Get cursor position and ensure proper text positioning
            cursor_pos = self.text_area.index(tk.INSERT)
            
            # Check if we're at the beginning of a line, if not, add newline before image
            line_start = cursor_pos.split('.')[0] + '.0'
            if cursor_pos != line_start and self.text_area.get(line_start, cursor_pos).strip():
                self.text_area.insert(cursor_pos, '\n')
                cursor_pos = self.text_area.index(tk.INSERT)
            
            # Insert image at cursor position with baseline alignment to fix text positioning
            image_name = self.text_area.image_create(cursor_pos, image=photo, align='baseline')
            
            # Move cursor to end of current line and add newline to ensure text goes below image
            current_line = cursor_pos.split('.')[0]
            line_end = f"{current_line}.end"
            self.text_area.mark_set(tk.INSERT, line_end)
            self.text_area.insert(tk.INSERT, '\n')
            
            # Make image selectable by creating a tag around it
            image_start = cursor_pos
            image_end = self.text_area.index(f"{cursor_pos}+1c")
            image_tag = f"image_{len(getattr(self, 'images', []))}"
            self.text_area.tag_add(image_tag, image_start, image_end)
            
            # Configure image tag to be selectable
            self.text_area.tag_configure(image_tag, 
                                       selectbackground=self.themes[self.current_theme]["select_bg"],
                                       selectforeground='white')
            
            # Keep a reference to prevent garbage collection
            if not hasattr(self, 'images'):
                self.images = []
            self.images.append(photo)
            
            # Add highlighting around the image in normal mode
            if not self.is_code_mode:
                self.highlight_pasted_image(image_start, image_name)
            
            self.on_text_change()
            self.status_bar.configure(text="Image pasted from clipboard")
        except Exception as e:
            self.status_bar.configure(text=f"Error pasting image: {str(e)}")
    
    def select_all(self):
        """Select all text"""
        self.text_area.tag_add(tk.SEL, "1.0", tk.END)
        self.text_area.mark_set(tk.INSERT, "1.0")
        self.text_area.see(tk.INSERT)
        self.status_bar.configure(text="All text selected")
    
    def setup_syntax_highlighting(self):
        """Setup syntax highlighting tags for code mode"""
        # Adjust colors based on theme
        if self.current_theme == "light":
            keyword_color = "#0000ff"
            string_color = "#008000"
            comment_color = "#808080"
            number_color = "#ff0000"
            function_color = "#800080"
        else:
            keyword_color = "#569cd6"
            string_color = "#ce9178"
            comment_color = "#6a9955"
            number_color = "#b5cea8"
            function_color = "#dcdcaa"
        
        # Keywords
        self.text_area.tag_configure("keyword", foreground=keyword_color)
        # Strings
        self.text_area.tag_configure("string", foreground=string_color)
        # Comments
        self.text_area.tag_configure("comment", foreground=comment_color)
        # Numbers
        self.text_area.tag_configure("number", foreground=number_color)
        # Functions
        self.text_area.tag_configure("function", foreground=function_color)
        # Line numbers
        self.text_area.tag_configure("line_number", foreground="#858585")
        # Image highlight tag for normal mode
        self.text_area.tag_configure("image_highlight", background="#0078d4", 
                                   relief="solid", borderwidth=2)
    
    def highlight_pasted_image(self, cursor_pos, image_name):
        """Add visual highlighting around pasted image in normal mode"""
        try:
            # Create a tag around the image position
            end_pos = f"{cursor_pos}+1c"
            tag_name = f"img_highlight_{len(self.images)}"
            
            # Use theme-appropriate highlight color
            highlight_color = "#0078d4"  # Blue for all themes
            border_color = "#ffffff" if self.current_theme != "light" else "#000000"
            
            # Configure the highlight tag
            self.text_area.tag_configure(tag_name, background=highlight_color, 
                                       relief="solid", borderwidth=2,
                                       foreground=border_color,
                                       bgstipple="gray25")
            
            # Apply the tag to the image position
            self.text_area.tag_add(tag_name, cursor_pos, end_pos)
            
            # Store the tag name for potential removal later
            if not hasattr(self, 'image_tags'):
                self.image_tags = []
            self.image_tags.append(tag_name)
            
            # Auto-remove highlight after 3 seconds to show it was selected
            self.root.after(3000, lambda: self.remove_image_highlight(tag_name))
            
        except Exception as e:
            print(f"Error highlighting image: {e}")
    
    def remove_image_highlight(self, tag_name):
        """Remove image highlight after delay"""
        try:
            self.text_area.tag_delete(tag_name)
            if hasattr(self, 'image_tags') and tag_name in self.image_tags:
                self.image_tags.remove(tag_name)
        except Exception as e:
            print(f"Error removing image highlight: {e}")
    
    def switch_to_code_mode(self):
        """Switch to code editor mode"""
        # Handle switching from spreadsheet mode
        if hasattr(self, 'spreadsheet_frame') and self.spreadsheet_frame.winfo_ismapped():
            self.spreadsheet_frame.pack_forget()
            # Clear spreadsheet data
            if hasattr(self, 'cells'):
                for cell_id, entry in self.cells.items():
                    entry.delete(0, tk.END)
            # Show text area and scrollbar
            self.text_area.pack(side='left', fill='both', expand=True)
            self.scrollbar.pack(side='right', fill='y')
        
        if not self.is_code_mode:
            self.is_code_mode = True
            self.text_area.configure(wrap='none', padx=5)
            
            # Make sure text area is packed before adding line numbers
            if not self.text_area.winfo_ismapped():
                self.text_area.pack(side='left', fill='both', expand=True)
                self.scrollbar.pack(side='right', fill='y')
            
            # Create line numbers after text area is packed
            self.create_line_numbers()
            
            # Show language selector and hide formatting toolbar
            self.language_frame.pack(side='left', padx=(0, 10), pady=0)
            self.formatting_frame.pack_forget()
            
            # Apply syntax highlighting based on selected language
            self.apply_syntax_highlighting()
            self.status_bar.configure(text=f"Switched to Code Mode - {self.current_language}")
            self.root.title(f"Code ({self.current_language}) - Modern Notepad")
    
    def switch_to_normal_mode(self):
        """Switch to normal notepad mode"""
        # Handle switching from code mode
        if self.is_code_mode:
            self.is_code_mode = False
            self.remove_line_numbers()
            self.text_area.configure(wrap='word', padx=15)
            self.clear_syntax_highlighting()
        
        # Handle switching from spreadsheet mode
        if hasattr(self, 'spreadsheet_frame') and self.spreadsheet_frame.winfo_ismapped():
            self.spreadsheet_frame.pack_forget()
            # Clear spreadsheet data
            if hasattr(self, 'cells'):
                for cell_id, entry in self.cells.items():
                    entry.delete(0, tk.END)
            # Show text area and scrollbar if they're hidden
            self.text_area.pack(side='left', fill='both', expand=True)
            self.scrollbar.pack(side='right', fill='y')
        
        # Make sure text area is visible
        if not self.text_area.winfo_ismapped():
            self.text_area.pack(side='left', fill='both', expand=True)
            self.scrollbar.pack(side='right', fill='y')
        
        # Show formatting toolbar
        self.formatting_frame.pack(side='left', padx=(0, 10), pady=0)
        
        # Hide language selector
        self.language_frame.pack_forget()
        
        self.status_bar.configure(text="Switched to Normal Mode")
        self.update_title()
    
    def switch_to_spreadsheet_mode(self):
        """Switch to spreadsheet mode with Excel-like grid"""
        # Clear existing spreadsheet data if it exists
        if hasattr(self, 'cells'):
            for cell_id, entry in self.cells.items():
                entry.delete(0, tk.END)
        
        # Hide text area and related components
        self.text_area.pack_forget()
        self.scrollbar.pack_forget()
        
        # Remove line numbers if in code mode
        if self.is_code_mode:
            self.is_code_mode = False
            self.remove_line_numbers()
            self.clear_syntax_highlighting()
        
        # Hide language selector and formatting toolbar
        self.language_frame.pack_forget()
        self.formatting_frame.pack_forget()
        
        # Create spreadsheet frame if it doesn't exist
        if not hasattr(self, 'spreadsheet_frame'):
            self.create_spreadsheet_view()
        
        # Show spreadsheet view
        self.spreadsheet_frame.pack(fill='both', expand=True)
        
        # Update status and title
        self.status_bar.configure(text="Switched to Spreadsheet Mode")
        self.root.title("Spreadsheet - Modern Notepad")
    
    def create_spreadsheet_view(self):
        """Create Excel-like spreadsheet view"""
        theme = self.themes[self.current_theme]
        
        # Initialize spreadsheet properties
        self.max_rows = 101  # 1-100 rows initially (0 is header)
        self.max_cols = 27   # 1-26 columns initially (0 is header, 1-26 for A-Z)
        
        # Create main spreadsheet frame
        self.spreadsheet_frame = tk.Frame(self.text_container, bg=theme["bg"])
        
        # Define consistent column width
        cell_width = 10
        row_header_width = 4
        
        # Create toolbar for spreadsheet actions
        self.spreadsheet_toolbar = tk.Frame(self.spreadsheet_frame, bg=theme["header_bg"])
        self.spreadsheet_toolbar.pack(fill='x', side='top')
        
        # Add Row button
        add_row_btn = tk.Button(self.spreadsheet_toolbar, text="Add Row", 
                              command=self.add_spreadsheet_row,
                              bg=theme["header_bg"], fg=theme["text_fg"],
                              activebackground=theme["select_bg"])
        add_row_btn.pack(side='left', padx=5, pady=2)
        
        # Add Column button
        add_col_btn = tk.Button(self.spreadsheet_toolbar, text="Add Column", 
                              command=self.add_spreadsheet_column,
                              bg=theme["header_bg"], fg=theme["text_fg"],
                              activebackground=theme["select_bg"])
        add_col_btn.pack(side='left', padx=5, pady=2)
        
        # Create header container frame
        self.header_container = tk.Frame(self.spreadsheet_frame, bg=theme["header_bg"])
        self.header_container.pack(fill='x', side='top')
        
        # Create corner cell (fixed)
        self.corner_frame = tk.Frame(self.header_container, bg=theme["header_bg"])
        self.corner_frame.pack(side='left', fill='y')
        corner_label = tk.Label(self.corner_frame, width=row_header_width, bg=theme["header_bg"], 
                              fg=theme["text_fg"], relief='raised', borderwidth=1)
        corner_label.pack(fill='both', expand=True)
        
        # Create header canvas for scrollable column headers
        self.header_canvas = tk.Canvas(self.header_container, bg=theme["header_bg"], 
                                    height=25, highlightthickness=0)  # Fixed height for header
        self.header_canvas.pack(side='left', fill='x', expand=True)
        
        # Create header frame inside canvas
        self.header_frame = tk.Frame(self.header_canvas, bg=theme["header_bg"])
        self.header_window = self.header_canvas.create_window((0, 0), window=self.header_frame, anchor='nw')
        
        # Column headers (A, B, C, ...)
        for col in range(1, self.max_cols):
            col_label = tk.Label(self.header_frame, text=chr(64 + col), width=cell_width, 
                               bg=theme["header_bg"], fg=theme["text_fg"],
                               relief='raised', borderwidth=1)
            col_label.grid(row=0, column=col-1, sticky='nsew')  # Column index starts at 0 in header frame
            self.header_frame.columnconfigure(col-1, minsize=cell_width*8)  # Consistent width
        
        # Create canvas for scrollable grid
        self.canvas_frame = tk.Frame(self.spreadsheet_frame)
        self.canvas_frame.pack(fill='both', expand=True, side='top')
        
        self.canvas = tk.Canvas(self.canvas_frame, bg=theme["text_bg"])
        self.canvas.pack(side='left', fill='both', expand=True)
        
        # Add scrollbars
        # Sync function for horizontal scrolling
        def on_xview(*args):
            # Update both canvases with exact same scroll position
            self.canvas.xview(*args)
            
            # Get the current scroll position from the main canvas
            # and apply exactly the same position to the header canvas
            if args and len(args) > 0 and args[0] == 'moveto':
                # For 'moveto' commands, use the exact same position
                self.header_canvas.xview_moveto(float(args[1]))
            else:
                # For other commands (like scroll units/pages), sync after the main canvas updates
                self.header_canvas.xview_moveto(self.canvas.xview()[0])
        
        self.x_scrollbar = ttk.Scrollbar(self.spreadsheet_frame, orient='horizontal', command=on_xview)
        self.x_scrollbar.pack(side='bottom', fill='x')
        
        self.y_scrollbar = ttk.Scrollbar(self.canvas_frame, orient='vertical', command=self.canvas.yview)
        self.y_scrollbar.pack(side='right', fill='y')
        
        # Configure scrolling for both canvases
        self.canvas.configure(xscrollcommand=self.x_scrollbar.set, yscrollcommand=self.y_scrollbar.set)
        self.header_canvas.configure(xscrollcommand=self.x_scrollbar.set)
        
        # Create frame for cells
        self.grid_frame = tk.Frame(self.canvas, bg=theme["text_bg"])
        self.canvas.create_window((0, 0), window=self.grid_frame, anchor='nw')
        
        # Configure column widths for grid frame
        self.grid_frame.columnconfigure(0, minsize=row_header_width*8)  # Row header column
        for col in range(1, self.max_cols):
            self.grid_frame.columnconfigure(col, minsize=cell_width*8)  # Consistent width
        
        # Row headers (1, 2, 3, ...)
        for row in range(1, self.max_rows):
            row_label = tk.Label(self.grid_frame, text=str(row), width=row_header_width, height=1,
                               bg=theme["header_bg"], fg=theme["text_fg"],
                               relief='raised', borderwidth=1)
            row_label.grid(row=row, column=0, sticky='nsew')
        
        # Create cells
        self.cells = {}
        for row in range(1, self.max_rows):
            # Configure row height
            self.grid_frame.rowconfigure(row, minsize=25)  # Consistent row height
            
            for col in range(1, self.max_cols):
                cell = tk.Entry(self.grid_frame, width=cell_width, bg=theme["text_bg"], 
                              fg=theme["text_fg"], borderwidth=1, relief='solid')
                cell.grid(row=row, column=col, sticky='nsew')
                cell.insert(0, "")  # Empty by default
                
                # Store cell reference
                cell_id = f"{chr(64 + col)}{row}"  # e.g., A1, B2, etc.
                self.cells[cell_id] = cell
                
                # Bind events for cell navigation and clipboard
                cell.bind("<Return>", lambda e, r=row, c=col: self.move_to_cell(r+1, c))
                cell.bind("<Tab>", lambda e, r=row, c=col: self.move_to_cell(r, c+1))
                cell.bind("<Shift-Tab>", lambda e, r=row, c=col: self.move_to_cell(r, c-1))
                cell.bind("<Up>", lambda e, r=row, c=col: self.move_to_cell(r-1, c))
                cell.bind("<Down>", lambda e, r=row, c=col: self.move_to_cell(r+1, c))
                cell.bind("<Left>", lambda e, r=row, c=col: self.move_to_cell(r, c-1))
                cell.bind("<Right>", lambda e, r=row, c=col: self.move_to_cell(r, c+1))
                
                # Clipboard bindings
                cell.bind("<Control-v>", self.paste_to_cells)
                cell.bind("<Control-c>", self.copy_from_cells)
        
        # Update canvas scroll region
        self.grid_frame.update_idletasks()
        self.header_frame.update_idletasks()
        
        # Configure scroll regions for both canvases
        self.canvas.config(scrollregion=self.canvas.bbox("all"))
        self.header_canvas.config(scrollregion=(0, 0, self.grid_frame.winfo_width(), self.header_frame.winfo_height()))
        
        # Make sure header canvas width matches grid width
        self.header_canvas.itemconfig(self.header_window, width=self.grid_frame.winfo_width())
        
        # Bind grid frame changes to update header canvas
        def on_frame_configure(event):
            # Update the scroll region of both canvases
            self.canvas.config(scrollregion=self.canvas.bbox("all"))
            self.header_canvas.config(scrollregion=(0, 0, self.grid_frame.winfo_width(), self.header_frame.winfo_height()))
            # Update header canvas width
            self.header_canvas.itemconfig(self.header_window, width=self.grid_frame.winfo_width())
        
        self.grid_frame.bind("<Configure>", on_frame_configure)
        
        # Hide spreadsheet frame initially
        self.spreadsheet_frame.pack_forget()
    
    def move_to_cell(self, row, col):
        """Move focus to specified cell"""
        # Ensure row and column are within bounds
        row = max(1, min(row, self.max_rows - 1))
        col = max(1, min(col, self.max_cols - 1))
        
        # Get cell ID and focus on it
        cell_id = f"{chr(64 + col)}{row}"
        if cell_id in self.cells:
            self.cells[cell_id].focus_set()
            
            # Ensure cell is visible by scrolling if needed
            self.canvas.yview_moveto((row - 1) / (self.max_rows - 1))
            self.canvas.xview_moveto((col - 1) / (self.max_cols - 1))
        
        return "break"  # Prevent default behavior
        
    def copy_from_cells(self, event=None):
        """Copy selected cell content to clipboard"""
        if event and event.widget:
            # Get content from the current cell
            content = event.widget.get()
            # Copy to clipboard
            self.clipboard_clear()
            self.clipboard_append(content)
            return "break"  # Prevent default behavior
    
    def paste_to_cells(self, event=None):
        """Paste clipboard content to cells"""
        if event and event.widget:
            try:
                # Get clipboard content
                clipboard_content = self.clipboard_get()
                
                # Find current cell position
                current_cell = event.widget
                current_cell_id = None
                current_row = None
                current_col = None
                
                # Find the cell ID and position
                for cell_id, cell in self.cells.items():
                    if cell == current_cell:
                        current_cell_id = cell_id
                        # Extract row and column from cell_id (e.g., 'A1' -> col=1, row=1)
                        current_col = ord(cell_id[0]) - 64  # A=1, B=2, etc.
                        current_row = int(cell_id[1:])  # Extract row number
                        break
                
                if current_row is not None and current_col is not None:
                    # Check if content has tab or newline characters (table data)
                    if '\t' in clipboard_content or '\n' in clipboard_content:
                        # Split by rows and columns
                        rows = clipboard_content.strip().split('\n')
                        
                        for r_idx, row_data in enumerate(rows):
                            # Split row by tabs or multiple spaces
                            import re
                            cells_data = re.split(r'\t|\s{2,}', row_data)
                            
                            for c_idx, cell_data in enumerate(cells_data):
                                # Calculate target cell position
                                target_row = current_row + r_idx
                                target_col = current_col + c_idx
                                
                                # Ensure we don't exceed grid boundaries
                                if target_row < self.max_rows and target_col < self.max_cols:
                                    target_cell_id = f"{chr(64 + target_col)}{target_row}"
                                    
                                    # Update cell if it exists
                                    if target_cell_id in self.cells:
                                        # Clear existing content
                                        self.cells[target_cell_id].delete(0, tk.END)
                                        # Insert new content
                                        self.cells[target_cell_id].insert(0, cell_data.strip())
                    else:
                        # Single cell paste
                        current_cell.delete(0, tk.END)
                        current_cell.insert(0, clipboard_content)
                
                return "break"  # Prevent default behavior
            except Exception as e:
                print(f"Paste error: {e}")
        
        return None
    
    def update_spreadsheet_scroll_regions(self):
        """Update scroll regions for spreadsheet canvases"""
        # Update canvas scroll region
        self.grid_frame.update_idletasks()
        self.header_frame.update_idletasks()
        
        # Configure scroll regions for both canvases
        self.canvas.config(scrollregion=self.canvas.bbox("all"))
        self.header_canvas.config(scrollregion=(0, 0, self.grid_frame.winfo_width(), self.header_frame.winfo_height()))
        
        # Make sure header canvas width matches grid width
        self.header_canvas.itemconfig(self.header_window, width=self.grid_frame.winfo_width())
    
    def add_spreadsheet_row(self):
        """Add a new row to the spreadsheet"""
        theme = self.themes[self.current_theme]
        cell_width = 10
        row_header_width = 4
        
        # Increment max rows
        new_row = self.max_rows
        self.max_rows += 1
        
        # Configure row height
        self.grid_frame.rowconfigure(new_row, minsize=25)  # Consistent row height
        
        # Add row header
        row_label = tk.Label(self.grid_frame, text=str(new_row), width=row_header_width, height=1,
                           bg=theme["header_bg"], fg=theme["text_fg"],
                           relief='raised', borderwidth=1)
        row_label.grid(row=new_row, column=0, sticky='nsew')
        
        # Add cells for the new row
        for col in range(1, self.max_cols):
            cell = tk.Entry(self.grid_frame, width=cell_width, bg=theme["text_bg"], 
                          fg=theme["text_fg"], borderwidth=1, relief='solid')
            cell.grid(row=new_row, column=col, sticky='nsew')
            cell.insert(0, "")  # Empty by default
            
            # Store cell reference
            cell_id = f"{chr(64 + col)}{new_row}"  # e.g., A101, B101, etc.
            self.cells[cell_id] = cell
            
            # Bind events for cell navigation and clipboard
            cell.bind("<Return>", lambda e, r=new_row, c=col: self.move_to_cell(r+1, c))
            cell.bind("<Tab>", lambda e, r=new_row, c=col: self.move_to_cell(r, c+1))
            cell.bind("<Shift-Tab>", lambda e, r=new_row, c=col: self.move_to_cell(r, c-1))
            cell.bind("<Up>", lambda e, r=new_row, c=col: self.move_to_cell(r-1, c))
            cell.bind("<Down>", lambda e, r=new_row, c=col: self.move_to_cell(r+1, c))
            cell.bind("<Left>", lambda e, r=new_row, c=col: self.move_to_cell(r, c-1))
            cell.bind("<Right>", lambda e, r=new_row, c=col: self.move_to_cell(r, c+1))
            cell.bind("<Control-v>", self.paste_to_cells)
            cell.bind("<Control-c>", self.copy_from_cells)
        
        # Update scroll regions
        self.update_spreadsheet_scroll_regions()
    
    def add_spreadsheet_column(self):
        """Add a new column to the spreadsheet"""
        theme = self.themes[self.current_theme]
        cell_width = 10
        
        # Increment max columns
        new_col = self.max_cols
        self.max_cols += 1
        
        # Configure column width for grid frame
        self.grid_frame.columnconfigure(new_col, minsize=cell_width*8)  # Consistent width
        
        # Add column header
        col_letter = chr(64 + new_col) if new_col <= 26 else chr(64 + (new_col // 26)) + chr(64 + (new_col % 26))
        col_label = tk.Label(self.header_frame, text=col_letter, width=cell_width, 
                           bg=theme["header_bg"], fg=theme["text_fg"],
                           relief='raised', borderwidth=1)
        col_label.grid(row=0, column=new_col-1, sticky='nsew')  # Column index starts at 0 in header frame
        self.header_frame.columnconfigure(new_col-1, minsize=cell_width*8)  # Consistent width
        
        # Add cells for the new column
        for row in range(1, self.max_rows):
            cell = tk.Entry(self.grid_frame, width=cell_width, bg=theme["text_bg"], 
                          fg=theme["text_fg"], borderwidth=1, relief='solid')
            cell.grid(row=row, column=new_col, sticky='nsew')
            cell.insert(0, "")  # Empty by default
            
            # Store cell reference
            cell_id = f"{col_letter}{row}"  # e.g., AA1, AA2, etc.
            self.cells[cell_id] = cell
            
            # Bind events for cell navigation and clipboard
            cell.bind("<Return>", lambda e, r=row, c=new_col: self.move_to_cell(r+1, c))
            cell.bind("<Tab>", lambda e, r=row, c=new_col: self.move_to_cell(r, c+1))
            cell.bind("<Shift-Tab>", lambda e, r=row, c=new_col: self.move_to_cell(r, c-1))
            cell.bind("<Up>", lambda e, r=row, c=new_col: self.move_to_cell(r-1, c))
            cell.bind("<Down>", lambda e, r=row, c=new_col: self.move_to_cell(r+1, c))
            cell.bind("<Left>", lambda e, r=row, c=new_col: self.move_to_cell(r, c-1))
            cell.bind("<Right>", lambda e, r=row, c=new_col: self.move_to_cell(r, c+1))
            cell.bind("<Control-v>", self.paste_to_cells)
            cell.bind("<Control-c>", self.copy_from_cells)
        
        # Update scroll regions
        self.update_spreadsheet_scroll_regions()
    
    def create_line_numbers(self):
        """Create line numbers for code mode"""
        # Create line numbers frame
        self.line_numbers_frame = tk.Frame(self.text_container, bg='#2d2d30', width=50)
        self.line_numbers_frame.pack(side='left', fill='y', before=self.text_area)
        
        # Create canvas for line numbers
        self.line_numbers_canvas = tk.Canvas(self.line_numbers_frame, 
                                           bg='#2d2d30', 
                                           highlightthickness=0,
                                           width=50)
        self.line_numbers_canvas.pack(fill='both', expand=True)
        
        # Synchronize line numbers with text scrolling
        self.text_area.bind("<<Modified>>", self.update_line_numbers)
        self.text_area.bind("<Configure>", self.update_line_numbers)
        self.text_area.bind("<MouseWheel>", self.update_line_numbers)
        
        # Update line numbers
        self.update_line_numbers()
    
    def remove_line_numbers(self):
        """Remove line numbers when switching back to normal mode"""
        if self.line_numbers_frame:
            self.line_numbers_frame.destroy()
            self.line_numbers_frame = None
            self.line_numbers_canvas = None
    
    def update_line_numbers(self, event=None):
        """Update line numbers display"""
        if not self.is_code_mode or not self.line_numbers_canvas:
            return
        
        self.line_numbers_canvas.delete('all')
        
        # Get visible lines
        first_line = int(self.text_area.index('@0,0').split('.')[0])
        last_line = int(self.text_area.index(f'@0,{self.text_area.winfo_height()}').split('.')[0])
        
        # Get text widget font and calculate line height
        font_obj = font.Font(font=self.text_area['font'])
        line_height = font_obj.metrics('linespace')
        
        # Draw line numbers aligned with text
        for line_num in range(first_line, last_line + 1):
            # Get y-coordinate of the line in the text widget
            dline = self.text_area.dlineinfo(f"{line_num}.0")
            if dline:  # Line is visible
                y = dline[1]  # y-coordinate of the line
                self.line_numbers_canvas.create_text(
                    45,  # x position (right-aligned)
                    y + line_height/2,  # y position (centered vertically with line)
                    text=str(line_num),
                    fill='#858585',
                    font=('Consolas', 9),
                    anchor='e'
                )
    
    def on_language_change(self, event=None):
        """Handle language change from dropdown"""
        self.current_language = self.language_var.get()
        self.apply_syntax_highlighting()
        self.status_bar.configure(text=f"Language changed to {self.current_language}")
        self.root.title(f"Code ({self.current_language}) - Modern Notepad")
    
    def apply_syntax_highlighting(self):
        """Apply syntax highlighting to the text"""
        if not self.is_code_mode:
            return
        
        content = self.text_area.get('1.0', tk.END)
        
        # Clear existing tags
        for tag in ['keyword', 'string', 'comment', 'number', 'function']:
            self.text_area.tag_remove(tag, '1.0', tk.END)
        
        # Get keywords based on selected language
        keywords = self.get_language_keywords()
        
        lines = content.split('\n')
        for line_num, line in enumerate(lines, 1):
            # Highlight keywords
            for keyword in keywords:
                pattern = r'\b' + re.escape(keyword) + r'\b'
                for match in re.finditer(pattern, line):
                    start = f"{line_num}.{match.start()}"
                    end = f"{line_num}.{match.end()}"
                    self.text_area.tag_add('keyword', start, end)
            
            # Highlight strings
            string_patterns = [r'".*?"', r"'.*?'"]
            for pattern in string_patterns:
                for match in re.finditer(pattern, line):
                    start = f"{line_num}.{match.start()}"
                    end = f"{line_num}.{match.end()}"
                    self.text_area.tag_add('string', start, end)
            
            # Highlight comments
            comment_match = re.search(r'#.*$', line)
            if comment_match:
                start = f"{line_num}.{comment_match.start()}"
                end = f"{line_num}.{comment_match.end()}"
                self.text_area.tag_add('comment', start, end)
            
            # Highlight numbers
            for match in re.finditer(r'\b\d+\.?\d*\b', line):
                start = f"{line_num}.{match.start()}"
                end = f"{line_num}.{match.end()}"
                self.text_area.tag_add('number', start, end)
            
            # Highlight function definitions
            func_match = re.search(r'def\s+(\w+)', line)
            if func_match:
                start = f"{line_num}.{func_match.start(1)}"
                end = f"{line_num}.{func_match.end(1)}"
                self.text_area.tag_add('function', start, end)
    
    def clear_syntax_highlighting(self):
        """Clear all syntax highlighting"""
        for tag in ['keyword', 'string', 'comment', 'number', 'function']:
            self.text_area.tag_remove(tag, '1.0', tk.END)
    
    def get_language_keywords(self):
        """Get keywords for the selected programming language"""
        language_keywords = {
            "Python": ['def', 'class', 'if', 'else', 'elif', 'for', 'while', 'try', 'except', 
                      'finally', 'with', 'as', 'import', 'from', 'return', 'yield', 'break', 
                      'continue', 'pass', 'and', 'or', 'not', 'in', 'is', 'lambda', 'True', 
                      'False', 'None', 'self', 'super', 'print'],
            
            "JavaScript": ['function', 'var', 'let', 'const', 'if', 'else', 'for', 'while', 'do', 
                         'switch', 'case', 'default', 'break', 'continue', 'return', 'try', 
                         'catch', 'finally', 'throw', 'new', 'this', 'typeof', 'instanceof', 
                         'null', 'undefined', 'true', 'false', 'class', 'extends', 'super', 
                         'import', 'export', 'async', 'await'],
            
            "Java": ['abstract', 'assert', 'boolean', 'break', 'byte', 'case', 'catch', 'char', 
                    'class', 'const', 'continue', 'default', 'do', 'double', 'else', 'enum', 
                    'extends', 'final', 'finally', 'float', 'for', 'if', 'implements', 'import', 
                    'instanceof', 'int', 'interface', 'long', 'native', 'new', 'package', 'private', 
                    'protected', 'public', 'return', 'short', 'static', 'strictfp', 'super', 'switch', 
                    'synchronized', 'this', 'throw', 'throws', 'transient', 'try', 'void', 'volatile', 'while'],
            
            "C++": ['auto', 'break', 'case', 'char', 'class', 'const', 'continue', 'default', 
                   'delete', 'do', 'double', 'else', 'enum', 'extern', 'float', 'for', 'friend', 
                   'goto', 'if', 'inline', 'int', 'long', 'namespace', 'new', 'operator', 'private', 
                   'protected', 'public', 'register', 'return', 'short', 'signed', 'sizeof', 'static', 
                   'struct', 'switch', 'template', 'this', 'throw', 'try', 'typedef', 'union', 
                   'unsigned', 'virtual', 'void', 'volatile', 'while'],
            
            "SQL": ['SELECT', 'FROM', 'WHERE', 'INSERT', 'UPDATE', 'DELETE', 'CREATE', 'ALTER', 
                   'DROP', 'TABLE', 'DATABASE', 'VIEW', 'INDEX', 'JOIN', 'INNER', 'LEFT', 'RIGHT', 
                   'FULL', 'OUTER', 'ON', 'GROUP', 'BY', 'HAVING', 'ORDER', 'ASC', 'DESC', 'LIMIT', 
                   'OFFSET', 'UNION', 'ALL', 'AS', 'DISTINCT', 'INTO', 'VALUES', 'SET', 'CONSTRAINT', 
                   'PRIMARY', 'KEY', 'FOREIGN', 'REFERENCES', 'NOT', 'NULL', 'DEFAULT', 'AUTO_INCREMENT']
        }
        
        # Return keywords for selected language or Python as fallback
        return language_keywords.get(self.current_language, language_keywords["Python"])
    
    def on_key_release(self, event=None):
        """Handle key release events"""
        self.update_status()
        if self.is_code_mode:
            self.apply_syntax_highlighting()
            self.update_line_numbers()
    
    def on_click(self, event=None):
        """Handle click events"""
        self.update_status()
        if self.is_code_mode:
            self.update_line_numbers()
    
    def change_theme(self, theme_name):
        """Change the application theme"""
        if theme_name == "system":
            self.themes["system"] = self.get_system_theme()
        
        self.current_theme = theme_name
        self.setup_styles()
        self.apply_theme_to_text_area()
        self.setup_syntax_highlighting()
        
        # Update context menu colors
        self.create_context_menu()
        
        # Re-apply syntax highlighting if in code mode
        if self.is_code_mode:
            self.apply_syntax_highlighting()
        
        self.status_bar.configure(text=f"Theme changed to {theme_name.title()}")
    
    def apply_theme_to_text_area(self):
        """Apply current theme to text area"""
        theme = self.themes[self.current_theme]
        self.text_area.configure(
            bg=theme["text_bg"],
            fg=theme["text_fg"],
            insertbackground=theme["text_fg"],
            selectbackground=theme["select_bg"]
        )
        
        # Update text container background
        self.text_container.configure(bg=theme["bg"])
        
        # Update line numbers if in code mode
        if self.is_code_mode and self.line_numbers_canvas:
            line_bg = "#2d2d30" if self.current_theme != "light" else "#f5f5f5"
            self.line_numbers_canvas.configure(bg=line_bg)
            self.line_numbers_frame.configure(bg=line_bg)
            self.update_line_numbers()
    
    def toggle_bold(self):
        """Toggle bold formatting for selected text"""
        try:
            if self.text_area.tag_ranges(tk.SEL):
                # Store selection range
                sel_start = self.text_area.index(tk.SEL_FIRST)
                sel_end = self.text_area.index(tk.SEL_LAST)
                
                current_tags = self.text_area.tag_names(tk.SEL_FIRST)
                if "bold" in current_tags:
                    self.text_area.tag_remove("bold", sel_start, sel_end)
                    self.current_formatting['bold'] = False
                else:
                    self.text_area.tag_add("bold", sel_start, sel_end)
                    self.text_area.tag_configure("bold", font=("Segoe UI", 12, "bold"))
                    self.current_formatting['bold'] = True
                
                # Restore selection
                self.text_area.tag_add(tk.SEL, sel_start, sel_end)
                self.text_area.mark_set(tk.INSERT, sel_end)
                
                # Update button states based on selection
                self.update_current_formatting()
            else:
                # Toggle for future text input
                self.current_formatting['bold'] = not self.current_formatting['bold']
                self.update_button_states()
        except tk.TclError:
            pass
    
    def toggle_italic(self):
        """Toggle italic formatting for selected text"""
        try:
            if self.text_area.tag_ranges(tk.SEL):
                # Store selection range
                sel_start = self.text_area.index(tk.SEL_FIRST)
                sel_end = self.text_area.index(tk.SEL_LAST)
                
                current_tags = self.text_area.tag_names(tk.SEL_FIRST)
                if "italic" in current_tags:
                    self.text_area.tag_remove("italic", sel_start, sel_end)
                    self.current_formatting['italic'] = False
                else:
                    self.text_area.tag_add("italic", sel_start, sel_end)
                    self.text_area.tag_configure("italic", font=("Segoe UI", 12, "italic"))
                    self.current_formatting['italic'] = True
                
                # Restore selection
                self.text_area.tag_add(tk.SEL, sel_start, sel_end)
                self.text_area.mark_set(tk.INSERT, sel_end)
                
                # Update button states based on selection
                self.update_current_formatting()
            else:
                # Toggle for future text input
                self.current_formatting['italic'] = not self.current_formatting['italic']
                self.update_button_states()
        except tk.TclError:
            pass
    
    def toggle_underline(self):
        """Toggle underline formatting for selected text"""
        try:
            if self.text_area.tag_ranges(tk.SEL):
                # Store selection range
                sel_start = self.text_area.index(tk.SEL_FIRST)
                sel_end = self.text_area.index(tk.SEL_LAST)
                
                current_tags = self.text_area.tag_names(tk.SEL_FIRST)
                if "underline" in current_tags:
                    self.text_area.tag_remove("underline", sel_start, sel_end)
                    self.current_formatting['underline'] = False
                else:
                    self.text_area.tag_add("underline", sel_start, sel_end)
                    self.text_area.tag_configure("underline", underline=True)
                    self.current_formatting['underline'] = True
                
                # Restore selection
                self.text_area.tag_add(tk.SEL, sel_start, sel_end)
                self.text_area.mark_set(tk.INSERT, sel_end)
                
                # Update button states based on selection
                self.update_current_formatting()
            else:
                # Toggle for future text input
                self.current_formatting['underline'] = not self.current_formatting['underline']
                self.update_button_states()
        except tk.TclError:
            pass
    
    def change_font_size(self, event=None):
        """Change font size for selected text"""
        try:
            size = int(self.size_var.get())
            self.current_formatting['size'] = size
            
            if self.text_area.tag_ranges(tk.SEL):
                # Store selection range
                sel_start = self.text_area.index(tk.SEL_FIRST)
                sel_end = self.text_area.index(tk.SEL_LAST)
                
                tag_name = f"size_{size}"
                self.text_area.tag_add(tag_name, sel_start, sel_end)
                self.text_area.tag_configure(tag_name, font=("Segoe UI", size))
                
                # Restore selection
                self.text_area.tag_add(tk.SEL, sel_start, sel_end)
                self.text_area.mark_set(tk.INSERT, sel_end)
        except (tk.TclError, ValueError):
            pass
    
    def change_text_color(self):
        """Change text color for selected text"""
        try:
            from tkinter import colorchooser
            color = colorchooser.askcolor(title="Choose text color")
            if color[1]:  # If a color was selected
                self.current_formatting['color'] = color[1]
                tag_name = f"color_{color[1].replace('#', '')}"
                
                if self.text_area.tag_ranges(tk.SEL):
                    # Store selection range
                    sel_start = self.text_area.index(tk.SEL_FIRST)
                    sel_end = self.text_area.index(tk.SEL_LAST)
                    
                    self.text_area.tag_add(tag_name, sel_start, sel_end)
                    self.text_area.tag_configure(tag_name, foreground=color[1])
                    
                    # Restore selection
                    self.text_area.tag_add(tk.SEL, sel_start, sel_end)
                    self.text_area.mark_set(tk.INSERT, sel_end)
        except tk.TclError:
            pass
    
    def change_highlight_color(self):
        """Change highlight color for selected text"""
        try:
            from tkinter import colorchooser
            color = colorchooser.askcolor(title="Choose highlight color")
            if color[1]:  # If a color was selected
                self.current_formatting['highlight'] = color[1]
                tag_name = f"highlight_{color[1].replace('#', '')}"
                
                if self.text_area.tag_ranges(tk.SEL):
                    # Store selection range
                    sel_start = self.text_area.index(tk.SEL_FIRST)
                    sel_end = self.text_area.index(tk.SEL_LAST)
                    
                    self.text_area.tag_add(tag_name, sel_start, sel_end)
                    self.text_area.tag_configure(tag_name, background=color[1])
                    
                    # Restore selection
                    self.text_area.tag_add(tk.SEL, sel_start, sel_end)
                    self.text_area.mark_set(tk.INSERT, sel_end)
        except tk.TclError:
            pass
     
    def update_current_formatting(self, event=None):
        """Update current formatting state based on cursor position"""
        try:
            cursor_pos = self.text_area.index(tk.INSERT)
            tags_at_cursor = self.text_area.tag_names(cursor_pos)
            
            # Update formatting state based on tags at cursor
            self.current_formatting['bold'] = 'bold' in tags_at_cursor
            self.current_formatting['italic'] = 'italic' in tags_at_cursor
            self.current_formatting['underline'] = 'underline' in tags_at_cursor
            
            # Check for size tags
            for tag in tags_at_cursor:
                if tag.startswith('size_'):
                    try:
                        size = int(tag.split('_')[1])
                        self.current_formatting['size'] = size
                        self.size_var.set(str(size))
                    except (ValueError, IndexError):
                        pass
            
            # Check for color tags
            for tag in tags_at_cursor:
                if tag.startswith('color_'):
                    self.current_formatting['color'] = f"#{tag.split('_')[1]}"
                elif tag.startswith('highlight_'):
                    self.current_formatting['highlight'] = f"#{tag.split('_')[1]}"
            
            self.update_button_states()
        except tk.TclError:
            pass
     
    def update_button_states(self):
        """Update button appearance based on current formatting state"""
        try:
            # Update button styles to show active state
            bold_style = 'Active.TButton' if self.current_formatting['bold'] else 'Header.TButton'
            italic_style = 'Active.TButton' if self.current_formatting['italic'] else 'Header.TButton'
            underline_style = 'Active.TButton' if self.current_formatting['underline'] else 'Header.TButton'
            
            self.bold_button.configure(style=bold_style)
            self.italic_button.configure(style=italic_style)
            self.underline_button.configure(style=underline_style)
        except tk.TclError:
            pass
     
    def on_key_press_format(self, event):
        """Apply current formatting to newly typed characters"""
        try:
            # Only apply formatting to printable characters
            if event.char and event.char.isprintable() and not self.is_code_mode:
                # Get current cursor position
                cursor_pos = self.text_area.index(tk.INSERT)
                
                # Schedule formatting application after character is inserted
                self.root.after_idle(lambda: self.apply_formatting_to_new_char(cursor_pos))
        except Exception:
            pass
     
    def apply_formatting_to_new_char(self, start_pos):
        """Apply current formatting to the newly inserted character"""
        try:
            # Calculate end position (one character after start)
            end_pos = f"{start_pos}+1c"
            
            # Apply active formatting
            if self.current_formatting['bold']:
                self.text_area.tag_add("bold", start_pos, end_pos)
                self.text_area.tag_configure("bold", font=("Segoe UI", self.current_formatting['size'], "bold"))
            
            if self.current_formatting['italic']:
                self.text_area.tag_add("italic", start_pos, end_pos)
                self.text_area.tag_configure("italic", font=("Segoe UI", self.current_formatting['size'], "italic"))
            
            if self.current_formatting['underline']:
                self.text_area.tag_add("underline", start_pos, end_pos)
                self.text_area.tag_configure("underline", underline=True)
            
            if self.current_formatting['size'] != 12:
                tag_name = f"size_{self.current_formatting['size']}"
                self.text_area.tag_add(tag_name, start_pos, end_pos)
                self.text_area.tag_configure(tag_name, font=("Segoe UI", self.current_formatting['size']))
            
            if self.current_formatting['color']:
                tag_name = f"color_{self.current_formatting['color'].replace('#', '')}"
                self.text_area.tag_add(tag_name, start_pos, end_pos)
                self.text_area.tag_configure(tag_name, foreground=self.current_formatting['color'])
            
            if self.current_formatting['highlight']:
                tag_name = f"highlight_{self.current_formatting['highlight'].replace('#', '')}"
                self.text_area.tag_add(tag_name, start_pos, end_pos)
                self.text_area.tag_configure(tag_name, background=self.current_formatting['highlight'])
        except tk.TclError:
            pass
    
    def handle_text_replacement(self, event):
        """Handle typing over selected text"""
        try:
            # Check if there's a selection and the key is a printable character
            if (self.text_area.tag_ranges(tk.SEL) and 
                len(event.char) == 1 and event.char.isprintable() and 
                event.keysym not in ['BackSpace', 'Delete', 'Return', 'Tab']):
                
                # Get the formatting from the selected text
                self.inherit_formatting_from_selection()
                
                # Delete the selected text
                self.text_area.delete(tk.SEL_FIRST, tk.SEL_LAST)
                
                # Insert the new character with inherited formatting
                insert_pos = self.text_area.index(tk.INSERT)
                self.text_area.insert(insert_pos, event.char)
                
                # Apply formatting to the new character
                self.apply_formatting_to_new_char(insert_pos)
                
                # Prevent the default behavior
                return "break"
                
        except tk.TclError:
            pass
        
        return None
    
    def handle_backspace(self, event):
        """Handle backspace key when text is selected"""
        try:
            if self.text_area.tag_ranges(tk.SEL):
                # Get formatting from selection before deleting
                self.inherit_formatting_from_selection()
                
                # Delete the selected text
                self.text_area.delete(tk.SEL_FIRST, tk.SEL_LAST)
                
                # Prevent default backspace behavior
                return "break"
        except tk.TclError:
            pass
        
        return None
    
    def handle_delete(self, event):
        """Handle delete key when text is selected"""
        try:
            if self.text_area.tag_ranges(tk.SEL):
                # Get formatting from selection before deleting
                self.inherit_formatting_from_selection()
                
                # Delete the selected text
                self.text_area.delete(tk.SEL_FIRST, tk.SEL_LAST)
                
                # Prevent default delete behavior
                return "break"
        except tk.TclError:
            pass
        
        return None
    
    def inherit_formatting_from_selection(self):
        """Inherit formatting from the selected text"""
        try:
            if self.text_area.tag_ranges(tk.SEL):
                start_pos = self.text_area.index(tk.SEL_FIRST)
                
                # Get all tags at the start of selection
                tags = self.text_area.tag_names(start_pos)
                
                # Reset current formatting
                self.current_formatting = {
                    'bold': False,
                    'italic': False,
                    'underline': False,
                    'size': 12,
                    'color': None,
                    'highlight': None
                }
                
                # Check for formatting tags
                for tag in tags:
                    if tag == 'bold':
                        self.current_formatting['bold'] = True
                    elif tag == 'italic':
                        self.current_formatting['italic'] = True
                    elif tag == 'underline':
                        self.current_formatting['underline'] = True
                    elif tag.startswith('size_'):
                        try:
                            size = int(tag.split('_')[1])
                            self.current_formatting['size'] = size
                            self.size_var.set(str(size))
                        except (ValueError, IndexError):
                            pass
                    elif tag.startswith('color_'):
                        color = '#' + tag.split('_')[1]
                        self.current_formatting['color'] = color
                    elif tag.startswith('highlight_'):
                        color = '#' + tag.split('_')[1]
                        self.current_formatting['highlight'] = color
                
                # Update button states
                self.update_button_states()
                
        except tk.TclError:
            pass
    
    def on_closing(self):
        """Handle window closing"""
        if self.check_unsaved_changes():
            self.root.destroy()
    
    def run(self):
        """Start the application"""
        self.root.mainloop()

if __name__ == "__main__":
    app = ModernNotepad()
    app.run()