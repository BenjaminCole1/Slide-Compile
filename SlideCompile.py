import tkinter as tk
from tkinter import filedialog, messagebox
from pptx import Presentation
from pptx.util import Inches

class SlideCompileEditor:

    def __init__(self, root):
        self.root = root
        self.root.title("SlideCompile")
        
        self.file_path = None

        self.text_area = tk.Text(self.root, wrap='word')
        self.text_area.pack(fill='both', expand=True)
        
        self.menu_bar = tk.Menu(self.root)
        self.file_menu = tk.Menu(self.menu_bar, tearoff=0)
        
        self.file_menu.add_command(label="Open", command=self.open_file)
        self.file_menu.add_command(label="Save", command=self.save_file)
        self.file_menu.add_command(label="Save As", command=self.save_as_file)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Exit", command=self.root.quit)
        
        self.menu_bar.add_cascade(label="File", menu=self.file_menu)
        self.root.config(menu=self.menu_bar)
        
        self.root.bind_all('<Control-s>', self.handle_save_shortcut)

        # Add Compile button
        self.compile_button = tk.Button(self.root, text="Compile", command=self.compile)
        self.compile_button.place(relx=1.0, rely=1.0, anchor='se', bordermode='outside')

    def handle_save_shortcut(self, event):
        self.save_file()
        
    def open_file(self):
        self.file_path = filedialog.askopenfilename(defaultextension=".txt", 
                                               filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if self.file_path:
            try:
                with open(self.file_path, "r") as file:
                    content = file.read()
                    self.text_area.delete(1.0, tk.END)
                    self.text_area.insert(tk.INSERT, content)
                    self.root.title(f"SlideCompile - {self.file_path}")
            except Exception as e:
                self.show_error(f"Could not open file: {e}")

    def save_file(self):
        if self.file_path:
            try:
                with open(self.file_path, "w") as file:
                    content = self.text_area.get(1.0, tk.END)
                    file.write(content)
                    self.root.title(f"SlideCompile - {self.file_path}")
            except Exception as e:
                self.show_error(f"Could not save file: {e}")
        else:
            self.save_as_file()

    def save_as_file(self):
        self.file_path = filedialog.asksaveasfilename(defaultextension=".txt", 
                                                      filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if self.file_path:
            try:
                with open(self.file_path, "w") as file:
                    content = self.text_area.get(1.0, tk.END)
                    file.write(content)
                    self.root.title(f"SlideCompile - {self.file_path}")
            except Exception as e:
                self.show_error(f"Could not save file: {e}")

    def add_slide(self, prs, format_number, title, content, line):
        try:
            slide_layout = prs.slide_layouts[format_number-1]
            slide = prs.slides.add_slide(slide_layout)

            if title:
                slide.shapes.title.text = title

            if content:
                bullet_points = content.split('\n')
                content_box = slide.shapes.placeholders[1]
                content_frame = content_box.text_frame
                for point in bullet_points:
                    p = content_frame.add_paragraph()
                    p.text = point
        except Exception as e:
            self.show_error(f"ERROR: {e}. Error line: {line}")
            return True
        return False

    def compile(self):
        self.save_file()
        print("Reading text")
        with open(self.file_path) as file:
            print(file)
            lines = file.readlines()

        current_slide_data = {'format_number': -1, 'title': '', 'content': ''}
        print(current_slide_data)

        prs = Presentation()
        print("made presentation file")
        left = top = Inches(1)

        for lineNumber, line in enumerate(lines, start=1):
            line = line.strip()
            print(line)

            print("parsing data")

            if line.lower() == "newslide" or line.lower() == "new_slide":
                if current_slide_data['format_number'] != -1:
                    errored = self.add_slide(prs, current_slide_data['format_number'], current_slide_data['title'], current_slide_data['content'], lineNumber)
                    if errored:
                        return
                current_slide_data['content'] = ''
            elif line.lower().startswith("formatnumber:") or line.lower().startswith("format_number"):
                _, format_number = line.split(':', 1)
                current_slide_data['format_number'] = int(format_number.strip())
            elif line.lower().startswith("title:"):
                _, title = line.split(':', 1)
                current_slide_data['title'] = title.strip()
            elif line.lower().startswith("content:"):
                _, content = line.split(':', 1)
                current_slide_data['content'] += content.strip() + '\n'
            else:
                print(f"error, didn't find any useable data at {line}")

        if current_slide_data['format_number'] != -1:
            errored = self.add_slide(prs, current_slide_data['format_number'], current_slide_data['title'], current_slide_data['content'], lineNumber)
            if errored:
                return

        file_path = filedialog.asksaveasfilename(defaultextension=".pptx")
        if file_path:
            print("Selected file path:", file_path)
            prs.save(file_path)

    def show_error(self, error_message):
        messagebox.showerror("Error", error_message)

if __name__ == "__main__":
    root = tk.Tk()
    app = SlideCompileEditor(root)
    root.mainloop()