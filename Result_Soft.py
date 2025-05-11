import tkinter as tk
from tkinter import filedialog, messagebox
import re
import pandas as pd
import os
import sys

class CBSEParserApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("CBSE Result")

        # Center window on screen
        window_width = 600
        window_height = 400
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = int((screen_width / 2) - (window_width / 2))
        y = int((screen_height / 2) - (window_height / 2))
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.resizable(False, False)

        # Load icon safely (for PyInstaller compatibility)
        def resource_path(relative_path):
            try:
                base_path = sys._MEIPASS
            except Exception:
                base_path = os.path.abspath(".")
            return os.path.join(base_path, relative_path)

        try:
            self.iconbitmap(resource_path("cbse_TNC_icon.ico"))
        except Exception as e:
            print("Icon load failed:", e)

        self.file_path = None
        self.sample_line = ""
        self.grade12 = False
        self.subject_count = 6
        self.init_frame1()

    def init_frame1(self):
        self.frame1 = tk.Frame(self)
        self.frame1.pack(fill='both', expand=True)

        tk.Label(self.frame1, text="Step 1: Upload TXT File", font=("Arial", 14)).pack(pady=10)
        tk.Button(self.frame1, text="Choose File", command=self.select_file,
                  font=("Tahoma", 12), width=15, bd=0, bg="pink").pack(pady=5)

        self.file_label = tk.Label(self.frame1, text="No file selected", font=("Arial", 10))
        self.file_label.pack(pady=5)

        tk.Label(self.frame1, text="Step 2: Paste Sample Student Line", font=("Arial", 14)).pack(pady=10)
        self.sample_entry = tk.Text(self.frame1, height=4, width=70)
        self.sample_entry.pack(pady=5)

        tk.Button(self.frame1, text="Next", command=self.handle_sample,
                  font=("Tahoma", 12), width=15, bd=0, bg="lightgreen").pack(pady=10)

    def select_file(self):
        path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
        if path:
            self.file_path = path
            self.file_label.config(text=os.path.basename(path))

    def handle_sample(self):
        self.sample_line = self.sample_entry.get("1.0", tk.END).strip()
        if not self.file_path or not self.sample_line:
            messagebox.showerror("Error", "Please select a file and enter a sample line.")
            return

        grade12_match = re.search(r'(A[1-2]|B[1-2]|C[1-2]|D[1-2])\s+(A[1-2]|B[1-2]|C[1-2]|D[1-2])\s+(A[1-2]|B[1-2]|C[1-2]|D[1-2])\s+(PASS|FAIL|COMP|ABST)', self.sample_line)
        self.grade12 = bool(grade12_match)
        self.subject_count = 5 if self.grade12 else 6

        self.frame1.destroy()
        self.init_frame2()

    def init_frame2(self):
        self.frame2 = tk.Frame(self)
        self.frame2.pack(fill='both', expand=True)

        tk.Label(self.frame2, text="Step 3: Click to Save Excel File", font=("Arial", 14)).pack(pady=40)

        tk.Button(self.frame2, text="Save and Generate Excel", command=self.generate_excel,
                  font=("Tahoma", 12), width=25, bd=0, bg="lightblue").pack(pady=20)

    def generate_excel(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Save Excel File As"
        )

        if not file_path:
            return  # User cancelled

        try:
            self.parse_and_save(file_path)
            response = messagebox.askyesno("Success", f"✅ File saved as:\n{file_path}\n\nDo you want to generate another file?")
            if response:
                self.frame2.destroy()
                self.init_frame1()
            else:
                self.frame2.destroy()
                self.init_final_frame()
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def parse_and_save(self, filename):
        with open(self.file_path, 'r', encoding='utf-8', errors='ignore') as f:
            lines = f.readlines()

        skip_patterns = [
            r'^DATE:-', r'^ROLL\s+F', r'^NO\s+L', r'^-+$',
            r'^SCHOOL\s+:\s+-', r'^TOTAL CANDIDATES', r'^\f', r'^$', r'^REGION:'
        ]

        data = []
        i = 0

        while i < len(lines) - 1:
            line1 = lines[i].strip()
            line2 = lines[i + 1].strip()

            if any(re.match(p, line1) for p in skip_patterns) or not re.match(r'^\d{8}', line1):
                i += 1
                continue

            match = re.match(r'^(\d{8})\s+([MF])\s+(.*?)\s+(\d{3})', line1)
            if not match:
                i += 2
                continue

            roll_no = match.group(1)
            gender = match.group(2)
            name = match.group(3).strip()
            subject_codes = re.findall(r'\b\d{3}\b', line1)[:self.subject_count]

            if self.grade12:
                result_match = re.search(r'(A[1-2]|B[1-2]|C[1-2]|D[1-2])\s+(A[1-2]|B[1-2]|C[1-2]|D[1-2])\s+(A[1-2]|B[1-2]|C[1-2]|D[1-2])\s+(PASS|FAIL|COMP|ABST)', line1)
                if not result_match:
                    i += 2
                    continue
                isg1, isg2, isg3, result = result_match.groups()
            else:
                result_match = re.search(r'(PASS|FAIL|COMP|ABST)', line1)
                result = result_match.group(1) if result_match else ""
                isg1 = isg2 = isg3 = ""

            marks_grades = re.findall(r'(\d{2,3})\s+([A-D][1-2])', line2)
            if len(subject_codes) < self.subject_count or len(marks_grades) < self.subject_count:
                i += 2
                continue

            row = [roll_no, gender, name] + subject_codes
            if self.grade12:
                row += [isg1, isg2, isg3]
            for mark, grade in marks_grades[:self.subject_count]:
                row += [mark, grade]
            row.append(result)

            data.append(row)
            i += 2

        columns = ['Roll No', 'Gender', 'Name'] + [f'Sub{i+1}' for i in range(self.subject_count)]
        if self.grade12:
            columns += [f'ISG{i+1}' for i in range(3)]
        columns += sum([[f'Marks{i+1}', f'Grade{i+1}'] for i in range(self.subject_count)], [])
        columns += ['Result']

        df = pd.DataFrame(data, columns=columns)
        df.to_excel(filename, index=False)

    def init_final_frame(self):
        self.final_frame = tk.Frame(self)
        self.final_frame.pack(fill='both', expand=True)

        tk.Label(self.final_frame, text="🎉 Thank you for using the CBSE Result Parser!", font=("Arial", 16), fg='green').pack(pady=60)
        tk.Button(self.final_frame, text="Exit", command=self.destroy,
                  font=("Tahoma", 12), width=20, bd=0, bg="red", fg="white").pack(pady=10)

# Run the app
if __name__ == "__main__":
    app = CBSEParserApp()
    app.mainloop()
