import tkinter as tk
from tkinter import filedialog, messagebox, ttk, Listbox, MULTIPLE, END
import json
import os
import sys
import re
from docx import Document
import pdfplumber

# -------------------------------------------------------------------
# Helpers for PyInstaller bundled files
# -------------------------------------------------------------------
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# -------------------------------------------------------------------
# Extract sections from PDF
# -------------------------------------------------------------------
def extract_sections_from_pdf(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text += t + "\n"
    if not text.strip():
        return ["No extractable text – PDF may be scanned."]

    # Split into sections using common GHS headings
    section_pattern = re.compile(
        r'(?:^|\n)((?:SECTION\s+\d+|'
        r'\d{1,2}\.\s*(?:HAZARDS?\s*IDENTIFICATION|COMPOSITION|'
        r'FIRST\s*AID|FIRE\s*FIGHTING|ACCIDENTAL\s*RELEASE|'
        r'HANDLING\s*AND\s*STORAGE|EXPOSURE\s*CONTROLS|'
        r'PHYSICAL\s*AND\s*CHEMICAL|STABILITY\s*AND\s*REACTIVITY|'
        r'TOXICOLOGICAL|ECOLOGICAL|DISPOSAL|TRANSPORT|REGULATORY|OTHER\s*INFORMATION)'
        r')(?:\s*:|\s*\n|$))',
        re.IGNORECASE
    )
    splits = section_pattern.split(text)
    sections = []
    # The regex returns list: [before, heading, after, heading, after...]
    if len(splits) > 1:
        # First part is text before any heading (discard – it’s usually supplier info)
        for i in range(1, len(splits), 2):
            heading = splits[i].strip()
            body = splits[i+1].strip() if (i+1) < len(splits) else ""
            # Keep the heading + body together
            sections.append(f"{heading}\n{body}")
    else:
        # No headings found – return whole text as one block
        sections = [text]
    return sections

# -------------------------------------------------------------------
# Load configs
# -------------------------------------------------------------------
def load_json(file_name):
    path = resource_path(file_name)
    if not os.path.exists(path):
        return {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

# -------------------------------------------------------------------
# Main Application
# -------------------------------------------------------------------
class SDSAuthoringTool(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SDS Authoring Tool")
        self.geometry("1100x700")
        self.resizable(True, True)

        # Data
        self.pdf_path = None
        self.source_sections = []
        self.mappings = {}            # field_label -> (section_index, section_text)
        self.template_fields = []
        self.identities = []
        self.current_identity = None

        # Load config files
        self.load_configs()

        # Build UI
        self.create_widgets()

    def load_configs(self):
        field_cfg = load_json("template_fields.json")
        self.template_fields = field_cfg.get("fields", [])
        # Add identity placeholders if needed (they are handled separately)
        self.identities = load_json("identities.json")

    def create_widgets(self):
        # Top bar: open PDF, identity selector
        top_frame = tk.Frame(self)
        top_frame.pack(fill="x", padx=10, pady=5)

        tk.Button(top_frame, text="Open PDF", command=self.load_pdf).pack(side="left", padx=5)
        tk.Label(top_frame, text="Identity set:").pack(side="left", padx=(20,5))
        self.identity_var = tk.StringVar()
        self.identity_combo = ttk.Combobox(top_frame, textvariable=self.identity_var,
                                           state="readonly", width=30)
        self.identity_combo.pack(side="left", padx=5)
        self.identity_combo.bind("<<ComboboxSelected>>", self.on_identity_selected)
        self.populate_identity_list()

        # Main paned window: left (source sections) and right (template fields)
        main_pane = tk.PanedWindow(self, orient="horizontal")
        main_pane.pack(fill="both", expand=True, padx=10, pady=5)

        # Left: Source sections list
        left_frame = tk.LabelFrame(main_pane, text="Source PDF Sections (drag from here)")
        self.source_listbox = Listbox(left_frame, selectmode="single", width=50, height=25)
        self.source_listbox.pack(fill="both", expand=True, padx=5, pady=5)
        # Enable drag
        self.source_listbox.bind("<Button-1>", self.on_start_drag)
        self.source_listbox.bind("<B1-Motion>", self.on_drag_motion)
        self.source_listbox.bind("<ButtonRelease-1>", self.on_drop)
        main_pane.add(left_frame)

        # Right: Template fields
        right_frame = tk.LabelFrame(main_pane, text="Template Fields (drop a section here)")
        self.fields_listbox = Listbox(right_frame, selectmode="single", width=50, height=25)
        self.fields_listbox.pack(fill="both", expand=True, padx=5, pady=5)
        # Populate field list with labels
        for field in self.template_fields:
            self.fields_listbox.insert(END, f"{field['label']} [empty]")
        main_pane.add(right_frame)

        # Bottom buttons
        bottom_frame = tk.Frame(self)
        bottom_frame.pack(fill="x", padx=10, pady=5)
        tk.Button(bottom_frame, text="Clear Mappings", command=self.clear_mappings).pack(side="left", padx=5)
        tk.Button(bottom_frame, text="Export DOCX", command=self.export_docx).pack(side="right", padx=5)

        # Status bar
        self.status = tk.Label(self, text="Ready", bd=1, relief="sunken", anchor="w")
        self.status.pack(side="bottom", fill="x")

        # Internal drag state
        self._drag_data = {"x": 0, "y": 0, "item": None, "source": None}

    def populate_identity_list(self):
        names = [ident["name"] for ident in self.identities]
        self.identity_combo["values"] = names
        if names:
            self.identity_combo.current(0)
            self.current_identity = self.identities[0]

    def on_identity_selected(self, event):
        idx = self.identity_combo.current()
        if 0 <= idx < len(self.identities):
            self.current_identity = self.identities[idx]
            self.status.config(text=f"Selected identity: {self.current_identity['name']}")

    # ----- PDF handling -----
    def load_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if not path:
            return
        self.pdf_path = path
        self.status.config(text="Extracting sections from PDF...")
        self.update()
        self.source_sections = extract_sections_from_pdf(path)
        self.source_listbox.delete(0, END)
        for sec in self.source_sections:
            # Truncate long sections for display
            preview = sec[:100].replace("\n", " ") + ("..." if len(sec)>100 else "")
            self.source_listbox.insert(END, preview)
        self.clear_mappings()
        self.status.config(text=f"Loaded {len(self.source_sections)} sections from PDF.")

    # ----- Drag & Drop logic -----
    def on_start_drag(self, event):
        widget = event.widget
        if widget == self.source_listbox:
            idx = widget.nearest(event.y)
            if idx >= 0:
                self._drag_data["item"] = idx
                self._drag_data["source"] = "source"
                self._drag_data["x"] = event.x
                self._drag_data["y"] = event.y

    def on_drag_motion(self, event):
        # No visual feedback needed – simple drag implementation
        pass

    def on_drop(self, event):
        if not self._drag_data["item"] is not None:
            return
        # Determine target widget
        target = event.widget
        if target == self.fields_listbox and self._drag_data["source"] == "source":
            # Get target field index
            target_idx = target.nearest(event.y)
            if target_idx >= 0 and target_idx < len(self.template_fields):
                source_idx = self._drag_data["item"]
                section_text = self.source_sections[source_idx]
                field_label = self.template_fields[target_idx]["label"]
                # Store mapping
                self.mappings[field_label] = section_text
                # Update display
                self.fields_listbox.delete(target_idx)
                preview = section_text[:80].replace("\n", " ") + "..."
                self.fields_listbox.insert(target_idx, f"{field_label} [mapped] – {preview}")
                self.status.config(text=f"Mapped section {source_idx+1} → {field_label}")
        # Reset drag data
        self._drag_data = {"x": 0, "y": 0, "item": None, "source": None}

    def clear_mappings(self):
        self.mappings = {}
        self.fields_listbox.delete(0, END)
        for field in self.template_fields:
            self.fields_listbox.insert(END, f"{field['label']} [empty]")

    # ----- Export -----
    def export_docx(self):
        if not self.pdf_path:
            messagebox.showerror("Error", "No PDF loaded.")
            return
        output_path = filedialog.asksaveasfilename(defaultextension=".docx",
                                                   filetypes=[("Word Document", "*.docx")])
        if not output_path:
            return

        template_path = resource_path("template.docx")
        if not os.path.exists(template_path):
            messagebox.showerror("Error", "template.docx not found.")
            return

        doc = Document(template_path)

        # Prepare replacement dict
        replacements = {}

        # Identity placeholders
        if self.current_identity:
            id_cfg = load_json("template_fields.json").get("identity_placeholders", {})
            for key, placeholder in id_cfg.items():
                value = self.current_identity.get(key, "")
                replacements[placeholder] = value

        # Section placeholders from mappings
        for field in self.template_fields:
            label = field["label"]
            placeholder = "{{" + field["placeholder"] + "}}"
            if label in self.mappings:
                replacements[placeholder] = self.mappings[label]
            else:
                replacements[placeholder] = ""   # leave blank if not mapped

        # Replace in paragraphs and tables
        for para in doc.paragraphs:
            for ph, val in replacements.items():
                if ph in para.text:
                    if para.text.strip() == ph:
                        para.clear()
                        para.add_run(val)
                    else:
                        para.text = para.text.replace(ph, val)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for ph, val in replacements.items():
                            if ph in para.text:
                                if para.text.strip() == ph:
                                    para.clear()
                                    para.add_run(val)
                                else:
                                    para.text = para.text.replace(ph, val)

        doc.save(output_path)
        messagebox.showinfo("Success", f"Document saved to:\n{output_path}")

if __name__ == "__main__":
    app = SDSAuthoringTool()
    app.mainloop()