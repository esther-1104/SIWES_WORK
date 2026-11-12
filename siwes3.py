import csv
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
# openpyxl is used to export reports in Excel (.xlsx) format.
# Install with: pip install openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
 
# ── Constants ──
# Path to the CSV file used for persistent storage of student records.
DATA_FILE = "students.csv"
# Column headers for the CSV file — each record has a name, score, and grade.
FIELDNAMES = ["name", "score", "grade"]
 
 
def get_grade(score):
    """Return the letter grade for a given numeric score (0-100).
 
    Uses the University of Lagos grading scale with ordered
    if/elif branches. The highest threshold is checked first
    so that only one branch executes per call.
    """
    if score >= 70:   return "A"   # Excellent
    elif score >= 60: return "B"   # Very Good
    elif score >= 50: return "C"   # Good
    elif score >= 45: return "D"   # Fair
    elif score >= 40: return "E"   # Pass
    else:             return "F"   # Fail
 
 
class StudentRecordApp:
    """Main application class — builds the GUI, manages records,
    and handles all user interactions via callback methods."""
 
    def __init__(self, root):
        """Initialise the application: set window properties,
        load any existing records from CSV, build the UI,
        and populate the table."""
        self.root = root
        self.root.title("Student Record Management System")
        self.root.geometry("820x520")
        self.root.resizable(False, False)
        # Load saved records from disk (returns [] if file missing)
        self.records = self.load_records()
        self.build_ui()        # construct all widgets
        self.refresh_table()   # populate Treeview with loaded data
 
    # ── CSV Persistence ──
    # Records are stored in a simple CSV file so that data
    # survives between application sessions.
 
    def load_records(self):
        """Read student records from CSV into a list of dicts.
        Returns an empty list if the data file does not exist."""
        if not os.path.exists(DATA_FILE):
            return []
        with open(DATA_FILE, "r", newline="") as f:
            reader = csv.DictReader(f)
            # Convert score from string to int on load
            return [{"name": r["name"],
                     "score": int(r["score"]),
                     "grade": r["grade"]} for r in reader]
 
    def save_records(self):
        """Overwrite the CSV file with the current in-memory
        records list. Called after every add, update, or delete
        to ensure persistence."""
        with open(DATA_FILE, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
            writer.writeheader()
            writer.writerows(self.records)
 
    # ── GUI Construction ──
 
    def build_ui(self):
        """Create all frames, labels, entry fields, buttons,
        the Treeview table, and the statistics sidebar."""
 
        # ── Input Frame ──
        # Top panel where the user enters a student name and score.
        inp = ttk.LabelFrame(self.root, text="Add Student")
        inp.grid(row=0, column=0, padx=10, pady=5, sticky="ew")
 
        ttk.Label(inp, text="Name:").grid(row=0, column=0, padx=5, pady=5)
        self.name_var = tk.StringVar()  # bound to the Name entry field
        ttk.Entry(inp, textvariable=self.name_var, width=20).grid(
            row=0, column=1, padx=5)
 
        ttk.Label(inp, text="Score:").grid(row=0, column=2, padx=5)
        self.score_var = tk.StringVar()  # bound to the Score entry field
        ttk.Entry(inp, textvariable=self.score_var, width=8).grid(
            row=0, column=3, padx=5)
 
        # Buttons trigger callback methods when clicked
        ttk.Button(inp, text="Add", command=self.add_student).grid(
            row=0, column=4, padx=5)
        ttk.Button(inp, text="Clear", command=self.clear_fields).grid(
            row=0, column=5, padx=5)
 
        # ── Search Frame ──
        # Real-time filter: the table updates as the user types.
        srch = ttk.LabelFrame(self.root, text="Search")
        srch.grid(row=1, column=0, padx=10, pady=2, sticky="ew")
 
        self.search_var = tk.StringVar()
        # trace_add fires on_search() every time the text changes
        self.search_var.trace_add("write", lambda *_: self.on_search())
        ttk.Entry(srch, textvariable=self.search_var, width=30).grid(
            row=0, column=0, padx=5, pady=5)
        ttk.Label(srch, text="Type to filter by name").grid(
            row=0, column=1, padx=5)
 
        # ── Table Frame ──
        # Treeview widget displays student records in a scrollable table.
        tbl_frame = ttk.Frame(self.root)
        tbl_frame.grid(row=2, column=0, padx=10, pady=5, sticky="nsew")
 
        cols = ("no", "name", "score", "grade")
        self.tree = ttk.Treeview(tbl_frame, columns=cols,
                                 show="headings", height=14)
        # Configure column headings and widths
        for col, hdr, w in [("no","No.",40), ("name","Name",200),
                            ("score","Score",80), ("grade","Grade",80)]:
            self.tree.heading(col, text=hdr)
            self.tree.column(col, width=w, anchor="center"
                             if col != "name" else "w")
        # Attach a vertical scrollbar for long record lists
        sb = ttk.Scrollbar(tbl_frame, orient="vertical",
                           command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
 
        # ── Action Buttons ──
        btn = ttk.Frame(self.root)
        btn.grid(row=3, column=0, padx=10, pady=5)
        ttk.Button(btn, text="Update Selected",
                   command=self.update_selected).pack(side="left", padx=5)
        ttk.Button(btn, text="Delete Selected",
                   command=self.delete_selected).pack(side="left", padx=5)
        ttk.Button(btn, text="Export Report",
                   command=self.export_report).pack(side="left", padx=5)
 
        # ── Statistics Sidebar ──
        # A read-only Text widget on the right showing live stats.
        stats = ttk.LabelFrame(self.root, text="Class Statistics")
        stats.grid(row=0, column=1, rowspan=4, padx=10, pady=5,
                   sticky="ns")
        self.stats_text = tk.Text(stats, width=22, height=28,
                                  state="disabled", font=("Courier", 9))
        self.stats_text.pack(padx=5, pady=5)
 
    # ── CRUD Operations ──
 
    def clear_fields(self):
        """Reset both entry fields to empty strings."""
        self.name_var.set("")
        self.score_var.set("")
 
    def add_student(self):
        """Validate input, add a new student record, save to CSV,
        and refresh the display. Names are converted to title case
        so that input is case-insensitive — 'ade', 'ADE', and 'Ade'
        all become 'Ade'."""
        name = self.name_var.get().strip()
        if not name:
            messagebox.showerror("Error", "Name cannot be empty.")
            return
 
        # Normalise the name to title case so that "ade", "ADE",
        # and "Ade" are all stored as "Ade" — this makes every
        # subsequent search, update, and delete case-insensitive.
        name = name.title()
 
        # Check for duplicate names (case-insensitive comparison)
        # by comparing the lowered versions of each stored name.
        for r in self.records:
            if r["name"].lower() == name.lower():
                messagebox.showerror(
                    "Duplicate",
                    f"'{name}' already exists. Use Update instead.")
                return
 
        # Validate score: must be an integer between 0 and 100
        try:
            score = int(self.score_var.get())
        except ValueError:
            messagebox.showerror("Error", "Score must be a whole number.")
            return
        if score < 0 or score > 100:
            messagebox.showerror("Error", "Score must be 0-100.")
            return
 
        # Classify the score into a letter grade
        grade = get_grade(score)
 
        # Append the new record and persist to CSV
        self.records.append({"name": name, "score": score,
                             "grade": grade})
        self.save_records()
        self.clear_fields()
        self.refresh_table()
        messagebox.showinfo("Success", f"Added {name} ({grade})")
 
    def update_selected(self):
        """Update the score of the currently selected record.
        Opens a dialog for the new score, re-classifies the grade,
        saves to CSV, and refreshes the table."""
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Warning", "Select a record first.")
            return
        # The first column value is the row number (1-based)
        idx = int(self.tree.item(sel[0])["values"][0]) - 1
        rec = self.records[idx]
        # askinteger handles range validation (0-100) automatically
        new = simpledialog.askinteger(
            "Update Score", f"Current: {rec['name']} — {rec['score']}\n"
            f"Enter new score (0-100):",
            minvalue=0, maxvalue=100)
        if new is None:   # user pressed Cancel
            return
        rec["score"] = new
        rec["grade"] = get_grade(new)  # re-classify the grade
        self.save_records()
        self.refresh_table()
 
    def delete_selected(self):
        """Delete the currently selected record after a
        confirmation dialog. Saves and refreshes on confirm."""
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Warning", "Select a record first.")
            return
        idx = int(self.tree.item(sel[0])["values"][0]) - 1
        rec = self.records[idx]
        # Ask for confirmation before permanently removing
        if messagebox.askyesno("Confirm",
                f"Delete {rec['name']} ({rec['score']}, {rec['grade']})?"):
            self.records.pop(idx)
            self.save_records()
            self.refresh_table()
 
    # ── Search & Display ──
 
    def on_search(self):
        """Called on every keystroke in the search field.
        Delegates to refresh_table() which applies the filter."""
        self.refresh_table()
 
    def refresh_table(self):
        """Clear the Treeview and re-populate it from the
        in-memory records list. If a search query is active,
        only show records whose name contains the query
        (case-insensitive: both sides are lowered before
        comparison, so 'ade' matches 'Adewale')."""
        # Remove all existing rows from the Treeview
        for row in self.tree.get_children():
            self.tree.delete(row)
        # Convert search query to lowercase for case-insensitive match
        query = self.search_var.get().strip().lower()
        for i, r in enumerate(self.records, 1):
            # Skip records that do not match the search filter
            if query and query not in r["name"].lower():
                continue
            # Alternate row colours for readability
            tag = "even" if i % 2 == 0 else "odd"
            self.tree.insert("", "end", values=(i, r["name"],
                             r["score"], r["grade"]), tags=(tag,))
        self.tree.tag_configure("even", background="#f0f0f0")
        # Update the statistics sidebar to reflect visible data
        self.update_statistics()
 
    # ── Analytics ──
 
    def update_statistics(self):
        """Recompute and display class statistics in the sidebar:
        total students, average score, highest/lowest score,
        pass rate, and a grade distribution bar chart."""
        self.stats_text.config(state="normal")   # unlock for editing
        self.stats_text.delete("1.0", "end")     # clear old content
        if not self.records:
            self.stats_text.insert("end", "No records yet.")
            self.stats_text.config(state="disabled")
            return
        # Extract all scores into a list for aggregate calculations
        scores = [r["score"] for r in self.records]
        passed = sum(1 for r in self.records if r["grade"] != "F")
        avg = sum(scores) / len(scores)
        # Build a frequency dictionary for grade distribution
        dist = {}
        for r in self.records:
            dist[r["grade"]] = dist.get(r["grade"], 0) + 1
        # Format the statistics text
        txt  = f"Students : {len(self.records)}\n"
        txt += f"Average  : {avg:.1f}\n"
        txt += f"Highest  : {max(scores)}\n"
        txt += f"Lowest   : {min(scores)}\n"
        txt += f"Pass Rate: {(passed/len(self.records))*100:.1f}%\n"
        txt += f"{'—'*20}\n"
        txt += "Grade Distribution\n"
        for g in ["A","B","C","D","E","F"]:
            c = dist.get(g, 0)
            txt += f"  {g}: {c:>3}  {'█'*c}\n"
        self.stats_text.insert("end", txt)
        self.stats_text.config(state="disabled")  # lock again
 
    # ── Report Export ──
 
    def export_report(self):
        """Generate a formatted Excel (.xlsx) report of all
        records and save it to a file chosen via a save dialog.
        Uses openpyxl to create a styled workbook with headers,
        student data, and summary statistics."""
        if not self.records:
            messagebox.showinfo("Info", "No records to export.")
            return
        # Open a native file-save dialog filtered to .xlsx
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")])
        if not path:   # user pressed Cancel
            return
 
        # Create a new Excel workbook and select the active sheet
        wb = Workbook()
        ws = wb.active
        ws.title = "Student Report"
 
        # ── Define styles ──
        header_font = Font(bold=True, size=12)
        header_fill = PatternFill("solid", fgColor="D9E2F3")
        thin_border = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"), bottom=Side(style="thin"))
        center = Alignment(horizontal="center")
 
        # ── Title row ──
        ws.merge_cells("A1:D1")
        ws["A1"] = "STUDENT RECORD MANAGEMENT SYSTEM"
        ws["A1"].font = Font(bold=True, size=14)
        ws["A1"].alignment = center
 
        # ── Column headers (row 3) ──
        headers = ["No.", "Name", "Score", "Grade"]
        for col, hdr in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col, value=hdr)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = center
 
        # ── Student data rows ──
        for i, r in enumerate(self.records, 1):
            row = i + 3   # data starts at row 4
            ws.cell(row=row, column=1, value=i).border = thin_border
            ws.cell(row=row, column=2, value=r["name"]).border = thin_border
            ws.cell(row=row, column=3, value=r["score"]).border = thin_border
            ws.cell(row=row, column=4, value=r["grade"]).border = thin_border
            # Centre-align the No., Score, and Grade columns
            ws.cell(row=row, column=1).alignment = center
            ws.cell(row=row, column=3).alignment = center
            ws.cell(row=row, column=4).alignment = center
 
        # ── Summary statistics ──
        scores = [r["score"] for r in self.records]
        passed = sum(1 for r in self.records if r["grade"] != "F")
        avg = sum(scores) / len(scores)
        stat_row = len(self.records) + 5   # leave a blank row
        stats = [
            ("Total Students", len(self.records)),
            ("Class Average", round(avg, 1)),
            ("Highest Score", max(scores)),
            ("Lowest Score", min(scores)),
            ("Pass Rate (%)", round(passed / len(self.records) * 100, 1)),
        ]
        for j, (label, val) in enumerate(stats):
            r = stat_row + j
            ws.cell(row=r, column=2, value=label).font = Font(bold=True)
            ws.cell(row=r, column=3, value=val)
 
        # ── Adjust column widths for readability ──
        ws.column_dimensions["A"].width = 6
        ws.column_dimensions["B"].width = 28
        ws.column_dimensions["C"].width = 12
        ws.column_dimensions["D"].width = 10
 
        # Save the workbook to the chosen path
        wb.save(path)
        messagebox.showinfo("Exported", f"Report saved to {path}")
 
 
# ── Entry Point ──
# Create the Tk root window and start the event loop.
if __name__ == "__main__":
    root = tk.Tk()
    StudentRecordApp(root)
    root.mainloop()
