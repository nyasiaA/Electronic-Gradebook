import json
from pathlib import Path
from tkinter import *
from tkinter import simpledialog, messagebox, ttk
import pandas as pd

# --- DATA & PATH CONFIGURATION ---
# finds the Desktop and create the specific folder for Ms. Carter
desktopPath = Path.home() / "Desktop"
folderPath = desktopPath / "Ms. Carter's Classes"
dataFile = folderPath / "classData.json"

# Creates folder if it doesn't exist (exist_ok=True prevents errors if it does, basically checking for if the program has run before, reducing redundancy)
folderPath.mkdir(parents=True, exist_ok=True)

# saves the classroom info to the json file on the Desktop
def saveClassData(data):
    with open(dataFile, "w") as f:  # 'with' autocloses file when code is finished | 'w' is to write file | 'as f' temporary nickname
        json.dump(data, f, indent=4)  # 'json.dump' converts python to data to json | 'ident=4' adds 4 spaces

# checks for class data and loads it, if not skips
def loadClassData():
    if dataFile.exists():  # checking if the file exists
        with open(dataFile, "r") as f:  # 'r' read mode
            return json.load(f)  # returns the json string to python readable
    return None  # if nothing is there program just continues

# --- EXCEL LOGIC ---
def generateExcelSheet(classInfo):
    # Path for the Excel file inside the "Ms. Carter's Classes" folder
    excelFile = folderPath / "Class Roster.xlsx"

    with pd.ExcelWriter(excelFile, engine='openpyxl') as writer:
        for key, info in classInfo.items():
            excelNumbers = info['students']
            excelClass = info['name']

            data = {
                'Student Number': [f"{i:04}" for i in range(1, excelNumbers + 1)],
                'Student Name': ["" for _ in range(excelNumbers)]
            }

            if excelNumbers == 0:
                data = {'Student Number': ['0001'], 'Student Name': ['']}

            df = pd.DataFrame(data)
            safeNameCheck = "".join([c for c in excelClass if c.isalnum() or c == ' '])[:31]
            df.to_excel(writer, sheet_name=safeNameCheck, index=False)
            worksheet = writer.sheets[safeNameCheck]
            worksheet.column_dimensions['A'].width = 20
            worksheet.column_dimensions['B'].width = 40

    messagebox.showinfo("Excel Created", f"Excel rosters have been saved to:\n{excelFile}\nPlease complete the names for each student.")

# --- INITIAL SETUP ---
def runSetup():
    classInfo = {}
    setupWindow = Tk()
    setupWindow.withdraw()
    messagebox.showinfo("Setup", "Welcome Ms. Carter! Let's set up your classes.")

    for i in range(1, 6):
        defaultName = f"Class {i}"
        setupClassName = simpledialog.askstring("Class Name", f"Please choose the name for Class {i}:", initialvalue=defaultName)
        setupClassName = setupClassName if setupClassName else defaultName
        setupStudentCount = simpledialog.askinteger("Student Count", f"How many students are in {setupClassName}?", minvalue=0, maxvalue=150)
        classInfo[f"class{i}"] = {"name": setupClassName, "students": setupStudentCount if setupStudentCount is not None else 0}

    saveClassData(classInfo)
    generateExcelSheet(classInfo)
    setupWindow.destroy()
    return classInfo

# --- INITIALIZATION ---
currentData = loadClassData()
if not currentData:
    currentData = runSetup()

# --- MAIN GUI WINDOW ---
selectionWindow = Tk()
selectionWindow.geometry("1920x1080")
selectionWindow.title("Ms. Carter's Classes")
selectionWindow.config(background="#222222")

def on_closing():
    saveClassData(currentData)
    selectionWindow.destroy()

selectionWindow.protocol("WM_DELETE_WINDOW", on_closing)


label = Label(selectionWindow, text="Ms. Carter's Classes",
              font=('Arial', 40, 'bold'),
              fg='#c8102e',
              bg='white',
              relief="raised",
              bd=10,
              padx=20,
              pady=20,
              compound='center')
label.pack(pady=50)

# --- CLASS WINDOW FUNCTION ---
def openClassWindow(classNum):
    classKey = f"class{classNum}"
    current_class_data = currentData[classKey]
    className = current_class_data['name']

    classWin = Toplevel(selectionWindow)
    classWin.title(f"Class: {className}")
    classWin.geometry("1200x800")
    classWin.config(background="#222222")

    # Load roster from Excel
    def load_roster_from_excel():
        try:
            excel_path = folderPath / "Class Roster.xlsx"
            df = pd.read_excel(excel_path, sheet_name=className)
            return df.to_dict(orient='records')
        except Exception as e:
            return []

    roster = load_roster_from_excel()

    # Initialize data structures for new features
    if 'grades' not in current_class_data:
        current_class_data['grades'] = {}
    if 'assignments_meta' not in current_class_data:
        current_class_data['assignments_meta'] = {}

    # --- TAB INTERFACE ---
    notebook = ttk.Notebook(classWin)
    roster_tab = Frame(notebook, bg="#222222")
    assignments_tab = Frame(notebook, bg="#222222")
    reports_tab = Frame(notebook, bg="#222222")

    notebook.add(roster_tab, text='Students')
    notebook.add(assignments_tab, text='Assignments')
    notebook.add(reports_tab, text='Grades & Reports')
    notebook.pack(expand=1, fill='both')

    # --- STUDENTS TAB ---
    canvas = Canvas(roster_tab, bg="#222222", highlightthickness=0)
    scrollbar = Scrollbar(roster_tab, orient="vertical", command=canvas.yview)
    scroll_frame = Frame(canvas, bg="#222222")
    scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side=LEFT, fill=BOTH, expand=True)
    scrollbar.pack(side=RIGHT, fill=Y)

    Label(scroll_frame, text="ID - Name", width=50, anchor='w', bg="#222222", fg="white").pack(side=TOP, padx=10, pady=5)

    def calculate_grade(student_id):
        # get grades for student
        grades = current_class_data['grades'].get(student_id, {})
        meta = current_class_data['assignments_meta']

        # if no grades or no assignments, default letter grade
        if not grades or not meta:
            return "F"  # default to F if no scores

        total_earned = sum(score for score in grades.values() if isinstance(score, (int, float)))
        total_max = sum(meta[name]['max'] for name in grades.keys())

        if total_max == 0:
            return "F"  # default to F if no max points exist

        percent = (total_earned / total_max) * 100

        # convert percentage to letter grade
        if percent >= 90:
            return "A"
        elif percent >= 80:
            return "B"
        elif percent >= 70:
            return "C"
        elif percent >= 60:
            return "D"
        else:
            return "F"

    for student in roster:
        student_id = str(student['Student Number'])
        student_name = student.get('Student Name', 'Unnamed')

        row_frame = Frame(scroll_frame, bg="#222222")
        row_frame.pack(fill=X, pady=2)

        # Calculates the numerical percentage for a student (used for at-risk highlighting)
        def calculate_percentage(student_id):
            grades = current_class_data['grades'].get(student_id, {})
            meta = current_class_data['assignments_meta']
            if not grades or not meta:
                return 0  # default 0% if no grades

            total_earned = sum(score for score in grades.values() if isinstance(score, (int, float)))
            total_max = sum(meta[name]['max'] for name in grades.keys())
            if total_max == 0:
                return 0

            return (total_earned / total_max) * 100

        # Button for each student to see detailed grades
        def show_student_details(sid=student_id, sname=student_name):
            student_win = Toplevel(classWin)
            student_win.title(f"{sname} - {sid}")
            student_win.geometry("500x600")
            student_win.config(bg="#222222")

            Label(student_win, text=f"{sname} (ID: {sid})", font=('Arial', 16, 'bold'), bg="#222222", fg="white").pack(
                pady=10)

            # Display overall grade with percentage
            overall_grade = calculate_grade(sid)
            percent = calculate_percentage(sid)  # ← add this line
            Label(student_win, text=f"Overall Grade: {overall_grade} ({percent:.1f}%)", font=('Arial', 14),
                  bg="#222222",
                  fg="white").pack(pady=5)

            # Frame for assignment list
            assignments_frame = Frame(student_win, bg="#222222")
            assignments_frame.pack(fill=BOTH, expand=True, pady=10, padx=10)

            # Header
            Label(assignments_frame, text="Assignment", font=('Arial', 12, 'bold'), width=20, anchor='w', bg="#222222",
                  fg="white").grid(row=0, column=0, padx=5, pady=2)
            Label(assignments_frame, text="Score", font=('Arial', 12, 'bold'), width=10, anchor='w', bg="#222222",
                  fg="white").grid(row=0, column=1, padx=5, pady=2)

            # Show each assignment
            for i, assignment in enumerate(current_class_data['assignments_meta'], start=1):
                max_points = current_class_data['assignments_meta'][assignment]['max']
                score = current_class_data['grades'].get(sid, {}).get(assignment, 0)

                # Check if missing
                if 'missing_grades' in current_class_data and sid in current_class_data[
                    'missing_grades'] and assignment in current_class_data['missing_grades'][sid]:
                    score_text = f"{score} (Missing)"
                    fg_color = "red"
                else:
                    score_text = f"{score} / {max_points}"
                    fg_color = "white"

                Label(assignments_frame, text=assignment, font=('Arial', 12), width=20, anchor='w', bg="#222222",
                      fg="white").grid(row=i, column=0, padx=5, pady=2)
                Label(assignments_frame, text=score_text, font=('Arial', 12), width=10, anchor='w', bg="#222222",
                      fg=fg_color).grid(row=i, column=1, padx=5, pady=2)

        Button(row_frame, text=f"{student_id} - {student_name} ", width=50, anchor='w',
               command=show_student_details).pack(side=LEFT, padx=10)

    # --- ASSIGNMENTS TAB ---
    def add_assignment():
        name = simpledialog.askstring("Assignment Name", "Enter Assignment Name:")
        if not name: return
        max_points = simpledialog.askinteger("Max Points", f"Enter max points for {name}:", minvalue=1)
        if not max_points: return
        current_class_data['assignments_meta'][name] = {'max': max_points}

        # Assign grades to each student (allow blank for missing)
        for student in roster:
            sid = str(student['Student Number'])
            student_name = student.get('Student Name', 'Unnamed')

            while True:  # Loop until valid input or blank
                score_str = simpledialog.askstring(
                    "Grade",
                    f"Enter score for {student_name} (ID: {sid}) [Leave blank for missing]:"
                )

                if score_str is None or score_str.strip() == "":  # Blank or cancel
                    score = 0
                    # Store missing grade
                    if 'missing_grades' not in current_class_data:
                        current_class_data['missing_grades'] = {}
                    if sid not in current_class_data['missing_grades']:
                        current_class_data['missing_grades'][sid] = []
                    current_class_data['missing_grades'][sid].append(name)
                    break  # Exit loop, accept 0
                else:
                    try:
                        score = float(score_str)
                        if 0 <= score <= max_points:  # Valid range
                            break  # Accept this score
                        else:  # Out-of-range counts as missing
                            messagebox.showwarning("Invalid",
                                                   f"Score must be between 0 and {max_points}. Treated as missing.")
                            score = 0
                            if 'missing_grades' not in current_class_data:
                                current_class_data['missing_grades'] = {}
                            if sid not in current_class_data['missing_grades']:
                                current_class_data['missing_grades'][sid] = []
                            current_class_data['missing_grades'][sid].append(name)
                            break
                    except ValueError:  # Non-numeric input
                        messagebox.showwarning("Invalid", "Non-numeric input. Treated as missing.")
                        score = 0
                        if 'missing_grades' not in current_class_data:
                            current_class_data['missing_grades'] = {}
                        if sid not in current_class_data['missing_grades']:
                            current_class_data['missing_grades'][sid] = []
                        current_class_data['missing_grades'][sid].append(name)
                        break

            # Store the grade
            if sid not in current_class_data['grades']:
                current_class_data['grades'][sid] = {}
            current_class_data['grades'][sid][name] = score

            saveClassData(currentData)

    def edit_assignment():
        if not current_class_data['assignments_meta']:
            messagebox.showwarning("Warning", "No assignments to edit!")
            return

        edit_win = Toplevel(classWin)
        edit_win.title("Edit Assignment")
        edit_win.geometry("500x350")

        Label(edit_win, text="Select Assignment:").pack(pady=5)

        assignment_var = StringVar()
        assignment_dropdown = ttk.Combobox(
            edit_win,
            textvariable=assignment_var,
            values=list(current_class_data['assignments_meta'].keys()),
            state="readonly"
        )
        assignment_dropdown.pack(pady=5)
        assignment_dropdown.current(0)

        Label(edit_win, text="New Name (optional):").pack(pady=5)
        name_entry = Entry(
            edit_win,
            bg="gray",
            fg="black",
            font=('Arial', 14),
            width=30,
            insertbackground="black"
        )
        name_entry.pack(pady=10, ipady=6)  # note: I increased padding a bit

        Label(edit_win, text="New Max Points (optional):").pack(pady=5)
        points_entry = Entry(
            edit_win,
            bg="gray",
            fg="black",
            font=('Arial', 14),
            width=15,
            insertbackground="black"
        )
        points_entry.pack(pady=10, ipady=6)

        def save_changes():
            old_name = assignment_var.get()  # the currently selected assignment
            new_name = name_entry.get().strip() or old_name
            try:
                new_max = int(points_entry.get())
            except ValueError:
                new_max = current_class_data['assignments_meta'][old_name]['max']

            # store the old max points BEFORE updating
            old_max = current_class_data['assignments_meta'][old_name]['max']

            # --- Update assignment meta & rename ---
            current_class_data['assignments_meta'][new_name] = {'max': new_max}
            if new_name != old_name:
                del current_class_data['assignments_meta'][old_name]

            # --- Move grades to new name if renamed ---
            for sid in current_class_data['grades']:
                if old_name in current_class_data['grades'][sid]:
                    current_class_data['grades'][sid][new_name] = current_class_data['grades'][sid][old_name]
                    if new_name != old_name:
                        del current_class_data['grades'][sid][old_name]

            # --- Only prompt teacher for new grades if max points changed ---
            if new_max != old_max:
                for student in roster:
                    sid = str(student['Student Number'])
                    student_name = student.get('Student Name', 'Unnamed')

                    while True:
                        new_score_str = simpledialog.askstring(
                            "Grade Update",
                            f"Enter score for {student_name} (ID: {sid}) [Leave blank for missing]:"
                        )

                        if new_score_str is None or new_score_str.strip() == "":  # Blank
                            new_score = 0
                            if 'missing_grades' not in current_class_data:
                                current_class_data['missing_grades'] = {}
                            if sid not in current_class_data['missing_grades']:
                                current_class_data['missing_grades'][sid] = []
                            current_class_data['missing_grades'][sid].append(new_name)
                            break
                        else:
                            try:
                                new_score = float(new_score_str)
                                if 0 <= new_score <= new_max:
                                    break
                                else:
                                    messagebox.showwarning("Invalid",
                                                           f"Score must be between 0 and {new_max}. Treated as missing.")
                                    new_score = 0
                                    if 'missing_grades' not in current_class_data:
                                        current_class_data['missing_grades'] = {}
                                    if sid not in current_class_data['missing_grades']:
                                        current_class_data['missing_grades'][sid] = []
                                    current_class_data['missing_grades'][sid].append(new_name)
                                    break
                            except ValueError:
                                messagebox.showwarning("Invalid", "Non-numeric input. Treated as missing.")
                                new_score = 0
                                if 'missing_grades' not in current_class_data:
                                    current_class_data['missing_grades'] = {}
                                if sid not in current_class_data['missing_grades']:
                                    current_class_data['missing_grades'][sid] = []
                                current_class_data['missing_grades'][sid].append(new_name)
                                break

                    # Store the score
                    if sid not in current_class_data['grades']:
                        current_class_data['grades'][sid] = {}
                    current_class_data['grades'][sid][new_name] = new_score

            saveClassData(currentData)
            messagebox.showinfo("Updated", f"Assignment '{old_name}' updated!")
            edit_win.destroy()

        Button(edit_win, text="Save Changes", command=save_changes).pack(pady=15)

    def remove_assignment():
        if not current_class_data['assignments_meta']:
            messagebox.showwarning("Warning", "No assignments to remove!")
            return

        remove_win = Toplevel(classWin)
        remove_win.title("Remove Assignment")
        remove_win.geometry("500x350")

        Label(remove_win, text="Select Assignment to Remove:").pack(pady=10)

        assignment_var = StringVar()
        assignment_dropdown = ttk.Combobox(
            remove_win,
            textvariable=assignment_var,
            values=list(current_class_data['assignments_meta'].keys()),
            state="readonly"
        )
        assignment_dropdown.pack(pady=5)
        assignment_dropdown.current(0)

        def confirm_remove():
            name = assignment_var.get()

            del current_class_data['assignments_meta'][name]
            for sid in current_class_data['grades']:
                current_class_data['grades'][sid].pop(name, None)

            saveClassData(currentData)
            messagebox.showinfo("Removed", f"Assignment '{name}' removed!")
            remove_win.destroy()

        Button(remove_win, text="Remove", fg="black", bg="red", command=confirm_remove).pack(pady=20)

    Button(assignments_tab, text="Add Assignment", command=add_assignment).pack(pady=10)
    Button(assignments_tab, text="Edit Assignment", command=edit_assignment).pack(pady=10)
    Button(assignments_tab, text="Remove Assignment", command=remove_assignment).pack(pady=10)

    def highlight_at_risk():
        report_win = Toplevel(classWin)
        report_win.title("At-Risk Students")
        report_win.geometry("400x600")

        Label(report_win, text="At Risk Students", font=('Arial', 14, 'bold')).pack(pady=10)

        # --- SCROLLABLE CANVAS SETUP ---
        canvas = Canvas(report_win, bg="#222222", highlightthickness=0)
        scrollbar = Scrollbar(report_win, orient="vertical", command=canvas.yview)
        scroll_frame = Frame(canvas, bg="#222222")

        scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

        # --- POPULATE AT-RISK STUDENTS ---
        for student in roster:
            sid = str(student['Student Number'])
            student_name = student.get('Student Name', 'Unnamed')

            grades = current_class_data['grades'].get(sid, {})
            meta = current_class_data['assignments_meta']
            if not grades or not meta:
                percent = 0
            else:
                total_earned = sum(score for score in grades.values() if isinstance(score, (int, float)))
                total_max = sum(meta[name]['max'] for name in grades.keys())
                percent = (total_earned / total_max) * 100 if total_max > 0 else 0

            # Only show at-risk students
            if 70 <= percent <= 76:
                color = "yellow"
            elif percent < 70:
                color = "red"
            else:
                continue  # Skip students who are not at risk

            Label(scroll_frame, text=f"{student_name} ({sid}) - {percent:.1f}%", bg=color, fg="black", anchor='w').pack(
                fill=X, padx=10, pady=2
            )

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def show_progress_report():
        report_win = Toplevel(classWin)
        report_win.title("Progress Report")
        report_win.geometry("700x600")
        report_win.config(bg="#222222")

        # Header
        Label(report_win, text="Progress Report", font=('Arial', 16, 'bold'), bg="#222222", fg="white").pack(pady=10)

        canvas = Canvas(report_win, bg="#222222", highlightthickness=0)
        scrollbar = Scrollbar(report_win, orient="vertical", command=canvas.yview)
        frame = Frame(canvas, bg="#222222")
        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

        # Table header
        Label(frame, text="Student", font=('Arial', 12, 'bold'), width=20, anchor='w', bg="#222222", fg="white").grid(
            row=0, column=0, padx=5, pady=2)
        Label(frame, text="Grade (%)", font=('Arial', 12, 'bold'), width=10, anchor='w', bg="#222222", fg="white").grid(
            row=0, column=1, padx=5, pady=2)
        Label(frame, text="Assignments", font=('Arial', 12, 'bold'), width=40, anchor='w', bg="#222222",
              fg="white").grid(row=0, column=2, padx=5, pady=2)

        # Fill in student data
        for i, student in enumerate(roster, start=1):
            sid = str(student['Student Number'])
            sname = student.get('Student Name', 'Unnamed')
            percent = calculate_percentage(sid)
            grade = calculate_grade(sid)

            # Assignments summary
            assignments_summary = []
            for assignment in current_class_data['assignments_meta']:
                score = current_class_data['grades'].get(sid, {}).get(assignment, 0)
                if 'missing_grades' in current_class_data and sid in current_class_data[
                    'missing_grades'] and assignment in current_class_data['missing_grades'][sid]:
                    assignments_summary.append(f"{assignment}: Missing")
                else:
                    max_points = current_class_data['assignments_meta'][assignment]['max']
                    assignments_summary.append(f"{assignment}: {score}/{max_points}")
            assignments_text = ", ".join(assignments_summary)

            # Fill in row
            Label(frame, text=sname, font=('Arial', 12), width=20, anchor='w', bg="#222222", fg="white").grid(row=i,
                                                                                                              column=0,
                                                                                                              padx=5,
                                                                                                              pady=2)
            Label(frame, text=f"{grade} ({percent:.1f}%)", font=('Arial', 12), width=10, anchor='w', bg="#222222",
                  fg="white").grid(row=i, column=1, padx=5, pady=2)
            Label(frame, text=assignments_text, font=('Arial', 12), width=40, anchor='w', bg="#222222",
                  fg="white").grid(row=i, column=2, padx=5, pady=2)

    def show_missing_assignments():
        if not current_class_data['assignments_meta']:
            messagebox.showwarning("Warning", "No assignments available!")
            return

        missing_win = Toplevel(classWin)
        missing_win.title("Missing Assignments")
        missing_win.geometry("500x600")
        missing_win.config(bg="#222222")

        Label(missing_win, text="Select Assignment:", font=('Arial', 14, 'bold'), bg="#222222", fg="white").pack(
            pady=10)

        assignment_var = StringVar()
        assignment_dropdown = ttk.Combobox(
            missing_win,
            textvariable=assignment_var,
            values=list(current_class_data['assignments_meta'].keys()),
            state="readonly",
            font=('Arial', 12)
        )
        assignment_dropdown.pack(pady=5)
        assignment_dropdown.current(0)

        # Frame to display missing students
        missing_frame = Frame(missing_win, bg="#222222")
        missing_frame.pack(fill=BOTH, expand=True, pady=10, padx=10)

        def update_missing_list(event=None):
            # Clear previous list
            for widget in missing_frame.winfo_children():
                widget.destroy()

            selected_assignment = assignment_var.get()
            missing_students = []

            for student in roster:
                sid = str(student['Student Number'])
                sname = student.get('Student Name', 'Unnamed')
                if 'missing_grades' in current_class_data and sid in current_class_data['missing_grades']:
                    if selected_assignment in current_class_data['missing_grades'][sid]:
                        missing_students.append(f"{sname} ({sid})")

            if missing_students:
                for student in missing_students:
                    Label(missing_frame, text=student, font=('Arial', 12), anchor='w', bg="#222222", fg="red").pack(
                        fill=X, pady=2)
            else:
                Label(missing_frame, text="No students missing this assignment!", font=('Arial', 12), anchor='w',
                      bg="#222222", fg="white").pack(fill=X, pady=2)

        assignment_dropdown.bind("<<ComboboxSelected>>", update_missing_list)
        update_missing_list()

    # Buttons for grades tab
    Button(reports_tab, text="Highlight At-Risk Students", command=highlight_at_risk).pack(pady=10)
    Button(reports_tab, text="Show Progress Report", command=show_progress_report).pack(pady=10)
    Button(reports_tab, text="Missing Assignments", command=show_missing_assignments).pack(pady=10)


# --- MAIN BUTTONS ---
for i in range(1, 6):
    classKey = f"class{i}"
    className = currentData[classKey]['name']
    studentCount = currentData[classKey]['students']

    classButton = Button(selectionWindow,
                         text=f"{className} ({studentCount} Students)",
                         font=('Arial', 20, 'bold'),
                         fg='#c8102e',
                         bg='white',
                         activebackground='#a00',
                         padx=30,
                         pady=10,
                         relief="raised",
                         borderwidth=10,
                         command=lambda c=i: openClassWindow(c))
    classButton.pack(pady=15)

selectionWindow.mainloop()