# ─── Flask and Related Utilities ────────────────────────────────────────────────
from flask import (
    Flask, request, render_template, redirect, url_for,
    flash, session, send_file, jsonify, send_from_directory
)

from flask import render_template, request, jsonify, send_file
import sqlite3
from datetime import datetime
from io import BytesIO

from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# ─── Security ───────────────────────────────────────────────────────────────────
from werkzeug.security import generate_password_hash, check_password_hash

# ─── Database and Data Handling ─────────────────────────────────────────────────
import sqlite3
from collections import defaultdict
from datetime import datetime
from flask_cors import CORS

# ─── File & I/O Operations ──────────────────────────────────────────────────────
from io import BytesIO
import io
import csv
from flask import abort
import re
import os
from werkzeug.utils import secure_filename

# ─── Excel, PDF & Data Processing ───────────────────────────────────────────────
from openpyxl import Workbook
import pandas as pd

# ─── ReportLab for PDF Export ───────────────────────────────────────────────────
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
import sqlite3


app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Required for session and flash

# Folder to store uploaded lecture files
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Allowed file types
ALLOWED_EXTENSIONS = {'pdf', 'ppt', 'pptx', 'doc', 'docx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


#home route
@app.route('/')
def home_page():
    return render_template('home/home.html')

# index route
@app.route('/index_page')
def index_page():
    return render_template('home/index.html')

#database naming 
DATABASE = 'database.db'

def get_db_connection():
    conn = sqlite3.connect('database.db')
    conn.row_factory = sqlite3.Row
    return conn

DB_NAME = 'database.db'
def init_db():
    with sqlite3.connect('database.db') as conn:
        cursor = conn.cursor()
        

        cursor.execute('')
        # Admin table
        cursor.execute('''CREATE TABLE IF NOT EXISTS admin (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            email TEXT UNIQUE NOT NULL,
            contact TEXT CHECK(length(contact) = 10),
            password TEXT NOT NULL
        )''')
        
        cursor.execute(' ')
        # Student table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS student (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            register_no TEXT UNIQUE,
            admission_no TEXT UNIQUE,
            full_name TEXT NOT NULL,
            email TEXT UNIQUE,
            contact TEXT UNIQUE,
            gender TEXT,
            dob TEXT,
            age INTEGER,
            department TEXT,
            batch_year INTEGER
        )''')

        
        cursor.execute('')
        # Teacher table
        cursor.execute('''CREATE TABLE IF NOT EXISTS teacher (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            teacher_id TEXT UNIQUE,
            full_name TEXT NOT NULL,
            email TEXT UNIQUE,
            contact TEXT UNIQUE,
            gender TEXT NOT NULL,
            dob TEXT NOT NULL,
            age INTEGER NOT NULL,
            department TEXT NOT NULL,
            role TEXT NOT NULL,
            experience INTEGER NOT NULL,
            qualification TEXT NOT NULL
        )''')

        cursor.execute('')
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS notification (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            topic TEXT NOT NULL,
            message TEXT NOT NULL,
            posted_date TEXT NOT NULL,
            posted_time TEXT NOT NULL,
            priority TEXT DEFAULT 'Medium',
            visibility TEXT DEFAULT 'All',
            attachments TEXT,
            expiry_date TEXT,
            status TEXT DEFAULT 'Active',
            created_by TEXT,
            last_updated TEXT
        )''')

        cursor.execute('')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS feedback (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            message TEXT NOT NULL,
            rating INTEGER NOT NULL
        )''')

        cursor.execute('')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS period(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            department TEXT,
            batch_year INTEGER,
            subject TEXT,
            day TEXT,
            period_number INTEGER
        )''')

        cursor.execute('')
        cursor.execute('''DELETE FROM period
        WHERE id IN (
        SELECT id FROM (
        SELECT id, ROW_NUMBER() OVER (PARTITION BY department, batch_year, day, period_number, subject ORDER BY id) AS rn
        FROM period
        ) t WHERE rn > 1    
        );''')
       
        cursor.execute('')
        # Attendance table
        cursor.execute('''CREATE TABLE IF NOT EXISTS attendance (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        register_no TEXT NOT NULL,
        full_name TEXT NOT NULL,
        department TEXT NOT NULL,
        batch_year TEXT NOT NULL,
        period_number INTEGER NOT NULL,
        subject TEXT NOT NULL,
        attendance_status TEXT NOT NULL CHECK(attendance_status IN ('Present','Late','Absent')),
        date TEXT NOT NULL,
        time TEXT NOT NULL,
        UNIQUE(register_no, date, period_number)  -- prevents duplicate marking
        );''')

        cursor.execute('''''')
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS lecture_notes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            department TEXT NOT NULL,
            semester TEXT NOT NULL,
            filename TEXT NOT NULL,
            filepath TEXT NOT NULL,
            upload_time TEXT NOT NULL,
            hidden INTEGER DEFAULT 0
        )""")


        cursor.execute('')
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS tasks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            department TEXT NOT NULL,
            semester TEXT NOT NULL,
            category TEXT NOT NULL,
            title TEXT NOT NULL,
            description TEXT,
            warning TEXT,
            color TEXT,
            deadline TEXT,
            hidden INTEGER DEFAULT 0
        )
        """)
        
        cursor.execute("""""")
        cursor.execute("""CREATE TABLE IF NOT EXISTS work_evaluation (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL,
            task_id INTEGER NOT NULL,
            status TEXT CHECK(status IN ('Submitted', 'Not Submitted')) DEFAULT 'Not Submitted',
            file_path TEXT,
            submitted_on TEXT,
            can_reupload INTEGER DEFAULT 0,
            done_status TEXT CHECK(done_status IN ('Done', 'Not Done')) DEFAULT 'Not Done',
            UNIQUE(student_id, task_id),
            FOREIGN KEY (student_id) REFERENCES student(id),
            FOREIGN KEY (task_id) REFERENCES tasks(id)
        );""")
        

        cursor.execute('')
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS submitted_works (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL,
            task_id INTEGER NOT NULL,
            file_path TEXT NOT NULL,
            submitted_on DATETIME DEFAULT CURRENT_TIMESTAMP,
            can_reupload INTEGER DEFAULT 0,
            FOREIGN KEY (student_id) REFERENCES student(id),
            FOREIGN KEY (task_id) REFERENCES tasks(id)
        );""")



        cursor.execute('')
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS student_mark (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        student_id INTEGER,
        task_id INTEGER,
        mark INTEGER DEFAULT 0,
        grade TEXT,
        entered_on DATETIME DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(student_id, task_id)
        );""")

        
        conn.commit()

#admin register
@app.route('/admin_register', methods=['GET', 'POST'])
def admin_register():
    if request.method == 'POST':
        name = request.form['name']
        email = request.form['email']
        contact = request.form['contact']
        password = request.form['password']
        confirm_password = request.form['confirm_password'] 

        if password != confirm_password:
            flash('Passwords do not match.')
            return redirect(url_for('admin_register'))

        hashed_password = generate_password_hash(password)

        try:
            with sqlite3.connect('database.db') as conn:
                cursor = conn.cursor()
                cursor.execute("INSERT INTO admin (name, email, contact, password) VALUES (?, ?, ?, ?)",
                               (name, email, contact, hashed_password))
                conn.commit()
                flash('Admin registered successfully.')
                return redirect(url_for('admin_login'))
        except sqlite3.IntegrityError:
            flash('Email already exists.')
            return redirect(url_for('admin_register'))

    return render_template('admin/register.html')

#admin login                          
@app.route('/admin_login', methods=['GET', 'POST'])
def admin_login():
    if request.method == 'POST':
        email = request.form['email']
        password = request.form['password']

        with sqlite3.connect('database.db') as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT id, name, email, contact, password FROM admin WHERE email = ?", (email,))
            admin = cursor.fetchone()

            if admin:
                stored_password = admin[4]  # password from DB
                if check_password_hash(stored_password, password):
                    session['admin_id'] = admin[0]  # Save admin ID to session
                    session['admin_name'] = admin[1]  # Save admin name if needed
                    flash('Login successful!')
                    return redirect(url_for('admin_dashboard'))
                else:
                    flash('Incorrect password.')
            else:
                flash('No admin account found with that email.')

        return redirect(url_for('admin_login'))

    return render_template('admin/login.html')

#admin dashboard
@app.route('/admin_dashboard')
def admin_dashboard():
    if 'admin_id' in session:
        admin_id = session['admin_id']
        with sqlite3.connect('database.db') as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT name, email FROM admin WHERE id = ?", (admin_id,))
            admin = cursor.fetchone()
            if admin:
                return render_template('admin/dashboard.html', admin_name=admin[0], admin_email=admin[1])
    flash('Please log in first.')
    return redirect(url_for('admin_login'))

# student login
@app.route('/student_login', methods=['GET', 'POST'])
def student_login():
    if request.method == 'POST':
        register_no = request.form['register_no'].upper()
        email = request.form['email'].lower()
        dob_input = request.form['dob']  # From input type="date", format is YYYY-MM-DD
        try:
            # Convert to the format used in DB: DD-MM-YYYY
            dob = datetime.strptime(dob_input, '%Y-%m-%d').strftime('%d-%m-%Y')
        except ValueError:
            flash('Invalid date format.')
            return redirect(url_for('student_login'))

        with sqlite3.connect('database.db') as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT id, full_name FROM student 
                WHERE register_no = ? AND email = ? AND dob = ?
            ''', (register_no, email, dob))
            student = cursor.fetchone()

            if student:
                session['student_id'] = student[0]
                session['student_name'] = student[1]
                flash('Login successful.')
                return redirect(url_for('student_dashboard'))  # Define this route separately
            else:
                flash('Invalid credentials. Please try again.')

    return render_template('student/login.html')


#student dashboard
@app.route('/student_dashboard')
def student_dashboard():
    if 'student_id' in session:
        with sqlite3.connect('database.db') as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT full_name, register_no, email FROM student WHERE id = ?', (session['student_id'],))
            student = cursor.fetchone()
            if student:
                return render_template('student/dashboard.html', student=student)
    flash('Please log in first.')
    return redirect(url_for('student_login'))

# Teacher Login 
@app.route('/teacher_login', methods=['GET', 'POST'])
def teacher_login():
    if request.method == 'POST':
        teacher_id = request.form['teacher_id'].upper()
        email = request.form['email'].lower()
        dob_input = request.form['dob']  # Format from input is YYYY-MM-DD

        try:
            dob = datetime.strptime(dob_input, '%Y-%m-%d').strftime('%d-%m-%Y')
        except ValueError:
            flash('Invalid DOB format.')
            return redirect(url_for('teacher_login'))

        with sqlite3.connect('database.db') as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT id, full_name FROM teacher
                WHERE teacher_id = ? AND email = ? AND dob = ?
            ''', (teacher_id, email, dob))
            teacher = cursor.fetchone()

            if teacher:
                session['teacher_id'] = teacher[0]
                session['teacher_name'] = teacher[1]
                flash('Login successful.')
                return redirect(url_for('teacher_dashboard'))  # Define this route
            else:
                flash('Invalid credentials. Please try again.')

    return render_template('teacher/login.html')



#teacher dashboard
@app.route('/teacher_dashboard')
def teacher_dashboard():
    if 'teacher_id' in session:
        with sqlite3.connect('database.db') as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT full_name, teacher_id, email FROM teacher WHERE id = ?', (session['teacher_id'],))
            teacher = cursor.fetchone()
            if teacher:
                return render_template('teacher/dashboard.html', teacher=teacher)
    flash('Please log in first.')
    return redirect(url_for('teacher_login'))

from flask import request, render_template, flash
from datetime import datetime
import sqlite3


@app.route('/students', methods=['GET', 'POST'])
def students():

    # =========================
    # POST : ADD STUDENT
    # =========================
    if request.method == 'POST':
        register_no = request.form['register_no'].upper()
        admission_no = request.form['admission_no']
        full_name = request.form['full_name'].upper()
        email = request.form['email'].lower()
        contact = request.form['contact']
        gender = request.form['gender'].upper()
        dob_input = request.form['dob']
        department = request.form['department'].upper()
        batch_year = int(request.form['batch_year'])

        # DOB formatting
        birth_date = datetime.strptime(dob_input, '%Y-%m-%d')
        dob = birth_date.strftime('%d-%m-%Y')

        # Age calculation
        today = datetime.today()
        age = today.year - birth_date.year - (
            (today.month, today.day) < (birth_date.month, birth_date.day)
        )

        try:
            with sqlite3.connect('database.db') as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO student
                    (register_no, admission_no, full_name, email, contact,
                     gender, dob, age, department, batch_year)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    register_no, admission_no, full_name, email,
                    contact, gender, dob, age, department, batch_year
                ))
                conn.commit()
                flash('Student added successfully.')

        except sqlite3.IntegrityError:
            flash('Register No, Admission No, or Email already exists.')

    # =========================
    # GET : FILTER STUDENTS
    # =========================
    search_reg = request.args.get('search_reg', '')
    search_name = request.args.get('search_name', '')
    search_department = request.args.get('search_department', '')
    batch_year_category = request.args.get('batch_year', '')
    year_filter = request.args.get('year_filter', '')

    # =========================
    # ACADEMIC YEAR LOGIC
    # =========================
    today = datetime.today()

    # Academic year starts in June
    # Jan–May belong to previous academic year
    if today.month <= 5:
        academic_year = today.year - 1
    else:
        academic_year = today.year

    # =========================
    # BUILD QUERY
    # =========================
    query = "SELECT * FROM student WHERE 1=1"
    params = []

    # Search filters
    if search_reg:
        query += " AND register_no LIKE ?"
        params.append(f"%{search_reg.upper()}%")

    if search_name:
        query += " AND full_name LIKE ?"
        params.append(f"%{search_name.upper()}%")

    if search_department:
        query += " AND department LIKE ?"
        params.append(f"%{search_department.upper()}%")

    # Year category dropdown
    if batch_year_category == "First Year":
        query += " AND (? - batch_year) = 0"
        params.append(academic_year)

    elif batch_year_category == "Second Year":
        query += " AND (? - batch_year) = 1"
        params.append(academic_year)

    elif batch_year_category == "Third Year":
        query += " AND (? - batch_year) = 2"
        params.append(academic_year)

    elif batch_year_category == "Fourth Year / Graduated":
        query += " AND (? - batch_year) = 3"
        params.append(academic_year)

    elif batch_year_category == "Graduated":
        query += " AND (? - batch_year) > 3"
        params.append(academic_year)

    # Studying / Studied radio filters
    if year_filter == 'studying':
        query += " AND (? - batch_year) <= 3"
        params.append(academic_year)

    elif year_filter == 'studied':
        query += " AND (? - batch_year) > 3"
        params.append(academic_year)

    query += " ORDER BY register_no ASC"

    # =========================
    # FETCH DATA
    # =========================
    with sqlite3.connect('database.db') as conn:
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute(query, params)
        students = cursor.fetchall()

    # =========================
    # RENDER TEMPLATE
    # =========================
    return render_template(
    'admin/student_handle.html',
    students=students,
    current_year=academic_year,   # ✅ FIX
    academic_year=academic_year,
    search_reg=search_reg,
    search_name=search_name,
    search_department=search_department
)


# Delete a student
@app.route('/delete_student/<int:id>')
def delete_student(id):
    with sqlite3.connect('database.db') as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM student WHERE id=?", (id,))
        conn.commit()
    flash('Student deleted successfully.')
    return redirect(url_for('students'))

#edit student handle
@app.route('/edit_student/<int:id>', methods=['GET', 'POST'])
def edit_student(id):
    conn = sqlite3.connect('database.db')
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    if request.method == 'POST':
        data = request.form
        try:
            # Parse and convert date of birth
            birth_date = datetime.strptime(data['dob'], '%Y-%m-%d')
            dob_formatted = birth_date.strftime('%d-%m-%Y')

            # Calculate age
            today = datetime.today()
            age = today.year - birth_date.year - ((today.month, today.day) < (birth_date.month, birth_date.day))

            # Update the student in the database
            cursor.execute('''
                UPDATE student SET 
                    register_no = ?, 
                    admission_no = ?, 
                    full_name = ?, 
                    email = ?, 
                    contact = ?, 
                    gender = ?, 
                    dob = ?, 
                    age = ?, 
                    department = ?, 
                    batch_year = ?
                WHERE id = ?
            ''', (
                data['register_no'].upper(),
                data['admission_no'],
                data['full_name'].upper(),
                data['email'].lower(),
                data['contact'],
                data['gender'].upper(),
                dob_formatted,
                age,
                data['department'].upper(),
                int(data['batch_year']),
                id
            ))

            conn.commit()
            flash("Student details updated successfully.")
            return redirect(url_for('students'))

        except ValueError as e:
            flash(f"Invalid input: {e}")

    cursor.execute("SELECT * FROM student WHERE id = ?", (id,))
    student = cursor.fetchone()

    # Add ISO date for form
    if student:
        dob_obj = datetime.strptime(student['dob'], '%d-%m-%Y')
        student = dict(student)
        student['dob_iso'] = dob_obj.strftime('%Y-%m-%d')

    return render_template("admin/edit_student.html", student=student)




#teacher handle
@app.route('/teacher_handle', methods=['GET', 'POST'])
def teacher_handle():
    conn = get_db_connection()
    
    if request.method == 'POST':
        form = request.form
        dob_raw = form['dob']
        dob_obj = datetime.strptime(dob_raw, "%Y-%m-%d")
        dob_formatted = dob_obj.strftime("%d-%m-%Y")
        today = datetime.today()
        age = today.year - dob_obj.year - ((today.month, today.day) < (dob_obj.month, dob_obj.day))

        try:
            conn.execute('''INSERT INTO teacher 
                (teacher_id, full_name, email, contact, gender, dob, age, department, role, experience, qualification)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                (form['teacher_id'], form['full_name'], form['email'], form['contact'],
                form['gender'], dob_formatted, age, form['department'],
                form['role'], int(form['experience']), form['qualification']))

            conn.commit()
            flash("Teacher added successfully!")
        except sqlite3.IntegrityError:
            flash("Error: Duplicate teacher ID, email, or contact.")
        return redirect(url_for('teacher_handle'))

    # Search & Filter
    search_name = request.args.get('search_name', '')
    search_department = request.args.get('search_department', '')
    sort_by = request.args.get('sort_by', '')
    page = int(request.args.get('page', 1))
    per_page = 10
    offset = (page - 1) * per_page

    query = "SELECT * FROM teacher WHERE 1=1"
    params = []

    if search_name:
        query += " AND full_name LIKE ?"
        params.append(f"%{search_name}%")
    if search_department:
        query += " AND department LIKE ?"
        params.append(f"%{search_department}%")

    query += f" ORDER BY {sort_by if sort_by in ['experience', 'department'] else 'id ASC'}"
    query += " LIMIT ? OFFSET ?"
    params.extend([per_page + 1, offset])

    teachers = conn.execute(query, params).fetchall()
    has_next = len(teachers) > per_page
    teachers = teachers[:per_page]

    return render_template('admin/teacher_handle.html',
                           teachers=teachers,
                           page=page,
                           has_next=has_next,
                           search_name=search_name,
                           search_department=search_department)

#edit teacher
@app.route("/edit_teacher/<int:id>", methods=["GET", "POST"])
def edit_teacher(id):
    conn = get_db_connection()

    if request.method == "POST":
        teacher_id = request.form["teacher_id"].upper()
        full_name = request.form["full_name"].upper()
        email = request.form["email"].lower()
        contact = request.form["contact"]
        gender = request.form["gender"]
        dob = request.form["dob"]  # YYYY-MM-DD from input
        department = request.form["department"]
        role = request.form["role"]
        experience = request.form["experience"]
        qualification = request.form["qualification"].upper()

        # Store DOB in DD-MM-YYYY format
        dob_display = datetime.strptime(dob, "%Y-%m-%d").strftime("%d-%m-%Y")

        conn.execute("""
            UPDATE teacher
            SET teacher_id=?, full_name=?, email=?, contact=?, gender=?, dob=?, department=?, role=?, experience=?, qualification=?
            WHERE rowid=?
        """, (teacher_id, full_name, email, contact, gender, dob_display,
              department, role, experience, qualification, id))
        conn.commit()
        conn.close()

        flash("Teacher updated successfully!")
        return redirect(url_for("teacher_handle"))

    # GET request: fetch teacher data
    teacher = conn.execute("SELECT rowid, * FROM teacher WHERE rowid=?", (id,)).fetchone()
    conn.close()

    if teacher is None:
        flash("Teacher not found!")
        return redirect(url_for("teacher_handle"))

    teacher = dict(teacher)
    # Convert DOB to YYYY-MM-DD for HTML date input
    try:
        teacher["dob"] = datetime.strptime(teacher["dob"], "%d-%m-%Y").strftime("%Y-%m-%d")
    except ValueError:
        pass  # already in correct format

    return render_template("admin/edit_teacher.html", teacher=teacher)

#delete teacher
@app.route('/delete_teacher/<int:id>')
def delete_teacher(id):
    conn = get_db_connection()
    conn.execute("DELETE FROM teacher WHERE id = ?", (id,))
    conn.commit()
    flash("Teacher deleted.")
    return redirect(url_for('teacher_handle'))


#export list
@app.route('/export_teachers')
def export_teachers():
    search_name = request.args.get('search_name', '')
    search_department = request.args.get('search_department', '')
    sort_by = request.args.get('sort_by', '')

    query = "SELECT * FROM teacher WHERE 1=1"
    params = []

    if search_name:
        query += " AND full_name LIKE ?"
        params.append(f"%{search_name}%")
    if search_department:
        query += " AND department LIKE ?"
        params.append(f"%{search_department}%")
    
    query += f" ORDER BY {sort_by if sort_by in ['experience', 'department'] else 'id DESC'}"

    conn = get_db_connection()
    teachers = conn.execute(query, params).fetchall()

    output = io.StringIO()
    writer = csv.writer(output)

    writer.writerow(['ID', 'Teacher ID', 'Name', 'Email', 'Contact', 'Gender',
                     'DOB', 'Age', 'Department', 'Role', 'Experience', 'Qualification'])

    for t in teachers:
        writer.writerow([t['id'], t['teacher_id'], t['full_name'], t['email'], t['contact'],
                         t['gender'], t['dob'], t['age'], t['department'], t['role'],
                         t['experience'], t['qualification']])

    output.seek(0)
    return send_file(io.BytesIO(output.read().encode('utf-8')),
                     mimetype='text/csv',
                     as_attachment=True,
                     download_name='filtered_teachers.csv')

#notification update
@app.route('/notification_update', methods=['GET', 'POST'])
def notification_update():
    conn = sqlite3.connect('database.db')
    cursor = conn.cursor()

    # Handle add notification
    if request.method == 'POST':
        topic = request.form['topic']
        message = request.form['message']
        priority = request.form.get('priority', 'Medium')
        visibility = request.form.get('visibility', 'All')
        attachments = request.form.get('attachments', '').strip()
        if attachments and not re.match(r'^https?://', attachments):
            flash('Invalid URL format in attachment.')
            return redirect(url_for('notification_update'))
        expiry_date_input = request.form['expiry_date']
        expiry_date = datetime.strptime(expiry_date_input, '%Y-%m-%d').strftime('%d-%m-%Y')

        posted_date = datetime.today().strftime('%d-%m-%Y')
        posted_time = datetime.now().strftime('%H:%M:%S')
        created_by = request.form.get('created_by', 'Admin')
        last_updated = posted_date

        cursor.execute('''
            INSERT INTO notification (
                topic, message, posted_date, posted_time, priority, visibility, 
                attachments, expiry_date, status, created_by, last_updated
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            topic, message, posted_date, posted_time, priority, visibility,
            '', expiry_date, 'Active', created_by, last_updated
        ))
        conn.commit()
        flash('Notification added successfully!')
        return redirect(url_for('notification_update'))

    # Filter logic
    search_topic = request.args.get('filter_topic', '').strip()
    search_date = request.args.get('filter_date', '').strip()
    search_status = request.args.get('filter_status', '').strip()


    query = "SELECT * FROM notification"
    conditions = []
    params = []

    if search_topic:
        conditions.append("topic LIKE ?")
        params.append(f"%{search_topic}%")
    if search_date:
        conditions.append("posted_date LIKE ?")
        params.append(f"{search_date}%")  # Date is usually YYYY-MM-DD format
    if search_status:
        conditions.append("status = ?")
        params.append(search_status)

    if conditions:
        query += " WHERE " + " OR ".join(conditions)
        cursor.execute(query, params)
    else:
        cursor.execute(query)

    notifications = cursor.fetchall()

    # Auto-update status for all notifications
    today = datetime.today()
    for notif in notifications:
        expiry_str = notif[8]
        status = notif[9]
        if expiry_str:
            try:
                expiry_date = datetime.strptime(expiry_str, '%d-%m-%Y')
                if today > expiry_date and status == 'Active':
                    cursor.execute("UPDATE notification SET status='Expired' WHERE id=?", (notif[0],))
            except ValueError:
                pass  # handle invalid date format if needed

    conn.commit()

    return render_template('admin/notification_update.html', notifications=notifications)

#delete notification
@app.route('/delete_notification/<int:id>')
def delete_notification(id):
    with sqlite3.connect('database.db') as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM notification WHERE id=?", (id,))
        conn.commit()
    flash('Notification deleted successfully!')
    return redirect(url_for('notification_update'))

#feedback update
@app.route('/feedback', methods=['GET', 'POST'])
def feedback():
    if request.method == 'POST':
        message = request.form['message']
        rating = int(request.form['rating'])

        with sqlite3.connect('database.db') as conn:
            conn.execute('INSERT INTO feedback (message, rating) VALUES (?, ?)', (message, rating))
        return redirect('/feedback')

    # GET request
    with sqlite3.connect('database.db') as conn:
        conn.row_factory = sqlite3.Row
        feedbacks = conn.execute('SELECT * FROM feedback ORDER BY id DESC').fetchall()
    return render_template('home/feedback.html', feedbacks=feedbacks)

# Route to handle and display feedback
@app.route('/feedback_handle')
def feedback_handle():
    conn = get_db_connection()
    feedbacks = conn.execute("SELECT * FROM feedback ORDER BY id DESC").fetchall()
    conn.close()
    return render_template('admin/feedback_handle.html', feedbacks=feedbacks)

# Route to delete feedback
@app.route('/feedback/delete/<int:id>', methods=['POST'])
def delete_feedback(id):
    conn = get_db_connection()
    conn.execute('DELETE FROM feedback WHERE id = ?', (id,))
    conn.commit()
    conn.close()
    flash('Feedback removed successfully.', 'success')
    return redirect(url_for('feedback_handle'))

# Route to examine (this can be extended later)
@app.route('/feedback/examine/<int:id>')
def examine_feedback(id):
    conn = get_db_connection()
    feedback = conn.execute('SELECT * FROM feedback WHERE id = ?', (id,)).fetchone()
    conn.close()
    if feedback:
        return f"<h3>Examine Feedback</h3><p>ID: {feedback['id']}</p><p>Rating: {feedback['rating']} ★</p><p>Message: {feedback['message']}</p><a href='/feedback_handle'>← Back</a>"
    else:
        return "<p>Feedback not found.</p><a href='/feedback_handle'>← Back</a>"
    

# =========================
# PERIOD MANAGEMENT
# =========================
@app.route("/periods", methods=["GET", "POST"])
def periods():
    conn = get_db_connection()

    if request.method == "POST":
        # Fetch form data
        department = request.form.get("department")
        batch_year = request.form.get("batch_year")
        subject = request.form.get("subject")
        day = request.form.get("day")
        period_number = request.form.get("period_number")

        # Validate required fields
        if not (department and batch_year and subject and day and period_number):
            return "All fields required!", 400

        # Insert into period table
        conn.execute(
            "INSERT INTO period (department, batch_year, subject, day, period_number) VALUES (?,?,?,?,?)",
            (department, int(batch_year), subject, day, int(period_number))
        )
        conn.commit()

    # Build timetable for display
    timetable = {day: {i: "" for i in range(1, 6)} for day in ["Monday","Tuesday","Wednesday","Thursday","Friday"]}
    rows = conn.execute("SELECT * FROM period").fetchall()

    batch_year_map = {1: "First Year", 2: "Second Year", 3: "Third Year", 4: "Fourth Year"}

    for row in rows:
        timetable[row["day"]][row["period_number"]] = f"{row['subject']}<br>({row['department']} - {batch_year_map.get(row['batch_year'],'N/A')})"

    conn.close()
    return render_template("teacher/period.html", timetable=timetable)

@app.route("/delete_period", methods=["POST"])
def delete_period():
    day = request.form.get("day")
    period_number = request.form.get("period_number")

    if not day or not period_number:
        return "Invalid request", 400

    conn = get_db_connection()
    conn.execute("DELETE FROM period WHERE day=? AND period_number=?", (day, int(period_number)))
    conn.commit()
    conn.close()

    return redirect("/periods")



# --- Attendance Page ---
@app.route("/attendance")
def attendance():
    return render_template("teacher/attendance.html")


# --- Fetch Students for Attendance ---
@app.route("/get_students")
def get_students():
    department = request.args.get("department")
    year_label = request.args.get("year")

    if not department or not year_label:
        return jsonify({"students": []})

    # -----------------------------
    # Academic year calculation
    # -----------------------------
    today = datetime.now()
    if today.month >= 6:      # June–December
        academic_year = today.year
    else:                     # January–May
        academic_year = today.year - 1

    # -----------------------------
    # Map year label → batch year
    # -----------------------------
    year_map = {
        "First Year": academic_year,
        "Second Year": academic_year - 1,
        "Third Year": academic_year - 2,
        "Fourth Year": academic_year - 3,
        "Forth Year / Graduated": academic_year - 3,
        "Graduated": None
    }

    target_batch = year_map.get(year_label)

    conn = get_db_connection()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    # -----------------------------
    # Fetch students
    # -----------------------------
    if year_label == "Graduated":
        students = cursor.execute(
            """
            SELECT register_no, full_name, department, batch_year
            FROM student
            WHERE department = ?
              AND (? - batch_year) > 3
            """,
            (department, academic_year)
        ).fetchall()
    else:
        students = cursor.execute(
            """
            SELECT register_no, full_name, department, batch_year
            FROM student
            WHERE department = ?
              AND batch_year = ?
            """,
            (department, target_batch)
        ).fetchall()

    conn.close()

    return jsonify({
        "students": [dict(stu) for stu in students]
    })

#attendance
@app.route("/attendance")
def loadps():
    return render_template("teacher/attendance.html")

@app.route("/load_data")
def load_data():
    department = request.args.get("department")
    year_label = request.args.get("year")

    if not department or not year_label:
        return jsonify({"students": [], "periods": []})

    # =========================
    # ACADEMIC YEAR LOGIC
    # =========================
    today = datetime.now()

    # Academic year starts in June
    if today.month >= 6:          # June–December
        academic_year = today.year
    else:                         # January–May
        academic_year = today.year - 1

    # =========================
    # Map year label → batch_year (students)
    # =========================
    year_map_students = {
        "First Year": academic_year,
        "Second Year": academic_year - 1,
        "Third Year": academic_year - 2,
        "Fourth Year": academic_year - 3,
        "Graduated": None
    }
    target_batch_student = year_map_students.get(year_label)

    # =========================
    # Map year label → study year (periods table)
    # =========================
    year_map_periods = {
        "First Year": 1,
        "Second Year": 2,
        "Third Year": 3,
        "Fourth Year": 4,
        "Graduated": None
    }
    target_batch_period = year_map_periods.get(year_label)

    conn = get_db_connection()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    # =========================
    # Fetch Students
    # =========================
    if year_label == "Graduated":
        students = cursor.execute(
            """
            SELECT register_no, full_name, department, batch_year
            FROM student
            WHERE department = ?
              AND (? - batch_year) > 3
            """,
            (department, academic_year)
        ).fetchall()
    else:
        students = cursor.execute(
            """
            SELECT register_no, full_name, department, batch_year
            FROM student
            WHERE department = ?
              AND batch_year = ?
            """,
            (department, target_batch_student)
        ).fetchall()

    # =========================
    # Fetch Periods (by study year)
    # =========================
    if target_batch_period:
        periods = cursor.execute(
            """
            SELECT DISTINCT day, period_number, subject
            FROM period
            WHERE department = ?
              AND batch_year = ?
              AND day = ?
            ORDER BY period_number
            """,
            (department, target_batch_period, today.strftime("%A"))
        ).fetchall()
    else:
        periods = []

    conn.close()

    return jsonify({
        "students": [dict(s) for s in students],
        "periods": [dict(p) for p in periods]
    })

@app.route("/submit_attendance", methods=["POST"])
def submit_attendance():
    data = request.get_json()
    department = data.get("department")
    year_label = data.get("year")
    period_number = data.get("period")
    records = data.get("records")  # { reg_no: {period, status}, ... }

    if not (department and year_label and period_number and records):
        return jsonify({"success": False, "error": "Invalid data"})

    conn = get_db_connection()
    cursor = conn.cursor()

    today_date = datetime.now().strftime("%Y-%m-%d")
    now_time = datetime.now().strftime("%H:%M:%S")

    # Map year_label to batch_year for period
    batch_year_map = {"First Year": 1, "Second Year": 2, "Third Year": 3, "Fourth Year": 4}
    batch_year = batch_year_map.get(year_label)

    # Get subject name for this period
    subject_row = cursor.execute(
        "SELECT subject FROM period WHERE department=? AND batch_year=? AND period_number=? AND day=?",
        (department, batch_year, period_number, datetime.now().strftime("%A"))
    ).fetchone()
    subject = subject_row["subject"] if subject_row else "Unknown"

    inserted, skipped = 0, 0
    for reg_no, rec in records.items():
        status = rec["status"]

        # Check student exists
        stu = cursor.execute(
            "SELECT register_no, full_name, department, batch_year FROM student WHERE register_no=?",
            (reg_no,)
        ).fetchone()
        if not stu:
            skipped += 1
            continue

        try:
            # Insert attendance, ignore if already exists for this date & period
            cursor.execute("""
                INSERT OR IGNORE INTO attendance
                (register_no, full_name, department, batch_year, period_number, subject, attendance_status, date, time)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                stu["register_no"],
                stu["full_name"],
                stu["department"],
                stu["batch_year"],
                period_number,
                subject,
                status,
                today_date,
                now_time
            ))
            if cursor.rowcount == 1:
                inserted += 1
            else:
                skipped += 1
        except Exception as e:
            skipped += 1

    conn.commit()
    conn.close()

    return jsonify({"success": True, "inserted": inserted, "skipped": skipped})

# Attendance details page (renders HTML)
@app.route("/attendance_details")
def attendance_details_page():
    return render_template("teacher/attendance_detail.html")

@app.route("/get_attendance")
def get_attendance():
    department = request.args.get("department")
    period = request.args.get("period")
    from_date = request.args.get("from_date")
    to_date = request.args.get("to_date")
    register_no = request.args.get("register_no")

    if not from_date or not to_date:
        return jsonify([])

    conn = get_db_connection()
    cursor = conn.cursor()

    # 🧠 Join attendance with period to get subject name dynamically
    base_query = """
        SELECT 
            a.register_no,
            a.full_name,
            a.department,
            a.batch_year,
            a.period_number,
            a.date,
            a.attendance_status
        FROM attendance a
        LEFT JOIN period p 
            ON a.department = p.department
           AND a.batch_year = p.batch_year
           AND a.period_number = p.period_number
    """

    # --- Case 1: Register number provided ---
    if register_no:
        query = base_query + """
            WHERE a.register_no = ?
              AND DATE(a.date) BETWEEN DATE(?) AND DATE(?)
            ORDER BY a.date ASC, a.register_no ASC
        """
        params = [register_no, from_date, to_date]

    # --- Case 2: Department + Period provided ---
    else:
        if not department or not period:
            return jsonify([])

        query = base_query + """
            WHERE a.department = ?
              AND a.period_number = ?
              AND DATE(a.date) BETWEEN DATE(?) AND DATE(?)
            ORDER BY a.date ASC, a.register_no ASC
        """
        params = [department, period, from_date, to_date]

    rows = cursor.execute(query, params).fetchall()
    conn.close()

    return jsonify([dict(row) for row in rows])


# --- PDF EXPORT ---
@app.route("/export_attendance_pdf", methods=["POST"])
def export_attendance_pdf():
    data = request.json.get("attendance", [])

    if not data:
        return jsonify({"error": "No data to export"}), 400

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    styles = getSampleStyleSheet()

    elements = []
    elements.append(Paragraph("Attendance Report", styles["Title"]))

    # Table header
    table_data = [["Register No", "Full Name", "Department", "Batch", "Period", "Date", "Status"]]

    # Table rows
    for row in data:
        combined = f"{row['period_number']}"
        table_data.append([
            row["register_no"],
            row["full_name"],
            row["department"],
            row["batch_year"],
            combined,
            row["date"],
            row["attendance_status"]
        ])

    table = Table(table_data, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))

    elements.append(table)
    doc.build(elements)
    buffer.seek(0)

    return send_file(buffer, as_attachment=True, download_name="attendance_report.pdf", mimetype="application/pdf")

#lecture note route
@app.route("/lecture_notes")
def lecture_notes():
    return render_template("teacher/lecture_notes.html")

@app.route('/upload_lecture', methods=['POST'])
def upload_lecture():
    department = request.form.get('department')
    semester = request.form.get('semester')
    file = request.files.get('file')

    if not file or file.filename == '':
        return jsonify({'message': 'No file uploaded.'}), 400

    filename = secure_filename(file.filename)
    upload_dir = app.config['UPLOAD_FOLDER']
    os.makedirs(upload_dir, exist_ok=True)
    upload_path = os.path.join(upload_dir, filename)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        cursor.execute("""
            SELECT COUNT(*) FROM lecture_notes 
            WHERE department = ? AND semester = ? AND filename = ?
        """, (department, semester, filename))
        exists = cursor.fetchone()[0]

        if exists:
            conn.close()
            return jsonify({'message': f'Lecture "{filename}" already exists for {department}, {semester}.'}), 400

        # Save file
        file.save(upload_path)
        upload_time = datetime.now().strftime("%dd-%mm-%yyyy %H:%M:%S")

        cursor.execute("""
            INSERT INTO lecture_notes (department, semester, filename, filepath, upload_time)
            VALUES (?, ?, ?, ?, ?)
        """, (department, semester, filename, upload_path, upload_time))
        conn.commit()

        print("✅ Lecture inserted:", department, semester, filename)

    except Exception as e:
        conn.rollback()
        print("❌ ERROR inserting lecture:", e)
        return jsonify({'message': 'Database insert failed.'}), 500

    finally:
        conn.close()

    return jsonify({'message': f'Lecture "{filename}" uploaded successfully!'})

@app.route("/get_lectures", methods=["GET"])
def get_lectures():
    conn = get_db_connection()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM lecture_notes ORDER BY upload_time DESC")
    lectures = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return jsonify(lectures)


@app.route('/toggle_visibility/<int:lecture_id>', methods=['POST'])
def toggle_visibility(lecture_id):
    data = request.get_json()
    hidden = data.get('hidden', False)
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("UPDATE lecture_notes SET hidden = ? WHERE id = ?", (hidden, lecture_id))
    conn.commit()
    conn.close()
    return jsonify({"message": "Lecture visibility updated."})


@app.route('/delete_lecture/<int:lecture_id>', methods=['DELETE'])
def delete_lecture(lecture_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM lecture_notes WHERE id = ?", (lecture_id,))
    conn.commit()
    conn.close()
    return jsonify({"message": "Lecture deleted successfully."})

#academic tasks
@app.route("/academic_task")
def academic_task():
    return render_template("teacher/academic_task.html")

def dict_factory(cursor, row):
    """Convert database rows to dictionaries."""
    d = {}
    for idx, col in enumerate(cursor.description):
        d[col[0]] = row[idx]
    return d
    
@app.route("/api/tasks", methods=["GET"])
def get_tasks():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = dict_factory
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM tasks")
    data = cursor.fetchall()
    conn.close()
    return jsonify(data)

@app.route("/api/tasks", methods=["POST"])
def add_task():
    data = request.json
    required_fields = ["department", "semester", "category", "title"]

    if not all(field in data and data[field] for field in required_fields):
        return jsonify({"error": "Missing required fields"}), 400

    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO tasks (department, semester, category, title, description, warning, color, deadline)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        data["department"],
        data["semester"],
        data["category"],
        data["title"],
        data.get("description", ""),
        data.get("warning", ""),
        data.get("color", "green"),
        data.get("deadline", "")
    ))
    conn.commit()
    conn.close()
    return jsonify({"message": "Task added successfully!"}), 201

@app.route("/api/tasks/<int:task_id>", methods=["DELETE"])
def delete_task(task_id):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM tasks WHERE id = ?", (task_id,))
    conn.commit()
    conn.close()
    return jsonify({"message": "Task deleted successfully!"})

@app.route("/api/tasks/<int:task_id>/toggle_hide", methods=["PATCH"])
def toggle_hide(task_id):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT hidden FROM tasks WHERE id = ?", (task_id,))
    result = cursor.fetchone()

    if not result:
        conn.close()
        return jsonify({"error": "Task not found"}), 404

    new_hidden = 0 if result[0] else 1
    cursor.execute("UPDATE tasks SET hidden = ? WHERE id = ?", (new_hidden, task_id))
    conn.commit()
    conn.close()
    return jsonify({"message": f"Task {'hidden' if new_hidden else 'unhidden'} successfully."})


# -----------------------------
# Academic Work Page (Student)
# -----------------------------
@app.route("/academic_work")
def academic_work():
    if "student_id" not in session:
        return redirect("/login")

    student_id = session["student_id"]

    # Fetch student info
    conn = sqlite3.connect("database.db")
    cur = conn.cursor()
    cur.execute("""
        SELECT full_name, department, register_no, batch_year
        FROM student
        WHERE id = ?
    """, (student_id,))
    row = cur.fetchone()
    conn.close()

    if not row:
        return "Student not found", 404

    student = {
        "full_name": row[0],
        "department": row[1],
        "register_no": row[2],
        "batch_year": int(row[3])
    }

    # Calculate current semester
    now = datetime.now()
    current_year = now.year
    current_month = now.month
    batch_year = int(student["batch_year"])

# -----------------------------
# Academic base year
# -----------------------------
# Academic year starts in June
    if current_month >= 6:          # June–Dec
        academic_base_year = current_year
    else:                           # Jan–May
        academic_base_year = current_year - 1

# -----------------------------
# Years completed since joining
# -----------------------------
    years_completed = academic_base_year - batch_year

# -----------------------------
# Semester calculation
# -----------------------------
    if 6 <= current_month <= 10:          # June – October
        semester_number = years_completed * 2 + 1
    elif current_month in (11, 12, 1, 2, 3):  # November – March
        semester_number = years_completed * 2 + 2
    else:                                   # April – May (gap)
        semester_number = years_completed * 2

# Clamp semester between 1 and 8
    semester_number = max(1, min(semester_number, 8))

    student["semester"] = f"Semester {semester_number}"

# -----------------------------
# Academic year range
# -----------------------------
    academic_year_start = academic_base_year
    academic_year_end = academic_base_year + 1

    student["academic_year"] = f"{academic_year_start}-{academic_year_end}"

    # Fetch tasks
    conn = sqlite3.connect("database.db")
    cur = conn.cursor()
    cur.execute("""
        SELECT id, category, title, description, deadline, warning
        FROM tasks
        WHERE department = ? AND semester = ?
    """, (student["department"], student["semester"]))
    task_rows = cur.fetchall()

    tasks = [
        {
            "id": r[0],
            "category": r[1],
            "title": r[2],
            "description": r[3],
            "deadline": r[4],
            "warning": r[5]
        }
        for r in task_rows
    ]

    # Fetch student's submissions
    cur.execute("""
        SELECT task_id, file_path, can_reupload
        FROM submitted_works
        WHERE student_id = ?
    """, (student_id,))
    done_records = {
        str(r[0]): {"file_path": r[1], "can_reupload": r[2]} for r in cur.fetchall()
    }
    conn.close()

    return render_template(
        "student/academic_work.html",
        student=student,
        tasks=tasks,
        done_records=done_records
    )


@app.route("/uploads/<path:filename>")
def uploaded_file(filename):
    return send_from_directory(app.config["UPLOAD_FOLDER"], filename)


@app.route("/upload_academic_work", methods=["POST"])
def upload_academic_work():
    if "student_id" not in session:
        return jsonify({"error": "Unauthorized"}), 401

    student_id = session["student_id"]
    task_id = request.form.get("task_id")
    file = request.files.get("file")

    if not task_id:
        return jsonify({"error": "Missing task ID"}), 400
    if not file:
        return jsonify({"error": "No file uploaded"}), 400
    if not allowed_file(file.filename):
        return jsonify({"error": "Invalid file type"}), 400

    filename = f"student_{student_id}_task_{task_id}_{secure_filename(file.filename)}"
    os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
    file_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)

    conn = sqlite3.connect("database.db")
    cur = conn.cursor()

    cur.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_unique_submission 
        ON submitted_works (student_id, task_id);
    """)

    # Check reupload permission
    cur.execute("""
        SELECT can_reupload FROM submitted_works 
        WHERE student_id=? AND task_id=?
    """, (student_id, task_id))
    existing = cur.fetchone()

    if existing and existing[0] == 0:
        conn.close()
        return jsonify({"error": "Re-upload not allowed for this submission"}), 403

    try:
        file.save(file_path)
    except Exception as e:
        conn.close()
        return jsonify({"error": f"File save failed: {str(e)}"}), 500

    cur.execute("""
        INSERT INTO submitted_works (student_id, task_id, file_path, submitted_on, can_reupload)
        VALUES (?, ?, ?, CURRENT_TIMESTAMP, 0)
        ON CONFLICT(student_id, task_id) DO UPDATE SET
            file_path = excluded.file_path,
            submitted_on = excluded.submitted_on,
            can_reupload = excluded.can_reupload
    """, (student_id, task_id, filename))

    conn.commit()
    conn.close()

    return jsonify({"success": True, "file_path": filename})


# -----------------------------
# Helper Function
# -----------------------------
def calculate_semester(batch_year: int) -> str:
    """Return semester name (Semester 1–8) based on batch year and current date."""
    try:
        batch_year = int(batch_year)
    except (TypeError, ValueError):
        return "Semester 1"

    now = datetime.now()
    current_year = now.year
    current_month = now.month
    years_passed = current_year - batch_year

    if 6 <= current_month <= 10:
        semester_number = years_passed * 2 + 1   # June–Oct → Odd semester
    elif current_month >= 11 or current_month <= 3:
        semester_number = years_passed * 2 + 2   # Nov–Mar → Even semester
    else:
        semester_number = years_passed * 2       # April–May break

    semester_number = max(1, min(semester_number, 8))
    return f"Semester {semester_number}"

# -----------------------------
# Work Valuate (Teacher)
# -----------------------------
@app.route("/work_valuate")
def work_valuate():
    return render_template("teacher/work_valuate.html")

@app.route("/get_submissions")
def get_submissions():
    department = request.args.get("department")
    semester = request.args.get("semester")
    work_type = request.args.get("work_type")

    if not department or not semester or not work_type:
        return jsonify({"error": "Missing parameters"}), 400

    conn = sqlite3.connect("database.db")
    cur = conn.cursor()

    # ✅ Fetch task details
    cur.execute("""
        SELECT id, department, semester, category, title, description, deadline
        FROM tasks
        WHERE department=? AND semester=? AND category=?
    """, (department, semester, work_type))
    
    task = cur.fetchone()
    if not task:
        conn.close()
        return jsonify({"error": "No task found"}), 404

    task_info = {
        "task_id": task[0],
        "department": task[1],
        "semester": task[2],
        "category": task[3],
        "title": task[4],
        "description": task[5],
        "deadline": task[6]
    }

    # ✅ Fetch all students in department
    cur.execute("""
        SELECT id, register_no, full_name
        FROM student
        WHERE department=?
    """, (department,))
    students = cur.fetchall()

    # ✅ Fetch submission records for this task
    cur.execute("""
        SELECT student_id, file_path, submitted_on, can_reupload
        FROM submitted_works
        WHERE task_id=?
    """, (task_info["task_id"],))
    submissions = cur.fetchall()

    conn.close()

    submission_dict = {
        s[0]: {
            "file_path": s[1],
            "submitted_on": s[2],
            "can_reupload": s[3]
        }
        for s in submissions
    }

    result = []
    for st in students:
        sid, reg_no, name = st
        if sid in submission_dict:
            sub = submission_dict[sid]
            result.append({
                "student_id": sid,
                "register_no": reg_no,
                "full_name": name, 
                "semester":semester,
                "status": "Submitted",
                "file_path": sub["file_path"],
                "submitted_on": sub["submitted_on"],
                "can_reupload": sub["can_reupload"]
            })
        else:
            result.append({
                "student_id": sid,
                "register_no": reg_no,
                "full_name": name,
                "semester":semester,
                "status": "Not Submitted",
                "file_path": None,
                "submitted_on": None,
                "can_reupload": 0
            })

    return jsonify({
        "task_info": task_info,
        "students": result
    })


# ✅ Allow re-upload & clear previous submission
@app.route("/allow_reupload/<int:student_id>/<int:task_id>", methods=["POST"])
def allow_reupload(student_id, task_id):
    try:
        conn = sqlite3.connect("database.db")
        cur = conn.cursor()

        # Clear file_path and allow re-upload
        cur.execute("""
            UPDATE submitted_works
            SET file_path = NULL,
                submitted_on = NULL,
                can_reupload = 1
            WHERE student_id=? AND task_id=?
        """, (student_id, task_id))

        # If no record yet, insert new one
        if cur.rowcount == 0:
            cur.execute("""
                INSERT INTO submitted_works (student_id, task_id, file_path, submitted_on, can_reupload)
                VALUES (?, ?, NULL, NULL, 1)
            """, (student_id, task_id))

        conn.commit()
        conn.close()

        return jsonify({"message": "✅ Re-upload allowed and previous file cleared."}), 200

    except Exception as e:
        print("Error:", e)
        return jsonify({"error": "Internal server error"}), 500

   
# -----------------------------
# Save DONE / NOT DONE to work_evaluation table
# -----------------------------
@app.route("/update_done_status", methods=["POST"])
def update_done_status():
    data = request.json
    student_id = data.get("student_id")
    task_id = data.get("task_id")
    done_status = data.get("done_status")   # "Done" or "Not Done"

    if not student_id or not task_id or not done_status:
        return jsonify({"error": "Missing data"}), 400

    conn = sqlite3.connect("database.db")
    cur = conn.cursor()

    # Insert or update the row
    cur.execute("""
        INSERT INTO work_evaluation (student_id, task_id, done_status)
        VALUES (?, ?, ?)
        ON CONFLICT(student_id, task_id)
        DO UPDATE SET done_status = excluded.done_status
    """, (student_id, task_id, done_status))

    conn.commit()
    conn.close()

    return jsonify({"message": "Done status updated successfully!"})


# ✅ Export Work Submissions PDF
@app.route("/export_submissions_pdf/<int:task_id>")
def export_submissions_pdf(task_id):
    try:
        conn = sqlite3.connect("database.db")
        cur = conn.cursor()

        # ✅ Fetch TASK details
        cur.execute("""
            SELECT department, semester, category, title, description, deadline
            FROM tasks
            WHERE id = ?
        """, (task_id,))
        task = cur.fetchone()

        if not task:
            conn.close()
            return jsonify({"error": "Task not found"}), 404

        department, semester, category, title, description, deadline = task

        # ✅ Fetch all students in that department + submission status
        cur.execute("""
            SELECT s.register_no, s.full_name, s.batch_year,
                   COALESCE(sw.submitted_on, '--') as submitted_on,
                   CASE WHEN sw.file_path IS NOT NULL THEN 'Submitted' ELSE 'Not Submitted' END AS status
            FROM student s
            LEFT JOIN submitted_works sw 
                ON s.id = sw.student_id AND sw.task_id = ?
            WHERE s.department = ?
            ORDER BY s.register_no
        """, (task_id, department))
        
        students = cur.fetchall()
        conn.close()

        # ✅ Create PDF
        buffer = io.BytesIO()
        pdf = SimpleDocTemplate(buffer, pagesize=A4)
        styles = getSampleStyleSheet()
        elements = []

        # ✅ Title
        elements.append(Paragraph("<b>Work Submission Report</b>", styles['Title']))
        elements.append(Spacer(1, 12))

        # ✅ Work Details
        details = f"""
        <b>Department:</b> 
        {department}<br/>
        <b>Semester:</b> {semester}<br/>
        <b>Work Type:</b> {category}<br/>
        <b>Title:</b> {title}<br/>
        <b>Deadline:</b> {deadline}<br/><br/>
        """
        elements.append(Paragraph(details, styles['Normal']))

        # ✅ Table Header
        table_data = [["Register No", "Student Name", "Batch", "Status", "Submitted On"]]

        submitted_count = 0
        not_submitted_count = 0

        for row in students:
            reg_no, name, batch, submitted_on, status = row

            if status == "Submitted":
                submitted_count += 1
            else:
                not_submitted_count += 1

            table_data.append([reg_no, name, batch, status, submitted_on])

        # ✅ Create table
        table = Table(table_data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1976D2")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ]))
        elements.append(table)
        elements.append(Spacer(1, 12))

        # ✅ Summary
        summary = f"""
        <b>Total Students:</b> {len(students)}<br/>
        <b>Submitted:</b> {submitted_count}<br/>
        <b>Not Submitted:</b> {not_submitted_count}
        """
        elements.append(Paragraph(summary, styles['Normal']))

        pdf.build(elements)
        buffer.seek(0)

        # ✅ AUTO DOWNLOAD PDF
        return send_file(
            buffer,
            as_attachment=True,
            download_name=f"Work_Submissions_{task_id}.pdf",
            mimetype="application/pdf"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500
    

#mark entry
@app.route("/mark_entry")
def mark_entry():
    return render_template("teacher/mark_entry.html")

@app.route('/fetch_students', methods=['GET'])
def fetch_students():
    department = request.args.get('department')
    semester = request.args.get('semester')
    work_type = request.args.get('work_type')

    conn = sqlite3.connect("database.db")
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    query = """
    SELECT
        s.id AS student_id,
        s.register_no,
        s.full_name,
        t.id AS task_id,
        IFNULL(w.done_status, 'Not Done') AS done_status,
        IFNULL(m.mark, 0) AS mark,
        IFNULL(m.grade, '') AS grade
    FROM student s
    JOIN tasks t
        ON t.department = s.department
       AND t.semester = ?
       AND t.category = ?
    LEFT JOIN work_evaluation w
        ON w.student_id = s.id
       AND w.task_id = t.id
    LEFT JOIN student_mark m
        ON m.student_id = s.id
       AND m.task_id = t.id
    WHERE s.department = ?
    ORDER BY s.register_no
    """

    cur.execute(query, (semester, work_type, department))
    rows = cur.fetchall()
    conn.close()

    return jsonify([dict(r) for r in rows])


def calculate_grade(mark):
    if mark >= 90:
        return "A+"
    elif mark >= 80:
        return "A"
    elif mark >= 70:
        return "B+"
    elif mark >= 60:
        return "B"
    elif mark >= 50:
        return "C"
    else:
        return "F"


def grade_color(grade):
    return {
        "A+": "green",
        "A": "darkgreen",
        "B+": "blue",
        "B": "orange",
        "C": "brown",
        "F": "red"
    }.get(grade, "black")


@app.route('/save_marks', methods=['POST'])
def save_marks():
    data = request.json  # list of students

    conn = sqlite3.connect("database.db")
    cur = conn.cursor()

    for s in data:
        mark = int(s['mark'])
        grade = calculate_grade(mark)

        cur.execute("""
        INSERT INTO student_mark (student_id, task_id, mark, grade)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(student_id, task_id)
        DO UPDATE SET
            mark = excluded.mark,
            grade = excluded.grade
        """, (s['student_id'], s['task_id'], mark, grade))

    conn.commit()
    conn.close()

    return jsonify({"status": "success"})

@app.route('/export_marks_pdf')
def export_marks_pdf():
    department = request.args.get('department')
    semester = request.args.get('semester')
    work_type = request.args.get('work_type')

    conn = sqlite3.connect("database.db")
    cur = conn.cursor()

    query = """
    SELECT
        s.register_no,
        s.full_name,
        IFNULL(w.done_status, 'Not Done') AS done_status,
        IFNULL(m.mark, 0) AS mark,
        IFNULL(m.grade, '') AS grade
    FROM student s
    JOIN tasks t
        ON t.department = s.department
       AND t.semester = ?
       AND t.category = ?
    LEFT JOIN work_evaluation w
        ON w.student_id = s.id
       AND w.task_id = t.id
    LEFT JOIN student_mark m
        ON m.student_id = s.id
       AND m.task_id = t.id
    WHERE s.department = ?
    ORDER BY s.register_no
    """

    cur.execute(query, (semester, work_type, department))
    rows = cur.fetchall()
    conn.close()

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = []

    # Title
    elements.append(Paragraph("<b>Mark Entry Report</b>", styles["Title"]))
    elements.append(Paragraph(f"Department: {department}", styles["Normal"]))
    elements.append(Paragraph(f"Semester: {semester}", styles["Normal"]))
    elements.append(Paragraph(f"Work Type: {work_type}", styles["Normal"]))
    elements.append(Paragraph("<br/>", styles["Normal"]))

    # Table data
    table_data = [
        ["Register No", "Student Name", "Status", "Mark", "Grade"]
    ]

    for r in rows:
        table_data.append([
            r[0],
            r[1],
            "Uploaded" if r[2] == "Done" else "Not Uploaded",
            str(r[3]),
            r[4]
        ])

    table = Table(table_data, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('GRID', (0,0), (-1,-1), 1, colors.black),
        ('ALIGN', (3,1), (-1,-1), 'CENTER'),
        ('FONT', (0,0), (-1,0), 'Helvetica-Bold'),
    ]))

    elements.append(table)
    doc.build(elements)

    buffer.seek(0)
    return send_file(
        buffer,
        as_attachment=True,
        download_name="mark_entry_report.pdf",
        mimetype="application/pdf"
    )

# -----------------------------
# Lecture Notes Page (Student)
# -----------------------------
@app.route("/student_lecture_notes")
def student_lecture_notes():

    # 🔐 Login check
    if "student_id" not in session:
        return redirect(url_for("student_login"))

    student_id = session["student_id"]

    conn = sqlite3.connect("database.db")
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    # -----------------------------
    # Fetch student info
    # -----------------------------
    cur.execute("""
        SELECT full_name, department, register_no, batch_year
        FROM student
        WHERE id = ?
    """, (student_id,))
    student_row = cur.fetchone()

    if not student_row:
        conn.close()
        return "Student not found", 404

    # Convert to dict
    student = {
        "full_name": student_row["full_name"],
        "department": student_row["department"],
        "register_no": student_row["register_no"],
        "batch_year": int(student_row["batch_year"])
    }


# -----------------------------
# Semester calculation
# -----------------------------
    from datetime import datetime

# -----------------------------
# Semester calculation
# -----------------------------
    now = datetime.now()
    current_year = now.year
    current_month = now.month
    batch_year = int(student["batch_year"])  # admission year

# -----------------------------
# Determine academic year base
# -----------------------------
# Academic year starts in June
    if current_month >= 6:        # June–Dec
        academic_base_year = current_year
    elif current_month <= 3:      # Jan–Mar
        academic_base_year = current_year - 1
    else:                         # April–May (ignored)
        academic_base_year = current_year - 1

# -----------------------------
# Years completed since joining
# -----------------------------
    years_completed = academic_base_year - batch_year

# -----------------------------
# Semester calculation
# -----------------------------
    if 6 <= current_month <= 10:        # June – October
        semester_no = years_completed * 2 + 1
    elif current_month in (11, 12, 1, 2, 3):  # Nov – March
        semester_no = years_completed * 2 + 2
    else:                                   # April – May
        semester_no = years_completed * 2   # last completed semester

# Clamp semester between 1 and 6
    semester_no = max(1, min(semester_no, 6))
    semester = f"Semester {semester_no}"

    student["semester"] = semester


    # -----------------------------
    # Fetch lecture notes
    # -----------------------------
    cur.execute("""
        SELECT id, filename, filepath, upload_time
        FROM lecture_notes
        WHERE department = ?
          AND semester = ?
          AND hidden = 0
        ORDER BY upload_time DESC
    """, (student["department"], semester))

    notes = cur.fetchall()
    conn.close()
    # -----------------------------
    # Render page
    # -----------------------------
    return render_template(
        "student/lecture_notes.html",
        student=student,
        notes=notes
    )

@app.route("/view_lecture_note/<int:note_id>")
def view_lecture_note(note_id):
    if "student_id" not in session:
        return redirect(url_for("student_login"))

    conn = sqlite3.connect("database.db")
    cur = conn.cursor()

    cur.execute("""
        SELECT filepath
        FROM lecture_notes
        WHERE id = ? AND hidden = 0
    """, (note_id,))

    row = cur.fetchone()
    conn.close()

    if not row:
        abort(404)

    full_path = row[0]

    if not os.path.exists(full_path):
        abort(404)

    return send_file(
        full_path,
        as_attachment=False   # VIEW
    )

@app.route("/download_lecture_note/<int:note_id>")
def download_lecture_note(note_id):
    if "student_id" not in session:
        return redirect(url_for("student_login"))

    conn = sqlite3.connect("database.db")
    cur = conn.cursor()

    cur.execute("""
        SELECT filepath
        FROM lecture_notes
        WHERE id = ? AND hidden = 0
    """, (note_id,))

    row = cur.fetchone()
    conn.close()

    if not row:
        abort(404)

    full_path = row[0]

    if not os.path.exists(full_path):
        abort(404)

    return send_file(
        full_path,
        as_attachment=True    # DOWNLOAD
    )

# ================================
# Page Route
# ================================
@app.route("/student_attendance_details")
def student_attendance_details():
    return render_template("admin/attendance_details.html")


# ================================
# Fetch Attendance (Admin)
# ================================
@app.route("/get_attendance_admin")
def get_attendance_admin():

    department = request.args.get("department")
    study_year = request.args.get("batch_year")  # 1,2,3,4 from UI
    period = request.args.get("period")
    register_no = request.args.get("register_no")
    from_date = request.args.get("from_date")
    to_date = request.args.get("to_date")

    current_year = datetime.now().year

    query = """
        SELECT register_no, full_name, department, batch_year,
               period_number, subject, attendance_status, date, time
        FROM attendance
        WHERE 1=1
    """
    params = []

    if department:
        query += " AND department = ?"
        params.append(department)

    # 🔥 Study year → admission year mapping
    if study_year:
        admission_year = current_year - int(study_year) + 1
        query += " AND batch_year = ?"
        params.append(admission_year)

    if period:
        query += " AND period_number = ?"
        params.append(period)

    if register_no:
        query += " AND register_no = ?"
        params.append(register_no)

    if from_date and to_date:
        query += " AND date BETWEEN ? AND ?"
        params.extend([from_date, to_date])

    query += " ORDER BY date, period_number"

    conn = sqlite3.connect("database.db")
    cur = conn.cursor()
    cur.execute(query, params)
    rows = cur.fetchall()
    conn.close()

    data = []

    for r in rows:
        admission_year = int(r[3])  # ✅ FIX HERE
        calculated_year = (current_year - admission_year) + 1

        year_label = {
            1: "First Year",
            2: "Second Year",
            3: "Third Year",
            4: "Fourth Year"
        }.get(calculated_year, f"Year {calculated_year}")

        data.append({
            "register_no": r[0],
            "full_name": r[1],
            "department": r[2],
            "batch_year": year_label,
            "period": r[4],
            "subject": r[5],
            "status": r[6],
            "date": str(r[7]),
            "time": str(r[8])
        })

    return jsonify(data)

@app.route("/attendance_pdf")
def attendance_pdf():

    department   = request.args.get("department")
    study_year   = request.args.get("batch_year")
    period       = request.args.get("period")
    register_no  = request.args.get("register_no")
    from_date    = request.args.get("from_date")
    to_date      = request.args.get("to_date")

    current_year = datetime.now().year

    query = """
        SELECT register_no, full_name, department, batch_year,
               period_number, subject, attendance_status, date, time
        FROM attendance
        WHERE 1=1
    """
    params = []

    if department:
        query += " AND department = ?"
        params.append(department)

    if study_year:
        admission_year = current_year - int(study_year) + 1
        query += " AND batch_year = ?"
        params.append(admission_year)

    if period:
        query += " AND period_number = ?"
        params.append(period)

    if register_no:
        query += " AND register_no = ?"
        params.append(register_no)

    if from_date and to_date:
        query += " AND date BETWEEN ? AND ?"
        params.extend([from_date, to_date])

    query += " ORDER BY date, period_number"

    conn = sqlite3.connect("database.db")
    cur = conn.cursor()
    cur.execute(query, params)
    rows = cur.fetchall()
    conn.close()

    # ---------------- PDF ----------------

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = []

    elements.append(Paragraph("<b>Attendance Details Report</b>", styles["Title"]))
    elements.append(Paragraph(f"Generated on: {datetime.now().strftime('%d-%m-%Y')}", styles["Normal"]))
    elements.append(Paragraph(" ", styles["Normal"]))

    table_data = [[
        "Reg No", "Name", "Department", "Year",
        "Period", "Subject", "Status", "Date", "Time"
    ]]

    for r in rows:
        admission_year = int(r[3])
        year_no = (current_year - admission_year) + 1

        year_label = {
            1: "First Year",
            2: "Second Year",
            3: "Third Year",
            4: "Fourth Year"
        }.get(year_no, f"Year {year_no}")

        table_data.append([
            r[0], r[1], r[2], year_label,
            r[4], r[5], r[6],
            str(r[7]), str(r[8])
        ])

    table = Table(table_data, repeatRows=1)
    table.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.5, colors.black),
        ("BACKGROUND", (0,0), (-1,0), colors.lightblue),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("FONT", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 8),
        ("BOTTOMPADDING", (0,0), (-1,0), 8)
    ]))

    elements.append(table)
    doc.build(elements)

    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="attendance_report.pdf",
        mimetype="application/pdf"
    )


#  --- mark details ----
@app.route("/student_mark_details")
def student_mark_details():
    return render_template("admin/mark_details.html")

# ================================
# Fetch Marks (Admin)
# ================================

@app.route("/get_marks_admin")
def get_marks_admin():

    department = request.args.get("department")
    study_year = request.args.get("batch_year")  # 1,2,3,4
    register_no = request.args.get("register_no")

    current_year = datetime.now().year

    query = """
        SELECT
            s.register_no,
            s.full_name,
            s.department,
            s.batch_year,
            t.category,
            m.mark,
            m.grade,
            m.entered_on
        FROM student_mark m
        JOIN student s ON s.id = m.student_id
        JOIN tasks t ON t.id = m.task_id
        WHERE 1=1
    """
    params = []

    if department:
        query += " AND s.department = ?"
        params.append(department)

    # 🔥 STUDY YEAR → ADMISSION YEAR
    if study_year:
        admission_year = current_year - int(study_year) + 1
        query += " AND s.batch_year = ?"
        params.append(admission_year)

    if register_no:
        query += " AND s.register_no = ?"
        params.append(register_no)

    query += " ORDER BY m.entered_on DESC"

    conn = sqlite3.connect("database.db")
    cur = conn.cursor()
    cur.execute(query, params)
    rows = cur.fetchall()
    conn.close()

    data = []

    for r in rows:
        admission_year = int(r[3])
        year_no = (current_year - admission_year) + 1

        year_label = {
            1: "First Year",
            2: "Second Year",
            3: "Third Year",
            4: "Fourth Year"
        }.get(year_no, f"Year {year_no}")

        data.append({
            "register_no": r[0],
            "full_name": r[1],
            "department": r[2],
            "year": year_label,
            "category": r[4],        # ✅ CATEGORY SHOWN
            "mark": r[5],
            "grade": r[6],
            "entered_on": str(r[7])
        })

    return jsonify(data)


import sqlite3
from datetime import date
from flask import session, request, render_template

DB = "database.db"

def get_db():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    return conn

@app.route("/teacher_notification")
def teacher_notification():

    session["role"] = "Teacher"
    role = session["role"]

    search = request.args.get("search", "").strip()
    today = datetime.today()

    conn = sqlite3.connect("database.db")
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    cursor.execute("""
        SELECT *
        FROM notification
        WHERE status = 'Active'
        AND visibility IN ('Teacher', 'Teachers', 'All')
        ORDER BY posted_date DESC, posted_time DESC
    """)
    rows = cursor.fetchall()
    conn.close()

    notifications = []

    for n in rows:
        # Handle expiry safely
        if n["expiry_date"]:
            try:
                expiry = datetime.strptime(n["expiry_date"], "%d-%m-%Y")
                if today > expiry:
                    continue
            except ValueError:
                continue

        # Search filter
        if search:
            if search.lower() not in n["message"].lower() and search.lower() not in n["topic"].lower():
                continue

        notifications.append(n)

    return render_template(
        "teacher/notifications.html",
        notifications=notifications,
        role=role
    )


@app.route("/student_notification")
def student_notification():

    session["role"] = "Student"
    role = session["role"]

    search = request.args.get("search", "").strip()
    today = datetime.today()

    conn = sqlite3.connect("database.db")
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    cursor.execute("""
        SELECT *
        FROM notification
        WHERE status = 'Active'
        AND visibility IN ('Student', 'Students', 'All')
        ORDER BY posted_date DESC, posted_time DESC
    """)
    rows = cursor.fetchall()
    conn.close()

    notifications = []

    for n in rows:
        # Expiry check
        if n["expiry_date"]:
            try:
                expiry = datetime.strptime(n["expiry_date"], "%d-%m-%Y")
                if today > expiry:
                    continue
            except ValueError:
                continue

        # Search filter
        if search:
            if search.lower() not in n["message"].lower() and search.lower() not in n["topic"].lower():
                continue

        notifications.append(n)

    return render_template(
        "student/notifications.html",
        notifications=notifications,
        role=role
    )


from datetime import timedelta

@app.route('/student_attendance', methods=['GET', 'POST'])
def student_attendance():

    # 🔐 Login protection
    if 'student_id' not in session:
        flash('Please login first')
        return redirect(url_for('student_login'))

    # 🔹 Default: show only TODAY
    if request.method == 'GET':
        from_date = to_date = date.today().strftime('%Y-%m-%d')
    else:
        from_date = request.form.get('from_date')
        to_date = request.form.get('to_date')

    conn = sqlite3.connect('database.db')
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    # 🔹 Student details
    cursor.execute("""
        SELECT register_no, full_name
        FROM student
        WHERE id = ?
    """, (session['student_id'],))
    student = cursor.fetchone()

    register_no = student['register_no']

    # 🔹 Fetch attendance
    cursor.execute("""
        SELECT date, period_number, subject, attendance_status
        FROM attendance
        WHERE register_no = ?
        AND date BETWEEN ? AND ?
        ORDER BY date, period_number
    """, (register_no, from_date, to_date))

    records = cursor.fetchall()
    conn.close()

    # 🔹 Map attendance by date
    attendance_map = {}
    for r in records:
        d = r['date']
        if d not in attendance_map:
            attendance_map[d] = []
        attendance_map[d].append(r)

    # 🔹 Generate ALL dates in range
    attendance = {}
    start = datetime.strptime(from_date, '%Y-%m-%d')
    end = datetime.strptime(to_date, '%Y-%m-%d')

    overall_present = 0
    overall_total = 0

    current = start
    while current <= end:
        d_str = current.strftime('%Y-%m-%d')
        day_name = current.strftime('%A')
        weekday = current.weekday()  # 5 = Saturday, 6 = Sunday

        attendance[d_str] = {
            'day': day_name,
            'periods': {},
            'present_count': 0,
            'total_periods': 0
        }

        if weekday >= 5:
            # Weekend → 0/0
            pass
        else:
            attendance[d_str]['total_periods'] = 5
            overall_total += 5

            if d_str in attendance_map:
                for row in attendance_map[d_str]:
                    attendance[d_str]['periods'][row['period_number']] = {
                        'status': row['attendance_status'],
                        'subject': row['subject']
                    }

                    if row['attendance_status'] == 'Present':
                        attendance[d_str]['present_count'] += 1
                        overall_present += 1

        current += timedelta(days=1)

    return render_template(
        'student/attendance.html',
        student=student,
        attendance=attendance,
        from_date=from_date,
        to_date=to_date,
        overall_present=overall_present,
        overall_total=overall_total
    )


from flask import request, session, flash, redirect, url_for, render_template
from datetime import datetime
import sqlite3
@app.route('/student_attendance_graph', methods=['GET', 'POST'])
def student_attendance_graph():

    # ─────────────────────────────
    # 1. Login check
    # ─────────────────────────────
    if 'student_id' not in session:
        flash('Please login first')
        return redirect(url_for('student_login'))

    # ─────────────────────────────
    # 2. DB connection
    # ─────────────────────────────
    conn = sqlite3.connect('database.db')
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    # ─────────────────────────────
    # 3. Student details
    # ─────────────────────────────
    cursor.execute("""
        SELECT register_no, full_name, batch_year
        FROM student
        WHERE id = ?
    """, (session['student_id'],))
    student = cursor.fetchone()

    register_no = student['register_no']
    batch_year = int(student['batch_year'])

    # ─────────────────────────────
    # 4. Semester structure (6 semesters)
    # ─────────────────────────────
    semester_months = {
        1: (6, 10),    # Jun – Oct
        2: (11, 3),    # Nov – Mar
        3: (6, 10),
        4: (11, 3),
        5: (6, 10),
        6: (11, 3)
    }

    today = datetime.now()
    sem_dates = {}
    current_semester = 1
    sem_start_year = batch_year

    for sem in range(1, 7):
        start_month, end_month = semester_months[sem]
        start_date = datetime(sem_start_year, start_month, 1)

        end_year = sem_start_year + 1 if end_month < start_month else sem_start_year
        end_day = 28 if end_month == 2 else 30 if end_month in [4, 6, 9, 11] else 31
        end_date = datetime(end_year, end_month, end_day)

        sem_dates[sem] = (
            start_date.strftime('%Y-%m-%d'),
            end_date.strftime('%Y-%m-%d')
        )

        if start_date <= today <= end_date:
            current_semester = sem

        if end_month < start_month:
            sem_start_year += 1

    # ─────────────────────────────
    # 5. Semester selection
    # ─────────────────────────────
    semesters = ['All'] + [f'Semester {i}' for i in range(1, 7)]
    selected_semester = request.form.get(
        'semester',
        f'Semester {current_semester}'
    )

    # ─────────────────────────────
    # 6. Fetch attendance
    # ─────────────────────────────
    if selected_semester != 'All':
        sem_num = int(selected_semester.split(' ')[1])
        start_date, end_date = sem_dates[sem_num]

        cursor.execute("""
            SELECT attendance_status
            FROM attendance
            WHERE register_no = ?
              AND date BETWEEN ? AND ?
        """, (register_no, start_date, end_date))
    else:
        cursor.execute("""
            SELECT attendance_status
            FROM attendance
            WHERE register_no = ?
        """, (register_no,))

    records = cursor.fetchall()
    conn.close()

    # FINAL attendance percentage logic
# ─────────────────────────────
    # ─────────────────────────────
# CLEAN & SAFE attendance logic
# ─────────────────────────────
    # ─────────────────────────────
# FINAL attendance percentage logic
# ─────────────────────────────

    absent = sum(1 for r in records if r['attendance_status'] == 'Absent')
    late   = sum(1 for r in records if r['attendance_status'] == 'Late')

# -----------------------------
# Determine TOTAL CLASSES
# -----------------------------

# Case 1: Current semester (live / till today)
    if selected_semester == f"Semester {current_semester}":
        total_classes = len(records)

# Case 2: All semesters (fixed 200 per semester)
    elif selected_semester == 'All':
        total_classes = 200 * current_semester

# Case 3: Any specific semester (fixed 200)
    else:
        total_classes = 200

# -----------------------------
# Calculate percentages
# -----------------------------

    absent_pct = 0.0
    late_pct = 0.0

    if total_classes > 0 and (absent > 0 or late > 0):
        class_pct = 100 / total_classes
        absent_pct = round(absent * class_pct, 2)
        late_pct   = round(late * class_pct, 2)

    present_pct = round(100 - (absent_pct + late_pct), 2)

# Safety clamp
    present_pct = max(0, min(100, present_pct))

    percentages = {
        'Present': present_pct,
        'Absent': absent_pct,
        'Late': late_pct
    }
    # ─────────────────────────────
    # 8. Render graph
    # ─────────────────────────────
    return render_template(
        'student/graph.html',
        student=student,
        percentages=percentages,
        semesters=semesters,
        selected_semester=selected_semester
    )



if __name__ == '__main__':
    init_db() 
    app.run(debug=True)