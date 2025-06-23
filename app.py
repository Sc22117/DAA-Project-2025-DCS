from flask import Flask, render_template, request, redirect, url_for, session, send_from_directory, flash
from werkzeug.utils import secure_filename
import os
import firebase_admin
from firebase_admin import credentials, firestore, db as firebase_db

 # It only initialize the app if it hasn't been initialized already
if not firebase_admin._apps:
    cred = credentials.Certificate("DCS(Firebase Key).json")
    firebase_admin.initialize_app(cred, {
        'databaseURL': 'https://dcs-simulation---testing-default-rtdb.asia-southeast1.firebasedatabase.app/'
    })

firestore_db = firestore.client()


# Flask App Setup
app = Flask(__name__)
app.secret_key = '92f7e2f9asca43b884c2dd0398fscc3b9f'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['RESULT_FOLDER'] = 'result'
ALLOWED_EXTENSIONS = {'xlsx', 'csv'}

# Ensure necessary folders exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['RESULT_FOLDER'], exist_ok=True)

# Check allowed file extension
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# ---------- Routes ----------
@app.route('/', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        # process login
        session['user'] = {
            'designation': request.form.get('designation'),
            'name': request.form.get('name')
        }
        return redirect(url_for('dashboard'))
    return render_template('login.html')

@app.route('/dashboard')
def dashboard():
    # optionally check if user logged in
    if 'user' not in session:
        return redirect(url_for('login'))
    return render_template('dashboard.html')

@app.route('/enquiry', methods=['GET', 'POST'])
def enquiry():
    if request.method == 'POST':
        search_type = request.form.get('option')
        query = request.form.get('query')
        institute = session.get('constraints', {}).get('institute')

        if not institute:
            flash("Institute not set. Please complete constraints first.")
            return redirect(url_for('constraints'))

        # Reference to the appropriate Firebase path
        db_ref = firebase_db.reference(f"{institute}/timetable")

        schedule = []

        if search_type == "Faculty":
            faculty_ref = db_ref.child("faculty").child(query)
            data = faculty_ref.get()
            if data:
                for day, periods in data.items():
                    for slot, detail in periods.items():
                        schedule.append(f"{day} - {slot}: {detail}")
        elif search_type == "Venue":
            venue_ref = db_ref.child("venue").child(query)
            data = venue_ref.get()
            if data:
                for day, periods in data.items():
                    for slot, detail in periods.items():
                        schedule.append(f"{day} - {slot}: {detail}")
        elif search_type == "Section":
            section_ref = db_ref.child("section").child(query)
            data = section_ref.get()
            if data:
                for day, periods in data.items():
                    for slot, detail in periods.items():
                        schedule.append(f"{day} - {slot}: {detail}")
        else:
            flash("Invalid option selected.")
            return redirect(url_for('enquiry'))

        return render_template('enquiry_result.html', schedule=schedule)

    return render_template('enquiry.html')

@app.route('/timetable')
def timetable():
    # This can redirect to your upload page or render a timetable management page
    return redirect(url_for('constraints'))


@app.route('/constraints', methods=['GET', 'POST'])
def constraints():
    if request.method == 'POST':
        # Store all constraint variables in session for next steps
        session['constraints'] = {
            'institute': request.form['institute'],
            'start_time': request.form['start_time'],
            'end_time': request.form['end_time'],
            'period_duration': request.form['period_duration'],
            'num_breaks': request.form['num_breaks'],
            'break_duration': request.form['break_duration'],
            'num_days': request.form['num_days']
        }
        return redirect(url_for('upload'))
    return render_template('constraints.html')

@app.route('/upload', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename("input.xlsx")
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            session['uploaded_file_path'] = file_path
            return redirect(url_for('process'))
        else:
            flash('Invalid file format. Allowed: xlsx, csv')
            return redirect(request.url)
    return render_template('upload.html')

@app.route('/process')
def process():
    try:
        print("Process started")  # Debug print to check route is hit
        
        from DAA_project_code import generate_timetable_pipeline

        constraints = session.get('constraints')
        file_path = session.get('uploaded_file_path')

        if not constraints or not file_path:
            return "Missing required data. Please complete previous steps.", 400

        institute = constraints['institute']
        start_time = constraints['start_time']
        end_time = constraints['end_time']
        duration = int(constraints['period_duration'])
        num_breaks = int(constraints['num_breaks'])
        break_duration = int(constraints['break_duration'])
        num_days = int(constraints['num_days'])

        output_file_path = os.path.join(app.config['RESULT_FOLDER'], 'final_report.xlsx')

        # Debug: Indicate before running the pipeline
        print(f"Running pipeline for institute: {institute}")

        # Call the combined pipeline function
        generate_timetable_pipeline(
            institute=institute,
            start_time_str=start_time,
            end_time_str=end_time,
            period_duration=duration,
            num_breaks=num_breaks,
            break_duration=break_duration,
            num_days=num_days,
            file_path=file_path,
            output_path=output_file_path
        )

        print("Pipeline completed successfully")

        return redirect(url_for('download_file'))

    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        print("Error during processing:")
        print(tb)  # Print full traceback for debugging
        return f"Error during processing: {e}", 500



@app.route('/download')
def download_file():
    filename = 'final_report.xlsx'
    file_path = os.path.join(app.config['RESULT_FOLDER'], filename)
    if not os.path.exists(file_path):
        return "Report not found, please run the process first.", 404
    return send_from_directory(app.config['RESULT_FOLDER'], filename, as_attachment=True)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

if __name__ == '__main__':
    app.run(debug=True)
