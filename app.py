from flask import Flask, render_template, request, redirect, url_for, session
from flask import send_from_directory
from werkzeug.utils import secure_filename
import os
import firebase_admin
from firebase_admin import credentials, firestore

# Firebase Initialization
cred = credentials.Certificate(r"firebase_credentials.json")
firebase_admin.initialize_app(cred)
db = firestore.client()

# Flask App Setup
app = Flask(__name__)
app.secret_key = '92f7e2f9asca43b884c2dd0398fscc3b9f'
app.config['UPLOAD_FOLDER'] = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'csv'}

# Ensure upload folder exists
if not os.path.exists('uploads'):
    os.makedirs('uploads')

# Check allowed file extension
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# ---------- Routes ----------

@app.route('/test')
def test():
    return "Flask is working!"

@app.route('/static/<path:filename>')
def serve_static(filename):
    return send_from_directory('static', filename)

@app.route('/', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        designation = request.form['designation']
        name = request.form['name']
        password = request.form['password']

        session['user'] = {'designation': designation, 'name': name}

        db.collection('users').document(name).set({
            'designation': designation,
            'password': password
        }, merge=True)

        return redirect(url_for('dashboard'))
    return render_template('login.html')

@app.route('/dashboard')
def dashboard():
    if 'user' not in session:
        return redirect(url_for('login'))
    return render_template('dashboard.html')

@app.route('/constraints', methods=['GET', 'POST'])
def constraints():
    if 'user' not in session:
        return redirect(url_for('login'))

    if request.method == 'POST':
        data = {
            'institute_name': request.form['institute_name'],
            'start_time': request.form['start_time'],
            'end_time': request.form['end_time'],
            'break_duration': int(request.form['break_duration']),
            'num_short_breaks': int(request.form['num_short_breaks']),
            'short_break_duration': int(request.form['short_break_duration'])
        }

        db.collection('constraints').document(session['user']['name']).set(data)
        return redirect(url_for('timetable'))  # go to file upload after saving

    return render_template('constraints.html')


@app.route('/timetable', methods=['GET', 'POST'])
def timetable():
    if 'user' not in session:
        return redirect(url_for('login'))

    if request.method == 'POST':
        institute_name = request.form.get('institute_name')
        start_time = request.form.get('start_time')
        end_time = request.form.get('end_time')
        period_duration = request.form.get('period_duration')
        num_short_breaks = request.form.get('num_short_breaks')
        short_break_duration = request.form.get('short_break_duration')

        # Save these values to session or database
        session['constraints'] = {
            'institute_name': institute_name,
            'start_time': start_time,
            'end_time': end_time,
            'period_duration': period_duration,
            'num_short_breaks': num_short_breaks,
            'short_break_duration': short_break_duration
        }

        return redirect(url_for('upload'))

    return render_template('constraints.html')

@app.route('/upload', methods=['GET', 'POST'])
def upload():
    if 'user' not in session:
        return redirect(url_for('login'))

    if request.method == 'POST':
        file = request.files['file']
        if file and allowed_file(file.filename):
            filename = secure_filename(f"{session['user']['name']}.xlsx")
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            return redirect(url_for('timetable'))  # or a success page

        return "Invalid file format."

    return render_template('upload.html')


@app.route('/enquiry', methods=['GET', 'POST'])
def enquiry():
    if 'user' not in session:
        return redirect(url_for('login'))

    if request.method == 'POST':
        option = request.form['option']  # Faculty, Venue, Section
        query = request.form['query']

        schedule_ref = db.collection('schedules').document(option).collection(query)
        docs = schedule_ref.stream()
        schedule = [doc.to_dict() for doc in docs]

        return render_template('enquiry_result.html', schedule=schedule)

    return render_template('enquiry.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

if __name__ == '__main__':
    app.run(debug=True)
