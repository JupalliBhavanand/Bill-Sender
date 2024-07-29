import os
import uuid
from flask import Flask, request, redirect, url_for, render_template, flash
import pandas as pd
import pywhatkit as kit
import datetime
import time
import pyautogui as pg
import razorpay

app = Flask(__name__)
app.secret_key = "super_secret_key"
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Initialize Razorpay client
razorpay_client = razorpay.Client(auth=("your_key_id", "your_key_secret"))

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('upload.html')

@app.route('/upload_file', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part'
    file = request.files['file']
    if file.filename == '':
        return 'No selected file'
    if file and allowed_file(file.filename):
        filename = str(uuid.uuid4()) + '_' + file.filename
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        process_file(file_path)
        return 'Messages have been sent and Excel sheet updated.'
    return 'File not allowed'

def process_file(file_path):
    df = pd.read_excel(file_path)
    for index, row in df.iterrows():
        if not pd.isna(row.get('paid_date')):
            continue
        phone_number = str(row['phone_number']).split('.')[0]
        if not phone_number.startswith('+'):
            phone_number = '+91' + phone_number  # Assuming country code is +91
        message = f"Dear {row['name']}, your bill is {row['bill_amount']}."
        message += f"\n⏰⏰ please pay before 10th\n"
        message += f"please inform any mistakes in bill\n\n"
        message += f"🌝🌝Sir/Madam, please inform morning hours only for next day no milk on whatsapp.\n\n"
        message += f"\nThanking you from BHOOMA PRABHAKAR "
        try:
            kit.sendwhatmsg_instantly(phone_number, message)
            print(f"Message sent to {phone_number}")
            time.sleep(3)
            pg.hotkey('alt', 'f4')
            time.sleep(2)
        except Exception as e:
            print(f"Failed to send message to {phone_number}: {e}")
    df.to_excel(file_path, index=False)
    print("Excel sheet updated.")

if __name__ == "__main__":
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    app.run(debug=False,host='0.0.0.0')
