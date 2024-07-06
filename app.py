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
razorpay_client = razorpay.Client(auth=("rzp_live_pmMQy6DshGmdX5", "eYEcYAIitDd8c6RdtYrixe63"))

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def calculate_attendance(row, df):
    attendance_columns = df.columns[2:32]  # Adjust based on the actual range of attendance columns
    attendance_count = row[attendance_columns].count()
    return attendance_count, attendance_columns

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = str(uuid.uuid4()) + '_' + file.filename
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            process_file(file_path)
            flash("Messages have been sent and Excel sheet updated.")
            return redirect(request.url)
    return render_template('upload.html')

def process_file(file_path):
    df = pd.read_excel(file_path)

    for index, row in df.iterrows():
        phone_number = str(row['phone_number']).split('.')[0]
        if not phone_number.startswith('+'):
            phone_number = '+' + phone_number

        name = row['EMRALD BLOCK ( BHOOMA SRI HARI )']
        amount_due = row['Bill']
        
        if pd.isna(amount_due):
            print(f"Skipping row {index} for {name} due to NaN amount_due.")
            continue
        
        try:
            amount_due = float(amount_due)  # Ensure amount_due is a numeric value
        except ValueError:
            print(f"Invalid amount_due for {name}. Skipping...")
            continue

        attendance, attendance_columns = calculate_attendance(row, df)
        june_attendance = row['last month Atten']
        july_actual = row['last month Act']
        june_diff = row['month diff']
        Aug_actual = row['this month Act']
        Aug_for_bill = row['this month for bill']
        Aug_milk_bill = row['Amt']
        Del_charge = row['Del charge']
        paper = row['Paper']
        old_bal = row['old bal']
        old_extra = row['old Extra']
        extraaaaa = row['extraaaaa ']

        # Create Razorpay order
        order_amount = int(amount_due * 100)  # Convert to paise
        order_currency = 'INR'
        order_receipt = str(uuid.uuid4())
        order_data = {
            'amount': order_amount,
            'currency': order_currency,
            'receipt': order_receipt,
            'payment_capture': '1'
        }

        try:
            razorpay_order = razorpay_client.order.create(data=order_data)
            payment_link = f"https://rzp.io/i/{razorpay_order['id']}"
            df.at[index, 'order_id'] = razorpay_order['id']
        except Exception as e:
            print(f"Failed to create Razorpay order: {e}")
            payment_link = "Payment link could not be generated."

        message = f"                   SRI LAXMI MILK AND NEWSPAPER SUPPLY               \n"
        message += f"Hello {name},🙋\n\n   "
        message += "        🗓🗓Daily attendance:\n"
        
        for date in attendance_columns:
            try:
                date_obj = pd.to_datetime(date)
                status = row[date]
                message += f"{date_obj.strftime('%Y-%m-%d')}: {status}\n"
            except Exception as e:
                print(f"Error processing date {date}: {e}")

        message += f"  Your last month attendance is: {june_attendance}\n"
        message += f"  Your this month actual attendance is: {july_actual}\n"
        message += f"  Your last month attendance difference is: {june_diff}\n"
        message += f"  Your month actual bill is: {Aug_actual}\n"
        message += f"  Your bill for the month is: {Aug_for_bill}\n"
        message += f"  Your paper bill is: {paper}\n"
        message += f"  Delivery charges is: {Del_charge}\n"
        message += f"  Your old balance is: {old_bal}\n"
        message += f"  Extra amount is: {extraaaaa}\n"
        message += f"  \nPlease make a payment of {amount_due}.  \n"
        message += f"  Payment link: 6305920076@ybl\n\n"
        message += "    🌝🌝Sir/Madam, please inform morning hours only for next day no milk.\n"
        message += "   🕣🕣🕣Please pay before 10th. Thank you\n"

        try:
            kit.sendwhatmsg_instantly(phone_number, message)
            print(f"Message sent to {phone_number}")
            time.sleep(6)

            pg.hotkey('alt', 'f4')
            time.sleep(2)

        except Exception as e:
            print(f"Failed to send message to {phone_number}: {e}")

    df.to_excel(file_path, index=False)
    print("Excel sheet updated.")

@app.route('/webhook', methods=['POST'])
def webhook():
    payload = request.get_json()
    event = payload.get('event')

    if event == 'payment.captured':
        payment_id = payload['payload']['payment']['entity']['id']
        order_id = payload['payload']['payment']['entity']['order_id']
        payment_date = datetime.datetime.now().strftime('%Y-%m-%d')

        update_paid_date(order_id, payment_date)
        return 'Webhook received', 200
    else:
        return 'Unhandled event', 400

def update_paid_date(order_id, payment_date):
    file_path = find_file_by_order_id(order_id)
    if file_path:
        df = pd.read_excel(file_path)
        df.loc[df['order_id'] == order_id, 'paid_date'] = payment_date
        df.to_excel(file_path, index=False)
        print(f"Updated paid date for order {order_id}")

def find_file_by_order_id(order_id):
    for root, dirs, files in os.walk(app.config['UPLOAD_FOLDER']):
        for file in files:
            file_path = os.path.join(root, file)
            df = pd.read_excel(file_path)
            if order_id in df['order_id'].values:
                return file_path
    return None

if __name__ == "__main__":
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    app.run(debug=True)
