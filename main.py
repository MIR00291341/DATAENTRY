import pandas as pd
from flask import Flask, render_template, request

app = Flask(__name__)

# Replace 'your_excel_file.xlsx' with the actual path to your Excel file
EXCEL_FILE = 'data_entry.xlsx'

try:
    # Try to load existing data from Excel
    df = pd.read_excel(EXCEL_FILE)
except FileNotFoundError:
    # If the file doesn't exist, create a new DataFrame with the required columns
    df = pd.DataFrame(columns=['Date', 'Buyer Name', 'DC No.', 'Item Name', 'Qty'])

@app.route('/', methods=['GET', 'POST'])
def data_entry():
    if request.method == 'POST':
        date = request.form['date']
        buyer_name = request.form['buyer_name']
        dc_no = request.form['dc_no']
        item_name = request.form['item_name']
        qty = request.form['qty']

        new_data = pd.DataFrame({'Date': [date], 'Buyer Name': [buyer_name], 
                                 'DC No.': [dc_no], 'Item Name': [item_name], 'Qty': [qty]})
        df = pd.concat([df, new_data], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False) # Save the updated DataFrame to Excel
        return "Data saved successfully!"

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
