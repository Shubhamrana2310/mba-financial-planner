from flask import Flask, request, jsonify
import xlwings as xw

# Your existing financial_planner function goes here or import it if in another file
def financial_planner(user_inputs):
    wb = xw.Book('FP.xlsx')
    mba_year_sheet = wb.sheets['MBA_year']
    post_mba_sheet = wb.sheets['post_MBA']
    
    # Set Excel inputs from user_inputs dictionary
    mba_year_sheet.range('B21').value = user_inputs.get('others_f', 0)
    mba_year_sheet.range('B34').value = user_inputs.get('others_m', 0)
    mba_year_sheet.range('C42').value = user_inputs.get('loan_disbursed', 5000000)
    
    post_mba_sheet.range('B3').value = user_inputs.get('yearly_salary', 80000)
    post_mba_sheet.range('B11').value = user_inputs.get('others_o', 0)
    post_mba_sheet.range('B23').value = user_inputs.get('others_m2', 0)
    post_mba_sheet.range('B25').value = user_inputs.get('personal_savings', 100)
    
    # Let Excel calculate
    mba_year_sheet.api.Calculate()
    post_mba_sheet.api.Calculate()
    
    # Extract output cells
    net_monthly_savings_1 = mba_year_sheet.range('B58').value
    net_monthly_savings_2 = post_mba_sheet.range('B41').value
    loan_repayment_time_years = post_mba_sheet.range('C55').value
    
    wb.save()
    wb.close()
    
    return {
        "total_savings_mba": net_monthly_savings_1,
        "total_savings_post_mba": net_monthly_savings_2,
        "loan_repayment_time_years": loan_repayment_time_years
    }

# Set up Flask app
app = Flask(__name__)

# Define API endpoint /calculate for POST requests
@app.route('/calculate', methods=['POST'])
def calculate():
    # Extract JSON data sent by client
    user_inputs = request.json
    
    # Call your existing processing function with user inputs
    results = financial_planner(user_inputs)
    
    # Send output back as JSON
    return jsonify(results)

if __name__ == '__main__':
    # Run the API server on your local machine, accessible on port 5000
    app.run(host='0.0.0.0', port=5000, debug=True)
