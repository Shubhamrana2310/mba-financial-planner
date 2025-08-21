import xlwings as xw

def financial_planner(user_inputs):
    # Open Excel workbook
    wb = xw.Book('FP.xlsx')
    
    # Access sheets
    mba_year_sheet = wb.sheets['MBA_year']
    post_mba_sheet = wb.sheets['post_MBA']
    
    # -- MBA_year inputs --
    print("Please input your expected first time expenditures. Following data has already been catered for:- \n'Note: All inputs to be in Euro and digits only.'")
    print("Administration Fee: €500, \nSecurity Deposit: €2,000, \nSetting up: €500, \nFLight: €500, \nVisa: €245, \nHealth Insurance: €150, \nBicycle: €150, \nShipping: €800")
    mba_year_sheet.range('B21').value = user_inputs.get('others_f', 0)  # input for other one-time expenditures
    print("Please input your expected expenditures you'll incure on monthly basis. Following data has already been catered for.")
    print("Rent (including utilities): €1,100,\n  Food: €300, \n Transport: €100, \n  Outings: €200, \n  Gym: €100, \n  Saloon: €200, \n  Misc: €150,")
    mba_year_sheet.range('B34').value = user_inputs.get('others_m', 0) # input for other monthly expenditures
    print("Mention your savings target for a month during MBA year.")
    # mba_year_sheet.range('D59').value = user_inputs.get('target_savings_mba', 0) # input for target savings during MBA year
    print("Sanctioned Loan Amount is ₹50,00,000. \n Please input the amount you want to avail as loan from this amount, enter only digits.")
    mba_year_sheet.range('C42').value = user_inputs.get('loan_disbursed', 5000000)  # input for loan amount
    # Add more inputs as needed, mapping your user_inputs keys to correct cells

    # -- post_MBA inputs --
    print("Please input your expected yealry salary post MBA.")
    post_mba_sheet.range('B3').value = user_inputs.get('yearly_salary', 80000)
    print("Please input your expected one time expenditures. Following data has already been catered for.")
    print("Administration Fee: €500, \n  Security Deposit: €2,000, \n  Shifting: €500, \n  VISA: €500, \n Visa: €250")
    post_mba_sheet.range('B11').value = user_inputs.get('others_o', 0)
    print("Please input your expected expenditures you'll incure on monthly basis. Following data has already been catered for.")
    print("Rent (including utilities): €1,800,\n  Food: €300, \n Transport: €200, \n  Outings: €300, \n  Gym: €200, \n  Saloon: €200, \n  Misc: €200,")
    post_mba_sheet.range('B23').value = user_inputs.get('others_m2', 0)
    print("Please enter any monthly savings you'd like to keep aside. These will kept aside for your future savings and not included in loan repayment.")
    post_mba_sheet.range('B25').value = user_inputs.get('personal_savings', 100)  # input for personal savings apart from loan repayment amount
    # Add more inputs here similarly
    
    # Let Excel calculate formulas
    mba_year_sheet.api.Calculate()
    post_mba_sheet.api.Calculate()
    
    # -- Extract Outputs (example) --
    net_monthly_savings_1 = mba_year_sheet.range('B58').value  # Change these to actual output cells
    net_monthly_savings_2 = post_mba_sheet.range('B41').value
    loan_repayment_time_years = post_mba_sheet.range('C55').value
    
    # Print outputs
    print("Total MBA Year Savings per Month:", net_monthly_savings_1)
    print("Total Post MBA Savings per month:", net_monthly_savings_2)
    print("Loan Repayment Time (Years):", loan_repayment_time_years)
    
    # Save and close workbook
    wb.save()
    wb.close()
    
    # Return output dictionary for use in app or further processing
    return {
        "total_savings_mba": net_monthly_savings_1,
        "total_savings_post_mba": net_monthly_savings_2,
        "loan_repayment_time_years": loan_repayment_time_years
    }


if __name__ == "__main__":
    # Example inputs, replace with actual user inputs later
    example_inputs = {
        'mba_expenditure': 2250,
        'loan_amount': 49150,
        'target_savings_mba': 1000,
        'monthly_salary': 6666.67,
        'monthly_expenses_post_mba': 3300,
        'target_savings_post_mba': 500
    }
    
    results = financial_planner(example_inputs)
    print("Results Dictionary:", results)
