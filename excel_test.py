import xlwings as xw

def test_excel():
    # Open your Excel file (make sure the filename is correct)
    wb = xw.Book('FP.xlsx')
    sheet = wb.sheets['MBA_year']  # Change if your sheet has a different name

    # Example: Put values into cells where inputs go
    sheet.range('C42').value = 4000000      # Example input 1 - Loan Amount 
    sheet.range('B12').value = 3000      # Example input 2 - Security Deposit

    # Ask Excel to calculate formulas after input
    sheet.api.Calculate()

    # Read output values from calculated cells
    output1 = sheet.range('B58').value # Net Monthly Savings
    output2 = sheet.range('C51').value # Total Monthly Payment fo Loan EMI

    print(f"Net Monthly Savings: {output1}")
    print(f"Monthly EMI: {output2}")

    # Save and close the workbook
    wb.save()
    wb.close()

if __name__ == "__main__":
    test_excel()
