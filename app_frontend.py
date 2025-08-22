import streamlit as st
import requests

st.title("MBA/Post-MBA Financial Planner")

st.write("Enter your details to see your savings and loan projections. Your data never leaves your device except to call the calculation API.")

others_f = st.number_input("Other first-time MBA expenses (Euro)", value=1000)
others_m = st.number_input("Other monthly MBA expenses (Euro)", value=200)
loan_disbursed = st.number_input("Loan disbursed (INR)", value=4000000)
yearly_salary = st.number_input("Expected yearly salary after MBA (Euro)", value=90000)
others_o = st.number_input("Other one-time post-MBA expenses (Euro)", value=1000)
others_m2 = st.number_input("Other monthly post-MBA expenses (Euro)", value=300)
personal_savings = st.number_input("Personal monthly savings after MBA (Euro)", value=150)

# Input your API endpoint here. That means every time you deploy the api_server.py, update this URL. 
# For testing, you can use a service like ngrok to expose your local server. 
api_url = st.text_input("API endpoint (e.g., https://abc123.ngrok.io/calculate)", value="https://fda01ab6ae1d.ngrok-free.app/calculate")

if st.button("Calculate My Plan"):
    input_data = {
        "others_f": others_f,
        "others_m": others_m,
        "loan_disbursed": loan_disbursed,
        "yearly_salary": yearly_salary,
        "others_o": others_o,
        "others_m2": others_m2,
        "personal_savings": personal_savings
    }   
    try:
        response = requests.post(api_url, json=input_data, timeout=30)
        if response.ok:
            st.success("Calculation Results:")
            st.json(response.json())
        else:
            st.error(f"API error: {response.text}")
    except Exception as e:
        st.error(f"Request failed: {str(e)}")
