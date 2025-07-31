import math
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
import io

class ImprovedProjectFinancingCalculator:
    def __init__(self):
        self.investment_options = {
            'Fixed Deposit': {
                'default_return': 7.0,
                'liquidity': 'Medium (penalty for early withdrawal)',
                'tax_treatment': 'Interest taxed as income',
                'risk': 'Very Low (insured up to ₹5L)',
                'compounding': 'quarterly'
            },
            'Liquid Mutual Funds': {
                'default_return': 6.5,
                'liquidity': 'High (T+1 settlement)',
                'tax_treatment': 'LTCG 20% after 3 years',
                'risk': 'Low',
                'compounding': 'daily'
            },
            'Debt Mutual Funds': {
                'default_return': 7.5,
                'liquidity': 'Medium (T+1 to T+3)',
                'tax_treatment': 'LTCG 20% after 3 years',
                'risk': 'Low to Medium',
                'compounding': 'daily'
            },
            'Equity Mutual Funds': {
                'default_return': 12.0,
                'liquidity': 'High (T+1 settlement)',
                'tax_treatment': 'LTCG 10% after 1 year',
                'risk': 'High',
                'compounding': 'daily'
            }
        }

    def calculate_emi(self, principal, annual_rate, tenure_years):
        """Calculate EMI using standard formula"""
        if principal <= 0 or tenure_years <= 0:
            return 0
        
        monthly_rate = annual_rate / (12 * 100)
        total_months = tenure_years * 12
        
        if annual_rate == 0:
            return principal / total_months
        
        emi = (principal * monthly_rate * (1 + monthly_rate)**total_months) / \
              ((1 + monthly_rate)**total_months - 1)
        
        return emi

    def calculate_investment_value(self, principal, annual_rate, years, compounding='quarterly'):
        """Calculate future value of investment"""
        if principal <= 0 or years <= 0:
            return principal
            
        rate = annual_rate / 100
        
        if compounding == 'daily':
            return principal * (1 + rate/365)**(365 * years)
        elif compounding == 'monthly':
            return principal * (1 + rate/12)**(12 * years)
        elif compounding == 'quarterly':
            return principal * (1 + rate/4)**(4 * years)
        else:  # annual
            return principal * (1 + rate)**years

    def calculate_scenarios(self, inputs):
        """Calculate three financing scenarios with clear definitions"""
        
        # ==========================================
        # SCENARIO 1: MAXIMUM OWN FUNDING
        # Use maximum possible own capital, minimize loan
        # ==========================================
        scenario1_own_used = min(inputs['own_capital'], inputs['project_cost'])
        scenario1_loan_needed = max(0, inputs['project_cost'] - scenario1_own_used)
        scenario1_remaining_capital = inputs['own_capital'] - scenario1_own_used
        
        # Loan calculations
        scenario1_emi = self.calculate_emi(
            scenario1_loan_needed * 100000,  # Convert to actual rupees
            inputs['loan_rate'], 
            inputs['loan_tenure']
        )
        scenario1_total_payments = scenario1_emi * 12 * inputs['loan_tenure']
        scenario1_total_interest = scenario1_total_payments - (scenario1_loan_needed * 100000)
        
        # Tax benefit on interest (if applicable)
        scenario1_interest_tax_benefit = 0
        if inputs['loan_interest_deductible']:
            scenario1_interest_tax_benefit = scenario1_total_interest * (inputs['tax_rate'] / 100)
        
        scenario1_net_interest_cost = scenario1_total_interest - scenario1_interest_tax_benefit
        
        # Investment of remaining capital (if any)
        scenario1_investment_value = scenario1_remaining_capital
        scenario1_investment_gain = 0
        scenario1_investment_gain_after_tax = 0
        
        if scenario1_remaining_capital > 0:
            scenario1_investment_value = self.calculate_investment_value(
                scenario1_remaining_capital,
                inputs['investment_return'],
                inputs['loan_tenure'],
                self.investment_options[inputs['investment_type']]['compounding']
            )
            scenario1_investment_gain = scenario1_investment_value - scenario1_remaining_capital
            scenario1_investment_gain_after_tax = scenario1_investment_gain * (1 - inputs['tax_rate']/100)
        
        # TOTAL CASH OUTFLOW = Own capital used + Net interest paid - Investment gains
        scenario1_total_cash_outflow = scenario1_own_used + (scenario1_net_interest_cost/100000) - scenario1_investment_gain_after_tax
        
        
        # ==========================================
        # SCENARIO 2: MAXIMUM LEVERAGE
        # Take full loan, invest all own capital
        # ==========================================
        scenario2_own_used = 0
        scenario2_loan_needed = inputs['project_cost']
        scenario2_capital_invested = inputs['own_capital']
        
        # Loan calculations
        scenario2_emi = self.calculate_emi(
            scenario2_loan_needed * 100000,
            inputs['loan_rate'],
            inputs['loan_tenure']
        )
        scenario2_total_payments = scenario2_emi * 12 * inputs['loan_tenure']
        scenario2_total_interest = scenario2_total_payments - (scenario2_loan_needed * 100000)
        
        # Tax benefit on interest
        scenario2_interest_tax_benefit = 0
        if inputs['loan_interest_deductible']:
            scenario2_interest_tax_benefit = scenario2_total_interest * (inputs['tax_rate'] / 100)
        
        scenario2_net_interest_cost = scenario2_total_interest - scenario2_interest_tax_benefit
        
        # Investment calculations
        scenario2_investment_value = self.calculate_investment_value(
            scenario2_capital_invested,
            inputs['investment_return'],
            inputs['loan_tenure'],
            self.investment_options[inputs['investment_type']]['compounding']
        )
        scenario2_investment_gain = scenario2_investment_value - scenario2_capital_invested
        scenario2_investment_gain_after_tax = scenario2_investment_gain * (1 - inputs['tax_rate']/100)
        
        # TOTAL CASH OUTFLOW = Net interest paid - Investment gains
        scenario2_total_cash_outflow = (scenario2_net_interest_cost/100000) - scenario2_investment_gain_after_tax
        
        
        # ==========================================
        # SCENARIO 3: BALANCED APPROACH
        # Custom mix of own funding and loan
        # ==========================================
        scenario3_own_used = min(inputs['custom_own_contribution'], inputs['project_cost'], inputs['own_capital'])
        scenario3_loan_needed = max(0, inputs['project_cost'] - scenario3_own_used)
        scenario3_remaining_capital = max(0, inputs['own_capital'] - scenario3_own_used)
        
        # Loan calculations
        scenario3_emi = self.calculate_emi(
            scenario3_loan_needed * 100000,
            inputs['loan_rate'],
            inputs['loan_tenure']
        ) if scenario3_loan_needed > 0 else 0
        
        scenario3_total_payments = scenario3_emi * 12 * inputs['loan_tenure']
        scenario3_total_interest = scenario3_total_payments - (scenario3_loan_needed * 100000)
        
        # Tax benefit on interest
        scenario3_interest_tax_benefit = 0
        if inputs['loan_interest_deductible']:
            scenario3_interest_tax_benefit = scenario3_total_interest * (inputs['tax_rate'] / 100)
        
        scenario3_net_interest_cost = scenario3_total_interest - scenario3_interest_tax_benefit
        
        # Investment of remaining capital
        scenario3_investment_value = scenario3_remaining_capital
        scenario3_investment_gain = 0
        scenario3_investment_gain_after_tax = 0
        
        if scenario3_remaining_capital > 0:
            scenario3_investment_value = self.calculate_investment_value(
                scenario3_remaining_capital,
                inputs['investment_return'],
                inputs['loan_tenure'],
                self.investment_options[inputs['investment_type']]['compounding']
            )
            scenario3_investment_gain = scenario3_investment_value - scenario3_remaining_capital
            scenario3_investment_gain_after_tax = scenario3_investment_gain * (1 - inputs['tax_rate']/100)
        
        # TOTAL CASH OUTFLOW = Own capital used + Net interest paid - Investment gains
        scenario3_total_cash_outflow = scenario3_own_used + (scenario3_net_interest_cost/100000) - scenario3_investment_gain_after_tax
        
        
        # ==========================================
        # COMPILE RESULTS
        # ==========================================
        results = {
            'scenario1': {
                'name': 'Maximum Own Funding',
                'own_capital_used': scenario1_own_used,
                'loan_amount': scenario1_loan_needed,
                'remaining_capital_invested': scenario1_remaining_capital,
                'monthly_emi': scenario1_emi,
                'total_interest_gross': scenario1_total_interest / 100000,
                'interest_tax_benefit': scenario1_interest_tax_benefit / 100000,
                'total_interest_net': scenario1_net_interest_cost / 100000,
                'investment_maturity_value': scenario1_investment_value,
                'investment_gain_gross': scenario1_investment_gain,
                'investment_gain_net': scenario1_investment_gain_after_tax,
                'total_cash_outflow': scenario1_total_cash_outflow
            },
            'scenario2': {
                'name': 'Maximum Leverage',
                'own_capital_used': scenario2_own_used,
                'loan_amount': scenario2_loan_needed,
                'remaining_capital_invested': scenario2_capital_invested,
                'monthly_emi': scenario2_emi,
                'total_interest_gross': scenario2_total_interest / 100000,
                'interest_tax_benefit': scenario2_interest_tax_benefit / 100000,
                'total_interest_net': scenario2_net_interest_cost / 100000,
                'investment_maturity_value': scenario2_investment_value,
                'investment_gain_gross': scenario2_investment_gain,
                'investment_gain_net': scenario2_investment_gain_after_tax,
                'total_cash_outflow': scenario2_total_cash_outflow
            },
            'scenario3': {
                'name': 'Balanced Approach',
                'own_capital_used': scenario3_own_used,
                'loan_amount': scenario3_loan_needed,
                'remaining_capital_invested': scenario3_remaining_capital,
                'monthly_emi': scenario3_emi,
                'total_interest_gross': scenario3_total_interest / 100000,
                'interest_tax_benefit': scenario3_interest_tax_benefit / 100000,
                'total_interest_net': scenario3_net_interest_cost / 100000,
                'investment_maturity_value': scenario3_investment_value,
                'investment_gain_gross': scenario3_investment_gain,
                'investment_gain_net': scenario3_investment_gain_after_tax,
                'total_cash_outflow': scenario3_total_cash_outflow
            }
        }
        
        # Find best scenario
        cash_outflows = [
            results['scenario1']['total_cash_outflow'],
            results['scenario2']['total_cash_outflow'],
            results['scenario3']['total_cash_outflow']
        ]
        
        best_scenario_index = cash_outflows.index(min(cash_outflows))
        best_scenario = ['scenario1', 'scenario2', 'scenario3'][best_scenario_index]
        
        results['recommendation'] = {
            'best_scenario': best_scenario,
            'savings_vs_worst': max(cash_outflows) - min(cash_outflows),
            'effective_loan_rate': inputs['loan_rate'] * (1 - inputs['tax_rate']/100 if inputs['loan_interest_deductible'] else 1),
            'net_investment_return': inputs['investment_return'] * (1 - inputs['tax_rate']/100)
        }
        
        return results

    def create_comparison_table(self, results):
        """Create a clean comparison table"""
        scenarios = ['scenario1', 'scenario2', 'scenario3']
        
        comparison_data = {
            'Financing Strategy': [results[s]['name'] for s in scenarios],
            'Own Capital Used (₹L)': [f"{results[s]['own_capital_used']:.1f}" for s in scenarios],
            'Loan Required (₹L)': [f"{results[s]['loan_amount']:.1f}" for s in scenarios],
            'Capital Invested (₹L)': [f"{results[s]['remaining_capital_invested']:.1f}" for s in scenarios],
            'Monthly EMI (₹)': [f"{results[s]['monthly_emi']:,.0f}" for s in scenarios],
            'Total Interest Cost (₹L)': [f"{results[s]['total_interest_net']:.2f}" for s in scenarios],
            'Investment Returns (₹L)': [f"{results[s]['investment_gain_net']:.2f}" for s in scenarios],
            'NET CASH OUTFLOW (₹L)': [f"{results[s]['total_cash_outflow']:.2f}" for s in scenarios]
        }
        
        return pd.DataFrame(comparison_data)

def main():
    st.set_page_config(page_title="Project Financing Calculator", layout="wide")
    
    st.title("🏗️ Project Financing Calculator")
    st.markdown("""
    **Compare three financing strategies for your project:**
    
    1. **Maximum Own Funding**: Use your capital first, minimize borrowing
    2. **Maximum Leverage**: Borrow the full amount, invest your capital separately  
    3. **Balanced Approach**: Custom mix of own funding and borrowing
    
    The calculator shows the **Net Cash Outflow** - your true out-of-pocket cost after considering loan interest, tax benefits, and investment returns.
    """)
    
    calc = ImprovedProjectFinancingCalculator()
    
    # Input sections
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📊 Project Details")
        project_cost = st.number_input("Project Cost (₹ Lakhs)", min_value=1.0, value=100.0, step=5.0)
        own_capital = st.number_input("Available Own Capital (₹ Lakhs)", min_value=0.0, value=80.0, step=5.0)
        
        st.subheader("🏦 Loan Details")
        loan_rate = st.number_input("Loan Interest Rate (% per annum)", min_value=1.0, value=10.5, step=0.25)
        loan_tenure = st.slider("Loan Tenure (Years)", min_value=1, max_value=20, value=7)
        loan_interest_deductible = st.checkbox("Loan interest is tax deductible", value=True)
        
        st.subheader("⚖️ Balanced Approach")
        custom_own_contribution = st.number_input(
            "Own capital to use directly (₹ Lakhs)", 
            min_value=0.0, 
            max_value=min(project_cost, own_capital),
            value=min(project_cost, own_capital) * 0.5,
            step=5.0
        )
    
    with col2:
        st.subheader("📈 Investment Details")
        investment_type = st.selectbox("Investment Type", list(calc.investment_options.keys()))
        investment_return = st.number_input(
            "Expected Investment Return (% per annum)", 
            min_value=1.0, 
            value=calc.investment_options[investment_type]['default_return'], 
            step=0.25
        )
        tax_rate = st.slider("Your Tax Rate (%)", min_value=0, max_value=45, value=30, step=5)
        
        st.info(f"""
        **{investment_type} Details:**
        - **Risk Level**: {calc.investment_options[investment_type]['risk']}
        - **Liquidity**: {calc.investment_options[investment_type]['liquidity']}
        - **Tax Treatment**: {calc.investment_options[investment_type]['tax_treatment']}
        """)
    
    # Calculate and display results
    if st.button("🔍 Calculate & Compare", type="primary"):
        inputs = {
            'project_cost': project_cost,
            'own_capital': own_capital,
            'loan_rate': loan_rate,
            'loan_tenure': loan_tenure,
            'loan_interest_deductible': loan_interest_deductible,
            'investment_type': investment_type,
            'investment_return': investment_return,
            'tax_rate': tax_rate,
            'custom_own_contribution': custom_own_contribution
        }
        
        results = calc.calculate_scenarios(inputs)
        
        # Display recommendation
        st.success(f"""
        🎯 **Recommended Strategy**: {results[results['recommendation']['best_scenario']]['name']}
        
        💰 **Savings vs Worst Option**: ₹{results['recommendation']['savings_vs_worst']:.2f} Lakhs
        """)
        
        # Display comparison table
        st.subheader("📊 Detailed Comparison")
        comparison_df = calc.create_comparison_table(results)
        st.dataframe(comparison_df, use_container_width=True)
        
        # Key insights
        st.subheader("🔍 Key Insights")
        col_insight1, col_insight2, col_insight3 = st.columns(3)
        
        with col_insight1:
            st.metric("Effective Loan Rate", f"{results['recommendation']['effective_loan_rate']:.2f}%")
        
        with col_insight2:
            st.metric("Net Investment Return", f"{results['recommendation']['net_investment_return']:.2f}%")
        
        with col_insight3:
            rate_spread = results['recommendation']['effective_loan_rate'] - results['recommendation']['net_investment_return']
            st.metric("Rate Spread", f"{rate_spread:.2f}%", 
                     delta="Borrowing costs more" if rate_spread > 0 else "Investing pays more")
        
        # Export options
        st.subheader("📥 Export Results")
        csv = comparison_df.to_csv(index=False)
        st.download_button(
            label="Download Comparison (CSV)",
            data=csv,
            file_name="project_financing_comparison.csv",
            mime="text/csv"
        )

if __name__ == "__main__":
    main()
