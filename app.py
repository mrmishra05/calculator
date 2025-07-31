import math
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import streamlit as st # Import streamlit

# Set a consistent style for matplotlib plots
plt.style.use('seaborn-v0_8')

class ProjectFinancingCalculator:
    def __init__(self):
        self.investment_options = {
            'FD': {
                'default_return': 7.0,
                'liquidity': 'Medium',
                'tax_efficiency': 'Taxed at slab rate',
                'compounding': 'quarterly',
                'notes': 'Safe, insured up to ₹5L'
            },
            'Liquid Funds': {
                'default_return': 6.75,
                'liquidity': 'High (T+1)',
                'tax_efficiency': 'Lower tax if held > 3 years',
                'compounding': 'daily',
                'notes': 'Great for idle cash, corporate treasuries'
            },
            'SGBs': {
                'default_return': 2.5,
                'liquidity': '8-year lock-in',
                'tax_efficiency': 'Tax-free maturity gains',
                'compounding': 'annual',
                'notes': 'Hedge + tax-free if held full term'
            },
            'Arbitrage Fund': {
                'default_return': 7.0,
                'liquidity': 'High (T+1)',
                'tax_efficiency': 'Equity taxation (10% after 1 yr)',
                'compounding': 'daily',
                'notes': 'Good for high net-worth safety seekers'
            },
            'Debt Funds': {
                'default_return': 7.5,
                'liquidity': 'Medium',
                'tax_efficiency': 'Debt tax rules (indexation gone)',
                'compounding': 'daily',
                'notes': 'Slightly better than FD on return'
            }
        }

    def calculate_emi(self, principal, rate, tenure_years):
        """Calculate EMI using the standard formula"""
        if principal == 0:
            return 0

        monthly_rate = rate / (12 * 100)
        months = tenure_years * 12

        if rate == 0:
            return principal / months

        emi = (principal * monthly_rate * math.pow(1 + monthly_rate, months)) / \
              (math.pow(1 + monthly_rate, months) - 1)

        return emi

    def calculate_investment_growth(self, principal, annual_rate, years, compounding='quarterly'):
        """Calculate investment growth with different compounding frequencies"""
        rate = annual_rate / 100

        if compounding == 'daily':
            return principal * math.pow(1 + rate/365, 365 * years)
        elif compounding == 'quarterly':
            return principal * math.pow(1 + rate/4, 4 * years)
        elif compounding == 'monthly':
            return principal * math.pow(1 + rate/12, 12 * years)
        else:  # annual
            return principal * math.pow(1 + rate, years)

    def generate_year_wise_data(self, inputs, results):
        """Generate year-wise breakdown for analysis"""
        data = []

        for year in range(inputs['loan_tenure'] + 1):
            # Option 1 calculations
            # Ensure calculations are done with full amounts, then converted to lakh for display if needed
            option1_loan_amount_actual = results['option1']['loan_amount'] * 100000
            option1_paid = results['option1']['emi'] * 12 * year
            option1_principal_paid = min(option1_paid, option1_loan_amount_actual)
            option1_interest_paid = max(0, option1_paid - option1_principal_paid) / 100000 # Convert to lakh

            # Option 2 calculations
            option2_loan_amount_actual = inputs['project_cost'] * 100000
            option2_paid = results['option2']['emi'] * 12 * year
            option2_principal_paid = min(option2_paid, option2_loan_amount_actual)
            option2_interest_paid = max(0, option2_paid - option2_principal_paid) / 100000 # Convert to lakh

            # Investment value
            if year == 0:
                investment_value = inputs['own_capital']
            else:
                investment_value = self.calculate_investment_growth(
                    inputs['own_capital'],
                    inputs['investment_return'],
                    year,
                    self.investment_options[inputs['investment_type']]['compounding']
                )

            investment_gain = investment_value - inputs['own_capital']
            post_tax_gain = investment_gain * (1 - inputs['tax_rate']/100)

            data.append({
                'Year': year,
                'Option1_Interest_Paid': option1_interest_paid,
                'Option2_Interest_Paid': option2_interest_paid,
                'Investment_Value': investment_value,
                'Investment_Gain': investment_gain,
                'Post_Tax_Gain': post_tax_gain,
                'Option2_Net_Cost': option2_interest_paid - post_tax_gain
            })

        return pd.DataFrame(data)

    def calculate_comparison(self, inputs):
        """Main calculation function"""
        # Option 1: Use own capital + small loan
        option1_loan_amount_lakh = max(0, inputs['project_cost'] - inputs['own_capital'])
        option1_loan_amount_actual = option1_loan_amount_lakh * 100000 # Convert to actual amount for EMI calc
        option1_emi = self.calculate_emi(option1_loan_amount_actual, inputs['loan_rate'], inputs['loan_tenure'])
        option1_total_payment = option1_emi * 12 * inputs['loan_tenure']
        option1_total_interest = option1_total_payment - option1_loan_amount_actual
        option1_net_outflow = (inputs['own_capital'] * 100000) + option1_total_interest # Convert own_capital to actual
        
        # Convert back to lakh for results
        option1_total_interest_lakh = option1_total_interest / 100000
        option1_net_outflow_lakh = option1_net_outflow / 100000

        # Option 2: Invest own capital + take full loan
        option2_loan_amount_lakh = inputs['project_cost']
        option2_loan_amount_actual = option2_loan_amount_lakh * 100000 # Convert to actual amount for EMI calc
        option2_emi = self.calculate_emi(option2_loan_amount_actual, inputs['loan_rate'], inputs['loan_tenure'])
        option2_total_payment = option2_emi * 12 * inputs['loan_tenure']
        option2_total_interest = option2_total_payment - option2_loan_amount_actual

        # Investment calculations
        investment_maturity_value = self.calculate_investment_growth(
            inputs['own_capital'], # Already in lakh
            inputs['investment_return'],
            inputs['loan_tenure'],
            self.investment_options[inputs['investment_type']]['compounding']
        )

        investment_gain = investment_maturity_value - inputs['own_capital']
        post_tax_gain = investment_gain * (1 - inputs['tax_rate']/100)
        option2_net_outflow = (option2_total_interest / 100000) - post_tax_gain # Convert total_interest to lakh

        # Determine recommendation
        recommendation = 'option1' if option1_net_outflow_lakh < option2_net_outflow else 'option2'
        savings = abs(option1_net_outflow_lakh - option2_net_outflow);

        results = {
            'option1': {
                'loan_amount': option1_loan_amount_lakh,
                'emi': option1_emi,
                'total_payment': option1_total_payment,
                'total_interest': option1_total_interest_lakh,
                'net_outflow': option1_net_outflow_lakh,
                'capital_used': inputs['own_capital']
            },
            'option2': {
                'loan_amount': option2_loan_amount_lakh,
                'emi': option2_emi,
                'total_payment': option2_total_payment,
                'total_interest': option2_total_interest / 100000, # Convert to lakh
                'investment_maturity': investment_maturity_value,
                'investment_gain': investment_gain,
                'post_tax_gain': post_tax_gain,
                'net_outflow': option2_net_outflow
            },
            'recommendation': recommendation,
            'savings': savings,
            'interest_spread': inputs['loan_rate'] - inputs['investment_return']
        }

        return results

    def get_recommendation_text(self, results, inputs):
        """Generate recommendation text"""
        interest_spread = results['interest_spread']

        if results['recommendation'] == 'option1':
            if interest_spread > 3:
                return "💡 Use own funds - loan rate is significantly higher than investment returns"
            else:
                return "💡 Use own funds - better capital preservation with lower total cost"
        else:
            if interest_spread < -2:
                return "💡 Use bank loan and invest - your investment returns significantly exceed loan costs"
            else:
                return "💡 Use bank loan and invest - maintains liquidity while generating positive returns"

    def print_detailed_report(self, inputs, results):
        """Print comprehensive analysis report using st.write"""
        st.subheader("📊 Detailed Analysis")
        st.markdown("---")

        # Input Summary
        st.markdown("#### 📋 Input Parameters:")
        st.write(f"**Total Project Cost:** ₹{inputs['project_cost']:.1f} lakh")
        st.write(f"**Own Capital Available:** ₹{inputs['own_capital']:.1f} lakh")
        st.write(f"**Bank Loan Interest Rate:** {inputs['loan_rate']:.2f}% p.a.")
        st.write(f"**Loan Tenure:** {inputs['loan_tenure']} years")
        st.write(f"**Investment Type:** {inputs['investment_type']}")
        st.write(f"**Investment Return:** {inputs['investment_return']:.2f}% p.a.")
        st.write(f"**Tax Rate:** {inputs['tax_rate']:.0f}%")

        investment_details = self.investment_options[inputs['investment_type']]
        st.markdown("##### Investment Details:")
        st.write(f" - **Liquidity:** {investment_details['liquidity']}")
        st.write(f" - **Tax Efficiency:** {investment_details['tax_efficiency']}")
        st.write(f" - **Compounding:** {investment_details['compounding']}")

        st.markdown("---")

        # Option 1 Analysis
        st.markdown(f"#### 💰 OPTION 1: Use Own Capital (₹{inputs['own_capital']:.1f}L) + Small Loan")
        st.write(f"**Loan Amount:** ₹{results['option1']['loan_amount']:.1f} lakh")
        if results['option1']['loan_amount'] > 0:
            st.write(f"**Monthly EMI:** ₹{results['option1']['emi']:,.0f}")
            st.write(f"**Total Payment:** ₹{results['option1']['total_payment']/100000:.2f} lakh")
            st.write(f"**Total Interest:** ₹{results['option1']['total_interest']:.1f} lakh")
        else:
            st.write("No loan required for Option 1.")
        st.write(f"**Capital Used:** ₹{results['option1']['capital_used']:.1f} lakh")
        st.markdown(f"**NET COST:** ₹{results['option1']['net_outflow']:.2f} lakh")

        st.markdown("---")

        # Option 2 Analysis
        st.markdown(f"#### 📈 OPTION 2: Invest Capital + Full Loan (₹{inputs['project_cost']:.1f}L)")
        st.write(f"**Loan Amount:** ₹{results['option2']['loan_amount']:.1f} lakh")
        st.write(f"**Monthly EMI:** ₹{results['option2']['emi']:,.0f}")
        st.write(f"**Total Payment:** ₹{results['option2']['total_payment']/100000:.2f} lakh")
        st.write(f"**Total Interest:** ₹{results['option2']['total_interest']:.1f} lakh")
        st.markdown("##### Investment Analysis:")
        st.write(f" - **Initial Investment:** ₹{inputs['own_capital']:.1f} lakh")
        st.write(f" - **Maturity Value:** ₹{results['option2']['investment_maturity']:.2f} lakh")
        st.write(f" - **Gross Gain:** ₹{results['option2']['investment_gain']:.2f} lakh")
        st.write(f" - **Post-Tax Gain:** ₹{results['option2']['post_tax_gain']:.2f} lakh")
        st.markdown(f"**NET COST:** ₹{results['option2']['net_outflow']:.2f} lakh")

        st.markdown("---")

        # Final Recommendation
        st.markdown("#### ✅ Recommendation:")
        recommendation_text = self.get_recommendation_text(results, inputs)
        st.success(recommendation_text)
        st.write(f"**Potential Savings:** ₹{results['savings']:.2f} lakh")

        # Additional Insights
        st.markdown("#### 🔍 Key Insights:")
        st.write(f"**Interest Rate Spread:** {results['interest_spread']:.2f}% (Loan Rate - Investment Return)")

        if results['recommendation'] == 'option1':
            st.write(" - ✓ Lower total cost")
            st.warning(" - ⚠ Capital gets locked in project")
            st.warning(" - ⚠ Reduced liquidity")
        else:
            st.write(f" - ✓ Maintains full liquidity of ₹{inputs['own_capital']:.1f}L")
            st.write(" - ✓ Emergency funds available")
            st.write(" - ✓ Opportunity for better investments")
            st.warning(" - ⚠ Higher total interest outgo")

        st.markdown("---")

    def create_visualizations(self, inputs, results):
        """Create comprehensive visualizations and display them with st.pyplot"""
        st.subheader("📈 Visual Analysis")

        df = self.generate_year_wise_data(inputs, results)

        fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(15, 12))

        # 1. Net Cost Comparison
        options = ['Option 1\n(Own Capital)', 'Option 2\n(Invest + Loan)']
        costs = [results['option1']['net_outflow'], results['option2']['net_outflow']]
        colors = ['#2E8B57' if results['recommendation'] == 'option1' else '#FF6B6B',
                  '#2E8B57' if results['recommendation'] == 'option2' else '#FF6B6B']

        bars1 = ax1.bar(options, costs, color=colors, alpha=0.7, edgecolor='black')
        ax1.set_title('Net Cost Comparison', fontsize=14, fontweight='bold')
        ax1.set_ylabel('Net Cost (₹ lakh)')

        for bar, cost in zip(bars1, costs):
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width()/2., height + 0.5,
                     f'₹{cost:.1f}L', ha='center', va='bottom', fontweight='bold')

        # 2. EMI Comparison
        emis = [results['option1']['emi'], results['option2']['emi']]
        bars2 = ax2.bar(options, emis, color=['#4CAF50', '#FF9800'], alpha=0.7, edgecolor='black')
        ax2.set_title('Monthly EMI Comparison', fontsize=14, fontweight='bold')
        ax2.set_ylabel('Monthly EMI (₹)')

        for bar, emi in zip(bars2, emis):
            height = bar.get_height()
            ax2.text(bar.get_x() + bar.get_width()/2., height + 5000,
                     f'₹{emi:,.0f}', ha='center', va='bottom', fontweight='bold')

        # 3. Year-wise Analysis
        ax3.plot(df['Year'], df['Option1_Interest_Paid'], marker='o',
                 linewidth=2, label='Option 1 Interest', color='#2196F3')
        ax3.plot(df['Year'], df['Option2_Interest_Paid'], marker='s',
                 linewidth=2, label='Option 2 Interest', color='#F44336')
        ax3.plot(df['Year'], df['Investment_Value'], marker='^',
                 linewidth=2, label='Investment Value', color='#4CAF50')

        ax3.set_title('Cost Evolution Over Time', fontsize=14, fontweight='bold')
        ax3.set_xlabel('Year')
        ax3.set_ylabel('Amount (₹ lakh)')
        ax3.legend()
        ax3.grid(True, alpha=0.3)

        # 4. Break-even Analysis
        net_costs_option2 = df['Option2_Net_Cost'].values
        option1_cost = results['option1']['net_outflow']

        ax4.plot(df['Year'], [option1_cost] * len(df),
                 linewidth=3, label='Option 1 Net Cost', color='#2196F3', linestyle='--')
        ax4.plot(df['Year'], net_costs_option2, marker='o',
                 linewidth=2, label='Option 2 Net Cost', color='#F44336')

        ax4.set_title('Break-even Analysis', fontsize=14, fontweight='bold')
        ax4.set_xlabel('Year')
        ax4.set_ylabel('Net Cost (₹ lakh)')
        ax4.legend()
        ax4.grid(True, alpha=0.3)
        ax4.axhline(y=0, color='black', linestyle='-', alpha=0.3)

        plt.tight_layout()
        st.pyplot(fig) # Display the figure in Streamlit
        plt.close(fig) # Close the figure to free memory

        return df

    def export_to_excel(self, inputs, results, df, filename='project_financing_analysis.xlsx'):
        """Export detailed analysis to Excel and provide a download button"""
        # Create an in-memory Excel file
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Summary sheet
            summary_data = {
                'Parameter': [
                    'Project Cost (₹ lakh)', 'Own Capital (₹ lakh)', 'Loan Rate (%)',
                    'Tenure (years)', 'Investment Type', 'Investment Return (%)', 'Tax Rate (%)',
                    '', 'OPTION 1 - Use Own Capital', 'Loan Amount (₹ lakh)', 'Monthly EMI (₹)',
                    'Total Interest (₹ lakh)', 'Net Cost (₹ lakh)', '', 'OPTION 2 - Invest + Loan',
                    'Loan Amount (₹ lakh)', 'Monthly EMI (₹)', 'Total Interest (₹ lakh)',
                    'Investment Maturity (₹ lakh)', 'Post-tax Gain (₹ lakh)', 'Net Cost (₹ lakh)',
                    '', 'RECOMMENDATION', 'Better Option', 'Savings (₹ lakh)'
                ],
                'Value': [
                    inputs['project_cost'], inputs['own_capital'], inputs['loan_rate'],
                    inputs['loan_tenure'], inputs['investment_type'], inputs['investment_return'],
                    inputs['tax_rate'], '', '', results['option1']['loan_amount'],
                    f"₹{results['option1']['emi']:,.0f}", results['option1']['total_interest'],
                    results['option1']['net_outflow'], '', '', results['option2']['loan_amount'],
                    f"₹{results['option2']['emi']:,.0f}", results['option2']['total_interest'],
                    results['option2']['investment_maturity'], results['option2']['post_tax_gain'],
                    results['option2']['net_outflow'], '', '',
                    'Option 1' if results['recommendation'] == 'option1' else 'Option 2',
                    results['savings']
                ]
            }

            pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)

            # Year-wise analysis
            df.to_excel(writer, sheet_name='Year_wise_Analysis', index=False)

            # Investment options reference
            inv_df = pd.DataFrame.from_dict(self.investment_options, orient='index')
            inv_df.to_excel(writer, sheet_name='Investment_Options')

        output.seek(0) # Rewind the buffer to the beginning

        st.download_button(
            label="Download Analysis as Excel",
            data=output,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Click to download the detailed financial analysis in Excel format."
        )


def main():
    """Main function to run the calculator with Streamlit UI"""
    st.set_page_config(layout="wide", page_title="Project Financing Calculator")
    st.title("🏦 Project Financing Calculator")
    st.markdown("""
        This tool helps you compare two project financing strategies:
        1.  **Option 1:** Use your own capital first, then take a loan for the remaining project cost.
        2.  **Option 2:** Invest your own capital and take a full loan for the entire project cost.
        
        Enter your project details below to see a detailed comparison and recommendation.
    """)

    calc = ProjectFinancingCalculator()

    # Create columns for input layout
    col1, col2 = st.columns(2)

    with col1:
        st.header("Project Details")
        project_cost = st.number_input(
            "Enter total project cost (₹ lakh):",
            min_value=0.1, value=150.0, step=10.0, format="%.1f",
            help="Total cost of the project in Lakhs (e.g., 150 for ₹1.5 Crore)"
        )
        own_capital = st.number_input(
            "Enter own capital available (₹ lakh):",
            min_value=0.0, value=100.0, step=10.0, format="%.1f",
            help="Your available capital in Lakhs to fund the project"
        )
        loan_rate = st.number_input(
            "Enter bank loan interest rate (% p.a.):",
            min_value=0.1, value=10.5, step=0.1, format="%.2f",
            help="Annual interest rate for the bank loan"
        )
        loan_tenure = st.slider(
            "Select loan tenure (years):",
            min_value=1, max_value=30, value=7, step=1,
            help="Duration of the loan in years"
        )

    with col2:
        st.header("Investment Details")
        investment_type = st.selectbox(
            "Select investment type for your capital:",
            options=list(calc.investment_options.keys()),
            index=0, # Default to FD
            help="Choose where you'd invest your own capital if you take a full loan"
        )
        # Get default return for selected investment type
        default_inv_return = calc.investment_options[investment_type]['default_return']
        investment_return = st.number_input(
            "Enter expected investment return (% p.a.):",
            min_value=0.0, value=default_inv_return, step=0.1, format="%.2f",
            help="Expected annual return from your chosen investment type"
        )
        tax_rate = st.slider(
            "Select your tax rate (%):",
            min_value=0, max_value=50, value=30, step=5,
            help="Your income tax slab rate relevant for investment gains"
        )
        st.markdown(f"**Selected Investment Notes:** {calc.investment_options[investment_type]['notes']}")
        st.markdown(f"**Selected Investment Liquidity:** {calc.investment_options[investment_type]['liquidity']}")
        st.markdown(f"**Selected Investment Tax Efficiency:** {calc.investment_options[investment_type]['tax_efficiency']}")
        st.markdown(f"**Selected Investment Compounding:** {calc.investment_options[investment_type]['compounding']}")


    # Store inputs in a dictionary
    inputs = {
        'project_cost': float(project_cost),
        'own_capital': float(own_capital),
        'loan_rate': float(loan_rate),
        'loan_tenure': int(loan_tenure),
        'investment_type': investment_type,
        'investment_return': float(investment_return),
        'tax_rate': float(tax_rate)
    }

    st.markdown("---")

    # Add a button to trigger calculation
    if st.button("Calculate Financing Options", type="primary"):
        # The problematic validation that caused KeyError has been removed.
        # The underlying calculation logic already handles cases where own_capital >= project_cost
        # by setting option1_loan_amount to 0.

        # Perform calculations
        results = calc.calculate_comparison(inputs)

        # Display detailed report
        calc.print_detailed_report(inputs, results)

        # Create and display visualizations
        df = calc.create_visualizations(inputs, results)

        # Provide Excel export
        import io # Import io here as it's only needed for Excel export
        calc.export_to_excel(inputs, results, df)
