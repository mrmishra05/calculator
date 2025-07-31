import math
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import streamlit as st # Import streamlit
import io # Needed for in-memory Excel export

# For PDF generation
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.lib import colors

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
        
        if monthly_rate == 0:
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
        # Define column names explicitly for DataFrame creation
        column_names = [
            'Year',
            'Option1_Interest_Paid',
            'Option2_Interest_Paid',
            'Option3_Interest_Paid',
            'Investment_Value_Option2',
            'Investment_Value_Option3',
            'Investment_Gain_Option2',
            'Investment_Gain_Option3',
            'Post_Tax_Gain_Option2',
            'Post_Tax_Gain_Option3',
            'Option1_Net_Cost_Cumulative',
            'Option2_Net_Cost_Cumulative',
            'Option3_Net_Cost_Cumulative'
        ]
        data = []

        for year in range(inputs['loan_tenure'] + 1):
            # Option 1 calculations (cumulative)
            option1_loan_amount_actual_at_start = results['option1']['loan_amount'] * 100000
            option1_paid_cumulative = results['option1']['emi'] * 12 * year
            option1_principal_paid_cumulative = min(option1_paid_cumulative, option1_loan_amount_actual_at_start)
            option1_interest_paid_cumulative = max(0, option1_paid_cumulative - option1_principal_paid_cumulative) / 100000 # Convert to lakh
            
            # Recalculate Option1_Net_Cost_Cumulative based on new definition
            # Net Cost = Project Cost + Effective Total Interest - Post-Tax Investment Gains (0 for Option 1)
            option1_effective_interest_cumulative = option1_interest_paid_cumulative
            if inputs['loan_interest_deductible']:
                option1_effective_interest_cumulative *= (1 - inputs['tax_rate']/100)
            option1_net_cost_cumulative = inputs['project_cost'] + option1_effective_interest_cumulative


            # Option 2 calculations (cumulative)
            option2_loan_amount_actual_at_start = inputs['project_cost'] * 100000
            option2_paid_cumulative = results['option2']['emi'] * 12 * year
            option2_principal_paid_cumulative = min(option2_paid_cumulative, option2_loan_amount_actual_at_start)
            option2_interest_paid_cumulative = max(0, option2_paid_cumulative - option2_principal_paid_cumulative) / 100000 # Convert to lakh

            option2_effective_interest_cumulative = option2_interest_paid_cumulative
            if inputs['loan_interest_deductible']:
                option2_effective_interest_cumulative *= (1 - inputs['tax_rate']/100)

            investment_value_option2 = inputs['own_capital']
            if year > 0:
                investment_value_option2 = self.calculate_investment_growth(
                    inputs['own_capital'],
                    inputs['investment_return'],
                    year,
                    self.investment_options[inputs['investment_type']]['compounding']
                )
            investment_gain_option2 = investment_value_option2 - inputs['own_capital']
            post_tax_gain_option2 = investment_gain_option2 * (1 - inputs['tax_rate']/100)
            
            # Recalculate Option2_Net_Cost_Cumulative based on new definition
            option2_net_cost_cumulative = inputs['project_cost'] + option2_effective_interest_cumulative - post_tax_gain_option2


            # Option 3 calculations (cumulative)
            option3_capital_used = inputs['custom_capital_contribution']
            option3_loan_amount_actual_at_start = max(0, inputs['project_cost'] - option3_capital_used) * 100000
            
            option3_paid_cumulative = results['option3']['emi'] * 12 * year if year > 0 else 0
            option3_principal_paid_cumulative = min(option3_paid_cumulative, option3_loan_amount_actual_at_start)
            option3_interest_paid_cumulative = max(0, option3_paid_cumulative - option3_principal_paid_cumulative) / 100000 # Convert to lakh

            option3_effective_interest_cumulative = option3_interest_paid_cumulative
            if inputs['loan_interest_deductible']:
                option3_effective_interest_cumulative *= (1 - inputs['tax_rate']/100)

            option3_remaining_own_capital = inputs['own_capital'] - option3_capital_used
            investment_value_option3 = option3_remaining_own_capital # At year 0
            if year > 0 and option3_remaining_own_capital > 0:
                investment_value_option3 = self.calculate_investment_growth(
                    option3_remaining_own_capital,
                    inputs['investment_return'],
                    year,
                    self.investment_options[inputs['investment_type']]['compounding']
                )
            investment_gain_option3 = investment_value_option3 - option3_remaining_own_capital
            post_tax_gain_option3 = investment_gain_option3 * (1 - inputs['tax_rate']/100)
            
            # Recalculate Option3_Net_Cost_Cumulative based on new definition
            option3_net_cost_cumulative = inputs['project_cost'] + option3_effective_interest_cumulative - post_tax_gain_option3


            data.append([
                year,
                option1_interest_paid_cumulative,
                option2_interest_paid_cumulative,
                option3_interest_paid_cumulative,
                investment_value_option2,
                investment_value_option3,
                investment_gain_option2,
                investment_gain_option3,
                post_tax_gain_option2,
                post_tax_gain_option3,
                option1_net_cost_cumulative,
                option2_net_cost_cumulative,
                option3_net_cost_cumulative
            ])

        # Create DataFrame with explicit column names
        return pd.DataFrame(data, columns=column_names)

    def calculate_comparison(self, inputs):
        """Main calculation function"""
        # Option 1: Use own capital + small loan
        option1_loan_amount_lakh = max(0, inputs['project_cost'] - inputs['own_capital'])
        option1_loan_amount_actual = option1_loan_amount_lakh * 100000 # Convert to actual amount for EMI calc
        option1_emi = self.calculate_emi(option1_loan_amount_actual, inputs['loan_rate'], inputs['loan_tenure'])
        option1_total_payment = option1_emi * 12 * inputs['loan_tenure']
        option1_total_interest = option1_total_payment - option1_loan_amount_actual
        
        # Apply tax deductibility for loan interest if applicable
        option1_effective_interest = option1_total_interest
        if inputs['loan_interest_deductible']:
            option1_effective_interest *= (1 - inputs['tax_rate']/100)
        
        option1_total_interest_lakh = option1_total_interest / 100000
        option1_effective_interest_lakh = option1_effective_interest / 100000

        # NEW NET COST DEFINITION: Project Cost + Effective Total Interest - Post-Tax Investment Gains (0 for Option 1)
        option1_net_outflow = inputs['project_cost'] + option1_effective_interest_lakh
        
        # NEW: Total Investment (Gross Cash Outflow)
        option1_total_investment_gross = (option1_total_payment / 100000) + inputs['own_capital']


        # Option 2: Invest own capital + take full loan
        option2_loan_amount_lakh = inputs['project_cost']
        option2_loan_amount_actual = option2_loan_amount_lakh * 100000 # Convert to actual amount for EMI calc
        option2_emi = self.calculate_emi(option2_loan_amount_actual, inputs['loan_rate'], inputs['loan_tenure'])
        option2_total_payment = option2_emi * 12 * inputs['loan_tenure']
        option2_total_interest = option2_total_payment - option2_loan_amount_actual
        
        # Apply tax deductibility for loan interest if applicable
        option2_effective_interest = option2_total_interest
        if inputs['loan_interest_deductible']:
            option2_effective_interest *= (1 - inputs['tax_rate']/100)
        
        # Investment calculations for Option 2
        investment_maturity_value_option2 = self.calculate_investment_growth(
            inputs['own_capital'], # Already in lakh
            inputs['investment_return'],
            inputs['loan_tenure'],
            self.investment_options[inputs['investment_type']]['compounding']
        )
        investment_gain_option2 = investment_maturity_value_option2 - inputs['own_capital']
        post_tax_gain_option2 = investment_gain_option2 * (1 - inputs['tax_rate']/100)

        # NEW NET COST DEFINITION: Project Cost + Effective Total Interest - Post-Tax Investment Gains
        option2_net_outflow = inputs['project_cost'] + (option2_effective_interest / 100000) - post_tax_gain_option2

        # NEW: Total Investment (Gross Cash Outflow)
        option2_total_investment_gross = (option2_total_payment / 100000) + inputs['own_capital']


        # Option 3: Custom Capital Contribution + Loan
        option3_capital_used = inputs['custom_capital_contribution']
        option3_loan_amount_lakh = max(0, inputs['project_cost'] - option3_capital_used)
        option3_loan_amount_actual = option3_loan_amount_lakh * 100000

        option3_emi = self.calculate_emi(option3_loan_amount_actual, inputs['loan_rate'], inputs['loan_tenure'])
        option3_total_payment = option3_emi * 12 * inputs['loan_tenure']
        option3_total_interest = option3_total_payment - option3_loan_amount_actual
        
        # Apply tax deductibility for loan interest if applicable
        option3_effective_interest = option3_total_interest
        if inputs['loan_interest_deductible']:
            option3_effective_interest *= (1 - inputs['tax_rate']/100)
        
        option3_total_interest_lakh = option3_total_interest / 100000
        option3_effective_interest_lakh = option3_effective_interest / 100000


        # Calculate investment for remaining own capital (if any) for Option 3
        option3_remaining_own_capital = inputs['own_capital'] - option3_capital_used
        option3_investment_maturity_value = 0
        option3_investment_gain = 0
        option3_post_tax_gain = 0

        if option3_remaining_own_capital > 0:
            option3_investment_maturity_value = self.calculate_investment_growth(
                option3_remaining_own_capital,
                inputs['investment_return'],
                inputs['loan_tenure'],
                self.investment_options[inputs['investment_type']]['compounding']
            )
            option3_investment_gain = option3_investment_maturity_value - option3_remaining_own_capital
            option3_post_tax_gain = option3_investment_gain * (1 - inputs['tax_rate']/100)

        # NEW NET COST DEFINITION: Project Cost + Effective Total Interest - Post-Tax Investment Gains
        option3_net_outflow = inputs['project_cost'] + option3_effective_interest_lakh - option3_post_tax_gain

        # NEW: Total Investment (Gross Cash Outflow)
        option3_total_investment_gross = (option3_total_payment / 100000) + inputs['own_capital']


        # Calculate annualized effective rates for display
        # Effective Loan Rate (Annualized)
        effective_loan_rate_annual = inputs['loan_rate']
        if inputs['loan_interest_deductible']:
            effective_loan_rate_annual = inputs['loan_rate'] * (1 - inputs['tax_rate']/100)

        # Effective Investment Return (Annualized Post-Tax CAGR)
        effective_investment_return_annual = 0
        if inputs['own_capital'] > 0 and inputs['loan_tenure'] > 0:
            # Calculate CAGR for Option 2's investment (using post-tax maturity value)
            # Ensure investment_maturity_value_option2 is > 0 to avoid log(0)
            if investment_maturity_value_option2 > 0:
                effective_investment_return_annual = ( (investment_maturity_value_option2 / inputs['own_capital'])**(1/inputs['loan_tenure']) - 1 ) * 100


        # Determine recommendation (now comparing 3 options)
        all_net_outflows = {
            'option1': option1_net_outflow,
            'option2': option2_net_outflow,
            'option3': option3_net_outflow
        }
        
        min_net_outflow_value = min(all_net_outflows.values())
        
        recommendation = ''
        for key, value in all_net_outflows.items():
            if value == min_net_outflow_value:
                recommendation = key
                break # Found the first matching option

        max_net_outflow_value = max(all_net_outflows.values())
        savings_against_worst = max_net_outflow_value - min_net_outflow_value

        # Calculate prepayment penalty
        prepayment_cost = 0
        if inputs['prepayment_penalty_pct'] > 0 and inputs['loan_tenure'] > 0:
            # For simplicity, calculate penalty on initial loan amount for Option 2 (full loan)
            # A more complex model would calculate it on outstanding principal at prepayment time
            prepayment_cost = (inputs['project_cost'] * inputs['prepayment_penalty_pct'] / 100)


        # Update the results dictionary
        results = {
            'option1': {
                'loan_amount': option1_loan_amount_lakh,
                'emi': option1_emi,
                'total_payment': option1_total_payment,
                'total_interest': option1_total_interest_lakh,
                'effective_interest': option1_effective_interest_lakh,
                'net_outflow': option1_net_outflow,
                'capital_used': inputs['own_capital'], # This is capital used directly for project
                'total_investment_gross': option1_total_investment_gross # NEW
            },
            'option2': {
                'loan_amount': option2_loan_amount_lakh,
                'emi': option2_emi,
                'total_payment': option2_total_payment,
                'total_interest': option2_total_interest / 100000,
                'effective_interest': option2_effective_interest / 100000,
                'investment_maturity': investment_maturity_value_option2,
                'investment_gain': investment_gain_option2,
                'post_tax_gain': post_tax_gain_option2,
                'net_outflow': option2_net_outflow,
                'capital_used': 0.0, # No own capital used directly for project
                'capital_invested': inputs['own_capital'], # All own capital is invested
                'total_investment_gross': option2_total_investment_gross # NEW
            },
            'option3': {
                'capital_used': option3_capital_used, # Capital used directly for project
                'loan_amount': option3_loan_amount_lakh,
                'emi': option3_emi,
                'total_payment': option3_total_payment,
                'total_interest': option3_total_interest_lakh,
                'effective_interest': option3_effective_interest_lakh,
                'remaining_own_capital_invested': option3_remaining_own_capital, # Capital invested
                'investment_maturity': option3_investment_maturity_value,
                'investment_gain': option3_investment_gain,
                'post_tax_gain': option3_post_tax_gain,
                'net_outflow': option3_net_outflow,
                'total_investment_gross': option3_total_investment_gross # NEW
            },
            'recommendation': recommendation,
            'savings': savings_against_worst,
            'interest_spread': inputs['loan_rate'] - inputs['investment_return'],
            'effective_loan_rate_annual': effective_loan_rate_annual,
            'effective_investment_return_annual': effective_investment_return_annual,
            'prepayment_cost': prepayment_cost
        }
        
        return results

    def get_recommendation_text(self, results, inputs):
        """Generate recommendation text considering 3 options"""
        interest_spread = results['interest_spread']
        recommendation_option = results['recommendation']

        base_text = ""
        if recommendation_option == 'option1':
            base_text = "💡 Option 1 (Use own funds directly) is recommended."
            if interest_spread > 3:
                base_text += " Loan rate is significantly higher than investment returns."
            else:
                base_text += " This offers better capital preservation with lower total cost."
        elif recommendation_option == 'option2':
            base_text = "💡 Option 2 (Use bank loan and invest) is recommended."
            if interest_spread < -2:
                base_text += " Your investment returns significantly exceed loan costs."
            else:
                base_text += " This maintains liquidity while generating positive returns."
        else: # recommendation_option == 'option3'
            base_text = "💡 Option 3 (Custom Contribution) is recommended."
            if results['option3']['remaining_own_capital_invested'] > 0 and results['option3']['loan_amount'] > 0:
                base_text += f" It's a balanced approach using ₹{results['option3']['capital_used']:.1f}L directly and investing ₹{results['option3']['remaining_own_capital_invested']:.1f}L."
            elif results['option3']['loan_amount'] == 0 and results['option3']['remaining_own_capital_invested'] == 0:
                base_text += f" You can fully fund the project with ₹{results['option3']['capital_used']:.1f}L directly, with no loan needed and no remaining capital to invest."
            elif results['option3']['loan_amount'] == 0 and results['option3']['remaining_own_capital_invested'] > 0:
                base_text += f" You fully fund the project directly and invest the remaining ₹{results['option3']['remaining_own_capital_invested']:.1f}L of your capital."
            else:
                base_text += " It provides the lowest net cost for your customized approach."
        
        if inputs['loan_interest_deductible']:
            base_text += f" (Note: Loan interest is considered tax-deductible, significantly reducing its effective cost.)"
        
        return base_text


    def print_detailed_report(self, inputs, results):
        """Print comprehensive analysis report using st.write"""
        st.subheader("📊 Detailed Analysis")
        st.markdown("---")

        # Recommendation at the top
        st.markdown("#### ✅ Recommendation:")
        recommendation_text = self.get_recommendation_text(results, inputs)
        st.success(recommendation_text)
        st.write(f"**Potential Savings (compared to worst option):** ₹{results['savings']:.2f} lakh")
        st.markdown("---")

        # Tabular Comparison of Options
        st.markdown("#### 📈 Financing Options Comparison")
        comparison_data = {
            'Metric': [
                'Capital Used Directly for Project',
                'Own Capital Invested',
                'Loan Amount',
                'Monthly EMI',
                'Gross Total Interest',
                'Effective Total Interest (After Tax)',
                'Investment Maturity Value (from invested capital)',
                'Gross Investment Gain (from invested capital)',
                'Post-Tax Investment Gain (from invested capital)',
                '**Total Investment (Gross Cash Outflow)**', # NEW COLUMN
                '**NET COST**'
            ],
            'Option 1: Own Capital + Small Loan': [
                f"₹{results['option1']['capital_used']:.1f} lakh",
                "₹0.0 lakh",
                f"₹{results['option1']['loan_amount']:.1f} lakh",
                f"₹{results['option1']['emi']:,.0f}",
                f"₹{results['option1']['total_interest']:.1f} lakh",
                f"₹{results['option1']['effective_interest']:.1f} lakh",
                "₹0.0 lakh",
                "₹0.0 lakh",
                "₹0.0 lakh",
                f"₹{results['option1']['total_investment_gross']:.2f} lakh", # NEW VALUE
                f"₹{results['option1']['net_outflow']:.2f} lakh"
            ],
            'Option 2: Invest Capital + Full Loan': [
                "₹0.0 lakh",
                f"₹{results['option2']['capital_invested']:.1f} lakh",
                f"₹{results['option2']['loan_amount']:.1f} lakh",
                f"₹{results['option2']['emi']:,.0f}",
                f"₹{results['option2']['total_interest']:.1f} lakh",
                f"₹{results['option2']['effective_interest']:.1f} lakh",
                f"₹{results['option2']['investment_maturity']:.2f} lakh",
                f"₹{results['option2']['investment_gain']:.2f} lakh",
                f"₹{results['option2']['post_tax_gain']:.2f} lakh",
                f"₹{results['option2']['total_investment_gross']:.2f} lakh", # NEW VALUE
                f"₹{results['option2']['net_outflow']:.2f} lakh"
            ],
            'Option 3: Custom Capital + Loan': [
                f"₹{results['option3']['capital_used']:.1f} lakh",
                f"₹{results['option3']['remaining_own_capital_invested']:.1f} lakh",
                f"₹{results['option3']['loan_amount']:.1f} lakh",
                f"₹{results['option3']['emi']:,.0f}",
                f"₹{results['option3']['total_interest']:.1f} lakh",
                f"₹{results['option3']['effective_interest']:.1f} lakh",
                f"₹{results['option3']['investment_maturity']:.2f} lakh" if results['option3']['remaining_own_capital_invested'] > 0 else "₹0.0 lakh",
                f"₹{results['option3']['investment_gain']:.2f} lakh" if results['option3']['remaining_own_capital_invested'] > 0 else "₹0.0 lakh",
                f"₹{results['option3']['post_tax_gain']:.2f} lakh" if results['option3']['remaining_own_capital_invested'] > 0 else "₹0.0 lakh",
                f"₹{results['option3']['total_investment_gross']:.2f} lakh", # NEW VALUE
                f"₹{results['option3']['net_outflow']:.2f} lakh"
            ]
        }
        comparison_df = pd.DataFrame(comparison_data).set_index('Metric')
        st.dataframe(comparison_df)

        # Net Cost Definition Expander
        with st.expander("❓ What are 'Net Cost' and 'Total Investment' in this context?"): # Updated title
            st.markdown("""
            **"Net Cost"** represents your ultimate **out-of-pocket expense to fund the project** over the loan tenure.
            It is calculated as:
            `Net Cost = Total Project Cost + Effective Total Loan Interest (After Tax) - Post-Tax Investment Gains`
            
            Let's break that down:
            * **Total Project Cost:** The fundamental cost of acquiring or building the project.
            * **Effective Total Loan Interest (After Tax):** This is the gross interest paid on the loan, reduced by any tax savings if the loan interest is tax-deductible as a business expense (based on your input `Tax Rate`).
            * **Post-Tax Investment Gains:** Any net profits you earn from investing your own capital (after taxes) are subtracted, as these gains effectively reduce your overall cost.
            
            This metric provides a true bottom-line figure for each financing option, allowing for direct and comparable evaluation of what you *truly spent* to get the project done.
            
            ---
            
            **"Total Investment (Gross Cash Outflow)"** represents the **total cash flow involved** in financing the project and managing your available capital over the loan tenure.
            It is calculated as:
            `Total Investment = Total Loan Payment (Principal + Gross Interest) + Own Capital Available`
            
            This figure is typically much higher than "Net Cost" because it does not account for tax benefits on interest or offsetting investment gains. It gives a sense of the gross funds that move through your financing strategy.
            """)
        st.markdown("---")

        # Input Summary
        st.markdown("#### 📋 Input Parameters Used:")
        st.write(f"**Total Project Cost:** ₹{inputs['project_cost']:.1f} lakh")
        st.write(f"**Own Capital Available:** ₹{inputs['own_capital']:.1f} lakh")
        st.write(f"**Bank Loan Interest Rate:** {inputs['loan_rate']:.2f}% p.a. ({inputs['loan_type']})")
        st.write(f"**Loan Tenure:** {inputs['loan_tenure']} years")
        st.write(f"**Loan Interest Tax Deductible:** {'Yes' if inputs['loan_interest_deductible'] else 'No'}")
        st.write(f"**Prepayment Penalty:** {inputs['prepayment_penalty_pct']:.2f}%")
        st.write(f"**Minimum Liquidity Target:** ₹{inputs['min_liquidity_target']:.1f} lakh")
        st.write(f"**Investment Type:** {inputs['investment_type']}")
        st.write(f"**Investment Return:** {inputs['investment_return']:.2f}% p.a.")
        st.write(f"**Tax Rate:** {inputs['tax_rate']:.0f}%")
        
        # Display custom capital contribution based on input type
        if inputs['custom_capital_input_type'] == 'Value':
            st.write(f"**Custom Capital Contribution (Option 3):** ₹{inputs['custom_capital_contribution']:.1f} lakh")
        else:
            st.write(f"**Custom Capital Contribution (Option 3):** {inputs['custom_capital_percentage']:.1f}% of Own Capital (₹{inputs['custom_capital_contribution']:.1f} lakh)")


        investment_details = self.investment_options[inputs['investment_type']]
        st.markdown("##### Investment Details:")
        st.write(f" - **Liquidity:** {investment_details['liquidity']}")
        st.write(f" - **Tax Efficiency:** {investment_details['tax_efficiency']}")
        st.write(f" - **Compounding:** {investment_details['compounding']}")
        st.markdown("---")

        # Key Insights & Strategic Considerations - Improved Representation
        st.markdown("#### 🔍 Key Insights & Strategic Considerations:")
        
        col_metrics_1, col_metrics_2, col_metrics_3 = st.columns(3)
        with col_metrics_1:
            st.metric(label="Effective Loan Interest Rate (After Tax) (Annualized)", value=f"{results['effective_loan_rate_annual']:.2f}% p.a.", delta_color="off")
        with col_metrics_2:
            st.metric(label="Effective Investment Return (After Tax) (Annualized)", value=f"{results['effective_investment_return_annual']:.2f}% p.a.", delta_color="off")
        with col_metrics_3:
            st.metric(label="Interest Rate Spread (Loan - Investment)", value=f"{results['interest_spread']:.2f}%", delta_color="off")
        
        st.markdown("""
        These annualized rates provide a clearer picture of the true cost of borrowing and the actual return on your investments, factoring in tax benefits.
        """)

        with st.expander("Risk & Scenario Analysis"):
            st.write(f" - **Loan Interest Rate Risk:** The loan is **{inputs['loan_type']}**. If floating, consider sensitivity to rate hikes and build in buffers.")
            st.write(f" - **Prepayment Cost:** A {inputs['prepayment_penalty_pct']:.2f}% penalty on a full loan (Option 2) would be **₹{results['prepayment_cost']:.2f} lakh**. Factor this into early exit scenarios and loan terms.")
            st.write(" - **Investment Volatility:** Assumed investment returns are estimates. Stress test your plan with $\pm 1-2\%$ returns to understand the impact on net cost and liquidity.")
            st.write(f" - **Liquidity Buffer:** Your target minimum liquidity is **₹{inputs['min_liquidity_target']:.1f} lakh**. Ensure the chosen option consistently maintains this, especially under stressed cash flow scenarios (e.g., delayed project revenue, unexpected expenses).")
        
        with st.expander("Nuanced Strategic Considerations"):
            if inputs['loan_interest_deductible']:
                st.write(" - **Tax Planning:** Loan interest is considered tax-deductible, significantly reducing the effective cost of borrowing. Ensure proper documentation for claiming this deduction.")
            else:
                st.write(" - **Tax Planning:** Loan interest is NOT considered tax-deductible, meaning the gross interest is the effective cost. Explore other tax optimization strategies.")
            st.write(" - **Credit Profile & Leverage:** Assess the impact of increased debt on your company's debt-to-equity ratio, credit rating, and future borrowing capacity. Excessive leverage can affect banking relationships and future funding flexibility.")
            st.write(" - **Regulatory Compliance:** Confirm all documentation (project reports, audited financials) are ready for smooth loan disbursal and subsequent annual reviews to avoid penalties or delays.")
            st.write(" - **Investment Liquidity/Market Risk:** Even 'liquid' investments carry some market risk (e.g., temporary impairment during market freezes). Maintain an emergency buffer in a highly liquid bank account (beyond investment) for immediate needs.")
            st.write(" - **Optimization Moves & Additional Value Levers:** Consider phased loan drawdown or capital deployment as per project schedule to optimize interest outflow and minimize idle funds.")
            st.write(" - **Step-up EMI/SWP Option:** If expecting higher business income post-project commissioning, opt for step-up EMI structure. Alternatively, consider Systematic Withdrawal Plan (SWP) from investments to sync cash flows with project needs.")
            st.write(" - **Insurance:** Insure the principal loan with adequate keyman/credit insurance, especially if individual guarantees are required, to mitigate personal risk.")
            st.write(" - **Renegotiation Flexibility:** Build in terms for possible loan prepayment, refinancing, or top-up without penal costs in your loan agreements.")
            st.write(" - **AR/AP Synchronization:** Consider peer business working capital requirements—using part of the liquid reserves to cover temporary spikes in Accounts Receivable (AR) / Accounts Payable (AP), thus reducing the need for additional short-term borrowings.")

        st.markdown("---")
        st.markdown("#### 🚀 Strategic Summary: Best Practices for CFOs")
        st.write(f" - **Quantify & Monitor Liquidity:** Always maintain a minimum liquidity threshold (e.g., **₹{inputs['min_liquidity_target']:.1f} lakh** always in instantly-accessible form) to ensure business agility and emergency preparedness.")
        st.write(" - **Dynamic Financial Planning:** Be prepared for faster loan prepayment or capital recycling if the business environment or investment returns shift in future years. Agility is key.")
        st.write(" - **Continuous Benchmarking:** Regularly benchmark your investment returns and borrowing costs against evolving market rates and new reinvestment opportunities to ensure optimal capital allocation.")
        st.write(" - **Dashboard Monitoring:** Implement and maintain a dynamic dashboard updated for promoters, CFO, and the board on all key financial parameters and risk indicators.")

        st.markdown("---")
        st.markdown("### **In short:**")
        st.info("""
        This section provides a concise, executive summary of the holistic advisory approach. It distills the comprehensive analysis into key actionable takeaways, ensuring that even busy stakeholders can quickly grasp the core value proposition: transforming "cost minimization" into **value optimization** by considering cash flow, risk, flexibility, regulatory, tax, and capital structure.
        """)
        st.markdown("---")

    def create_visualizations(self, inputs, results):
        """Create comprehensive visualizations and display them with st.pyplot"""
        st.subheader("📈 Visual Analysis")

        df = self.generate_year_wise_data(inputs, results)

        fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(15, 12))

        # 1. Net Cost Comparison - ADD OPTION 3
        options = ['Option 1\n(Own Capital)', 'Option 2\n(Invest + Loan)', 'Option 3\n(Custom + Loan)']
        costs = [results['option1']['net_outflow'], results['option2']['net_outflow'], results['option3']['net_outflow']]
        
        # Dynamic colors based on recommendation
        colors = []
        if results['recommendation'] == 'option1': colors.extend(['#2E8B57', '#FF6B6B', '#FF6B6B'])
        elif results['recommendation'] == 'option2': colors.extend(['#FF6B6B', '#2E8B57', '#FF6B6B'])
        else: colors.extend(['#FF6B6B', '#FF6B6B', '#2E8B57'])

        bars1 = ax1.bar(options, costs, color=colors, alpha=0.7, edgecolor='black')
        ax1.set_title('Net Cost Comparison', fontsize=14, fontweight='bold')
        ax1.set_ylabel('Net Cost (₹ lakh)')

        for bar, cost in zip(bars1, costs):
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width()/2., height + 0.5,
                     f'₹{cost:.1f}L', ha='center', va='bottom', fontweight='bold')

        # 2. EMI Comparison (add Option 3 EMI)
        emis = [results['option1']['emi'], results['option2']['emi'], results['option3']['emi']]
        # Adjust colors for EMI if needed, or keep generic
        emi_colors = ['#4CAF50', '#FF9800', '#2196F3']
        bars2 = ax2.bar(options, emis, color=emi_colors, alpha=0.7, edgecolor='black')
        ax2.set_title('Monthly EMI Comparison', fontsize=14, fontweight='bold')
        ax2.set_ylabel('Monthly EMI (₹)')

        for bar, emi in zip(bars2, emis):
            height = bar.get_height()
            ax2.text(bar.get_x() + bar.get_width()/2., height + 5000,
                     f'₹{emi:,.0f}', ha='center', va='bottom', fontweight='bold')


        # 3. Year-wise Analysis - ADD OPTION 3 Interest and Investment Value
        ax3.plot(df['Year'], df['Option1_Interest_Paid'], marker='o',
                 linewidth=2, label='Option 1 Interest', color='#2196F3')
        ax3.plot(df['Year'], df['Option2_Interest_Paid'], marker='s',
                 linewidth=2, label='Option 2 Interest', color='#F44336')
        ax3.plot(df['Year'], df['Option3_Interest_Paid'], marker='D',
                 linewidth=2, label='Option 3 Interest', color='#9C27B0')
        ax3.plot(df['Year'], df['Investment_Value_Option2'], marker='^',
                 linewidth=2, label='Option 2 Investment Value', color='#4CAF50')
        ax3.plot(df['Year'], df['Investment_Value_Option3'], marker='v',
                 linewidth=2, label='Option 3 Investment Value', color='#FFC107')


        ax3.set_title('Cost & Investment Evolution Over Time', fontsize=14, fontweight='bold')
        ax3.set_xlabel('Year')
        ax3.set_ylabel('Amount (₹ lakh)')
        ax3.legend()
        ax3.grid(True, alpha=0.3)

        # 4. Break-even Analysis - ADD OPTION 3 Net Cost
        ax4.plot(df['Year'], df['Option1_Net_Cost_Cumulative'],
                 linewidth=3, label='Option 1 Net Cost', color='#2196F3', linestyle='--')
        ax4.plot(df['Year'], df['Option2_Net_Cost_Cumulative'], marker='o',
                 linewidth=2, label='Option 2 Net Cost', color='#F44336')
        ax4.plot(df['Year'], df['Option3_Net_Cost_Cumulative'], marker='D',
                 linewidth=2, label='Option 3 Net Cost', color='#9C27B0', linestyle='-.')


        ax4.set_title('Cumulative Net Cost Over Time', fontsize=14, fontweight='bold')
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
                    'Loan Type',
                    'Loan Interest Tax Deductible',
                    'Prepayment Penalty (%)',
                    'Minimum Liquidity Target (₹ lakh)',
                    'Tenure (years)', 'Investment Type', 'Investment Return (%)', 'Tax Rate (%)',
                    'Custom Capital Contribution (Option 3) (₹ lakh)',
                    '', 'OPTION 1 - Use Own Capital',
                    'Capital Used Directly for Project (₹ lakh)',
                    'Own Capital Invested (₹ lakh)',
                    'Loan Amount (₹ lakh)', 'Monthly EMI (₹)',
                    'Gross Total Interest (₹ lakh)', 'Effective Total Interest (After Tax) (₹ lakh)',
                    'Total Investment (Gross Cash Outflow) (₹ lakh)', # NEW
                    'Net Cost (₹ lakh)',
                    '', 'OPTION 2 - Invest + Loan',
                    'Capital Used Directly for Project (₹ lakh)',
                    'Own Capital Invested (₹ lakh)',
                    'Loan Amount (₹ lakh)', 'Monthly EMI (₹)',
                    'Gross Total Interest (₹ lakh)', 'Effective Total Interest (After Tax) (₹ lakh)',
                    'Investment Maturity (₹ lakh)', 'Post-tax Gain (₹ lakh)',
                    'Total Investment (Gross Cash Outflow) (₹ lakh)', # NEW
                    'Net Cost (₹ lakh)',
                    '', 'OPTION 3 - Custom Capital Contribution + Loan',
                    'Capital Used Directly for Project (₹ lakh)',
                    'Own Capital Invested (₹ lakh)',
                    'Loan Amount (₹ lakh)',
                    'Monthly EMI (₹)',
                    'Gross Total Interest (₹ lakh)', 'Effective Total Interest (After Tax) (₹ lakh)',
                    'Investment Maturity (₹ lakh) (Option 3)',
                    'Post-tax Gain (₹ lakh) (Option 3)',
                    'Total Investment (Gross Cash Outflow) (₹ lakh)', # NEW
                    'Net Cost (₹ lakh) (Option 3)',
                    '', 'Overall Effective Loan Rate (After Tax) (%) (Annualized)', # Updated label
                    'Overall Effective Investment Return (After Tax) (%) (Annualized)', # Updated label
                    'Estimated Prepayment Cost (₹ lakh)',
                    '', 'RECOMMENDATION', 'Better Option', 'Savings (₹ lakh)'
                ],
                'Value': [
                    inputs['project_cost'], inputs['own_capital'], inputs['loan_rate'],
                    inputs['loan_type'],
                    'Yes' if inputs['loan_interest_deductible'] else 'No',
                    inputs['prepayment_penalty_pct'],
                    inputs['min_liquidity_target'],
                    inputs['loan_tenure'], inputs['investment_type'], inputs['investment_return'],
                    inputs['tax_rate'],
                    inputs['custom_capital_contribution'],
                    '', '',
                    results['option1']['capital_used'],
                    0.0,
                    results['option1']['loan_amount'],
                    f"₹{results['option1']['emi']:,.0f}", results['option1']['total_interest'],
                    results['option1']['effective_interest'],
                    results['option1']['total_investment_gross'], # NEW
                    results['option1']['net_outflow'], '', '',
                    results['option2']['capital_used'],
                    results['option2']['capital_invested'],
                    results['option2']['loan_amount'],
                    f"₹{results['option2']['emi']:,.0f}", results['option2']['total_interest'],
                    results['option2']['effective_interest'],
                    results['option2']['investment_maturity'], results['option2']['post_tax_gain'],
                    results['option2']['total_investment_gross'], # NEW
                    results['option2']['net_outflow'],
                    '', '',
                    results['option3']['capital_used'],
                    results['option3']['remaining_own_capital_invested'],
                    results['option3']['loan_amount'],
                    f"₹{results['option3']['emi']:,.0f}",
                    results['option3']['total_interest'],
                    results['option3']['effective_interest'],
                    results['option3']['investment_maturity'],
                    results['option3']['post_tax_gain'],
                    results['option3']['total_investment_gross'], # NEW
                    results['option3']['net_outflow'],
                    '', results['effective_loan_rate_annual'],
                    results['effective_investment_return_annual'],
                    results['prepayment_cost'],
                    '',
                    'Option 1' if results['recommendation'] == 'option1' else ('Option 2' if results['recommendation'] == 'option2' else 'Option 3'),
                    results['savings']
                ]
            }

            pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)

            # Year-wise analysis (df already contains the new columns from generate_year_wise_data)
            df.to_excel(writer, sheet_name='Year_wise_Analysis', index=False)

            # Investment options reference
            inv_df = pd.DataFrame.from_dict(self.investment_options, orient='index')
            inv_df.to_excel(writer, sheet_name='Investment_Options')

            # Ensure all sheets are visible and the first one is active
            if writer.book.sheetnames:
                for sheet_name in writer.book.sheetnames:
                    ws = writer.book[sheet_name]
                    ws.sheet_state = 'visible' # Explicitly set visibility
                writer.book.active = 0 # Set the first sheet as active

        output.seek(0) # Rewind the buffer to the beginning

        st.download_button(
            label="Download as Excel",
            data=output,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Click to download the detailed financial analysis in Excel format."
        )

    def export_to_csv(self, df, filename='analysis_data.csv', label="Download Year-wise CSV"):
        """Export DataFrame to CSV and provide a download button"""
        csv_output = df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label=label,
            data=csv_output,
            file_name=filename,
            mime="text/csv",
            key=f"csv_download_{filename.replace('.', '_')}" # Ensure unique key
        )

    def generate_pdf_report(self, inputs, results, comparison_df, year_wise_df, filename='project_financing_report.pdf'):
        """Generates a PDF report using ReportLab."""
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=letter,
                                rightMargin=inch/2, leftMargin=inch/2,
                                topMargin=inch/2, bottomMargin=inch/2)
        styles = getSampleStyleSheet()
        story = []

        # Title
        story.append(Paragraph("Project Financing Analysis Report", styles['h1']))
        story.append(Spacer(1, 0.2 * inch))

        # Recommendation
        story.append(Paragraph("Recommendation:", styles['h2']))
        story.append(Paragraph(self.get_recommendation_text(results, inputs), styles['Normal']))
        story.append(Paragraph(f"Potential Savings (compared to worst option): ₹{results['savings']:.2f} lakh", styles['Normal']))
        story.append(Spacer(1, 0.2 * inch))

        # Financing Options Comparison Table
        story.append(Paragraph("Financing Options Comparison:", styles['h2']))
        
        # Prepare data for ReportLab Table
        table_data = [list(comparison_df.columns)] + comparison_df.reset_index().values.tolist()
        
        # Apply formatting for ReportLab Table
        # Convert all values to string for table rendering
        table_data_str = [[str(item) for item in row] for row in table_data]

        table_style = TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#F0F2F6')), # Header background
            ('TEXTCOLOR', (0,0), (-1,0), colors.black),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0,0), (-1,0), 12),
            ('BACKGROUND', (0,1), (-1,-1), colors.white),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('FONTSIZE', (0,0), (-1,-1), 9),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('LEFTPADDING', (0,0), (-1,-1), 6),
            ('RIGHTPADDING', (0,0), (-1,-1), 6),
            ('TOPPADDING', (0,0), (-1,-1), 6),
            ('BOTTOMPADDING', (0,0), (-1,-1), 6),
            # Highlight NET COST row
            ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
            ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#E0FFE0')), # Light green for best option
        ])

        table = Table(table_data_str, colWidths=[1.8*inch] + [1.8*inch]*3)
        table.setStyle(table_style)
        story.append(table)
        story.append(Spacer(1, 0.2 * inch))

        # Net Cost Definition
        story.append(Paragraph("What are 'Net Cost' and 'Total Investment' in this context?", styles['h3'])) # Updated title
        story.append(Paragraph("""
        <b>"Net Cost"</b> represents your ultimate <b>out-of-pocket expense to fund the project</b> over the loan tenure.
        It is calculated as:
        `Net Cost = Total Project Cost + Effective Total Loan Interest (After Tax) - Post-Tax Investment Gains`
        """, styles['Normal']))
        story.append(Paragraph("""
        Let's break that down:
        <ul>
            <li><b>Total Project Cost:</b> The fundamental cost of acquiring or building the project.</li>
            <li><b>Effective Total Loan Interest (After Tax):</b> This is the gross interest paid on the loan, reduced by any tax savings if the loan interest is tax-deductible as a business expense (based on your input `Tax Rate`).</li>
            <li><b>Post-Tax Investment Gains:</b> Any net profits you earn from investing your own capital (after taxes) are subtracted, as these gains effectively reduce your overall cost.</li>
        </ul>
        This metric provides a true bottom-line figure for each financing option, allowing for direct and comparable evaluation of what you <i>truly spent</i> to get the project done.
        """, styles['Normal']))
        story.append(Spacer(1, 0.1 * inch))
        story.append(Paragraph("""
        ---
        """, styles['Normal']))
        story.append(Spacer(1, 0.1 * inch))
        story.append(Paragraph("""
        <b>"Total Investment (Gross Cash Outflow)"</b> represents the <b>total cash flow involved</b> in financing the project and managing your available capital over the loan tenure.
        It is calculated as:
        `Total Investment = Total Loan Payment (Principal + Gross Interest) + Own Capital Available`
        """, styles['Normal']))
        story.append(Paragraph("""
        This figure is typically much higher than "Net Cost" because it does not account for tax benefits on interest or offsetting investment gains. It gives a sense of the gross funds that move through your financing strategy.
        """, styles['Normal']))
        story.append(Spacer(1, 0.2 * inch))


        # Input Parameters Used
        story.append(Paragraph("Input Parameters Used:", styles['h2']))
        input_params = [
            ["Total Project Cost:", f"₹{inputs['project_cost']:.1f} lakh"],
            ["Own Capital Available:", f"₹{inputs['own_capital']:.1f} lakh"],
            ["Bank Loan Interest Rate:", f"{inputs['loan_rate']:.2f}% p.a. ({inputs['loan_type']})"],
            ["Loan Tenure:", f"{inputs['loan_tenure']} years"],
            ["Loan Interest Tax Deductible:", 'Yes' if inputs['loan_interest_deductible'] else 'No'],
            ["Prepayment Penalty:", f"{inputs['prepayment_penalty_pct']:.2f}%"],
            ["Minimum Liquidity Target:", f"₹{inputs['min_liquidity_target']:.1f} lakh"],
            ["Investment Type:", inputs['investment_type']],
            ["Investment Return:", f"{inputs['investment_return']:.2f}% p.a."],
            ["Tax Rate:", f"{inputs['tax_rate']:.0f}%"]
        ]
        if inputs['custom_capital_input_type'] == 'Value':
            input_params.append(["Custom Capital Contribution (Option 3):", f"₹{inputs['custom_capital_contribution']:.1f} lakh"])
        else:
            input_params.append(["Custom Capital Contribution (Option 3):", f"{inputs['custom_capital_percentage']:.1f}% of Own Capital (₹{inputs['custom_capital_contribution']:.1f} lakh)"])
        
        # Investment Details
        investment_details = self.investment_options[inputs['investment_type']]
        input_params.append(["", ""]) # Spacer
        input_params.append([Paragraph("<b>Investment Details:</b>", styles['Normal']), ""]) # Bold in PDF
        input_params.append([" - Liquidity:", investment_details['liquidity']])
        input_params.append([" - Tax Efficiency:", investment_details['tax_efficiency']])
        input_params.append([" - Compounding:", investment_details['compounding']])

        table_input_params = Table(input_params, colWidths=[2.5*inch, 3.5*inch])
        table_input_params.setStyle(TableStyle([
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('FONTSIZE', (0,0), (-1,-1), 9),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
            ('TOPPADDING', (0,0), (-1,-1), 2),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2),
            ('FONTNAME', (0,0), (0,-1), 'Helvetica'), # Removed bold from first column, will use Paragraph for specific bolding
        ]))
        story.append(table_input_params)
        story.append(Spacer(1, 0.2 * inch))

        # Key Insights
        story.append(Paragraph("Key Insights & Strategic Considerations:", styles['h2']))
        story.append(Paragraph(f"Effective Loan Interest Rate (After Tax) (Annualized): <b>{results['effective_loan_rate_annual']:.2f}% p.a.</b>", styles['Normal']))
        story.append(Paragraph(f"Effective Investment Return (After Tax) (Annualized): <b>{results['effective_investment_return_annual']:.2f}% p.a.</b>", styles['Normal']))
        story.append(Paragraph(f"Interest Rate Spread (Loan - Investment): <b>{results['interest_spread']:.2f}%</b>", styles['Normal']))
        story.append(Paragraph("These annualized rates provide a clearer picture of the true cost of borrowing and the actual return on your investments, factoring in tax benefits.", styles['Normal']))
        story.append(Spacer(1, 0.1 * inch))

        # Risk & Scenario Analysis
        story.append(Paragraph("Risk & Scenario Analysis:", styles['h3']))
        story.append(Paragraph(f" - Loan Interest Rate Risk: The loan is <b>{inputs['loan_type']}</b>. If floating, consider sensitivity to rate hikes and build in buffers.", styles['Normal']))
        story.append(Paragraph(f" - Prepayment Cost: A {inputs['prepayment_penalty_pct']:.2f}% penalty on a full loan (Option 2) would be <b>₹{results['prepayment_cost']:.2f} lakh</b>. Factor this into early exit scenarios and loan terms.", styles['Normal']))
        story.append(Paragraph(" - Investment Volatility: Assumed investment returns are estimates. Stress test your plan with &plusmn;1-2% returns to understand the impact on net cost and liquidity.", styles['Normal']))
        story.append(Paragraph(f" - Liquidity Buffer: Your target minimum liquidity is <b>₹{inputs['min_liquidity_target']:.1f} lakh</b>. Ensure the chosen option consistently maintains this, especially under stressed cash flow scenarios (e.g., delayed project revenue, unexpected expenses).", styles['Normal']))
        story.append(Spacer(1, 0.1 * inch))

        # Nuanced Strategic Considerations
        story.append(Paragraph("Nuanced Strategic Considerations:", styles['h3']))
        if inputs['loan_interest_deductible']:
            story.append(Paragraph(" - Tax Planning: Loan interest is considered tax-deductible, significantly reducing the effective cost of borrowing. Ensure proper documentation for claiming this deduction.", styles['Normal']))
        else:
            story.append(Paragraph(" - Tax Planning: Loan interest is NOT considered tax-deductible, meaning the gross interest is the effective cost. Explore other tax optimization strategies.", styles['Normal']))
        story.append(Paragraph(" - Credit Profile & Leverage: Assess the impact of increased debt on your company's debt-to-equity ratio, credit rating, and future borrowing capacity. Excessive leverage can affect banking relationships and future funding flexibility.", styles['Normal']))
        story.append(Paragraph(" - Regulatory Compliance: Confirm all documentation (project reports, audited financials) are ready for smooth loan disbursal and subsequent annual reviews to avoid penalties or delays.", styles['Normal']))
        story.append(Paragraph(" - Investment Liquidity/Market Risk: Even 'liquid' investments carry some market risk (e.g., temporary impairment during market freezes). Maintain an emergency buffer in a highly liquid bank account (beyond investment) for immediate needs.", styles['Normal']))
        story.append(Spacer(1, 0.1 * inch))

        # Optimization Moves & Additional Value Levers
        story.append(Paragraph("Optimization Moves & Additional Value Levers:", styles['h3']))
        story.append(Paragraph(" - Staggered Drawdown: Consider phased loan drawdown or capital deployment as per project schedule to optimize interest outflow and minimize idle funds.", styles['Normal']))
        story.append(Paragraph(" - Step-up EMI/SWP Option: If expecting higher business income post-project commissioning, opt for step-up EMI structure. Alternatively, consider Systematic Withdrawal Plan (SWP) from investments to sync cash flows with project needs.", styles['Normal']))
        story.append(Paragraph(" - Insurance: Insure the principal loan with adequate keyman/credit insurance, especially if individual guarantees are required, to mitigate personal risk.", styles['Normal']))
        story.append(Paragraph(" - Renegotiation Flexibility: Build in terms for possible loan prepayment, refinancing, or top-up without penal costs in your loan agreements.", styles['Normal']))
        story.append(Paragraph(" - AR/AP Synchronization: Consider peer business working capital requirements—using part of the liquid reserves to cover temporary spikes in Accounts Receivable (AR) / Accounts Payable (AP), thus reducing the need for additional short-term borrowings.", styles['Normal']))
        story.append(Spacer(1, 0.2 * inch))

        # Strategic Summary
        story.append(Paragraph("Strategic Summary: Best Practices for CFOs", styles['h2']))
        story.append(Paragraph(f" - Quantify & Monitor Liquidity: Always maintain a minimum liquidity threshold (e.g., <b>₹{inputs['min_liquidity_target']:.1f} lakh</b> always in instantly-accessible form) to ensure business agility and emergency preparedness.", styles['Normal']))
        story.append(Paragraph(" - Dynamic Financial Planning: Be prepared for faster loan prepayment or capital recycling if the business environment or investment returns shift in future years. Agility is key.", styles['Normal']))
        story.append(Paragraph(" - Continuous Benchmarking: Regularly benchmark your investment returns and borrowing costs against evolving market rates and new reinvestment opportunities to ensure optimal capital allocation.", styles['Normal']))
        story.append(Paragraph(" - Dashboard Monitoring: Implement and maintain a dynamic dashboard updated for promoters, CFO, and the board on all key financial parameters and risk indicators.", styles['Normal']))
        story.append(Spacer(1, 0.2 * inch))

        # In short
        story.append(Paragraph("In short:", styles['h3']))
        story.append(Paragraph("""
        This section provides a concise, executive summary of the holistic advisory approach. It distills the comprehensive analysis into key actionable takeaways, ensuring that even busy stakeholders can quickly grasp the core value proposition: transforming "cost minimization" into value optimization—giving the CFO not just the "cheapest" route, but the safest and most strategic path for long-term business health and growth.
        """, styles['Normal']))
        
        # Add year-wise data as a new page
        story.append(PageBreak())
        story.append(Paragraph("Year-wise Analysis:", styles['h2']))
        year_wise_data_list = [list(year_wise_df.columns)] + year_wise_df.values.tolist()
        year_wise_data_str = [[str(f"{item:.2f}" if isinstance(item, (int, float)) else item) for item in row] for row in year_wise_data_list] # Format numbers

        year_wise_table_style = TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#F0F2F6')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.black),
            ('ALIGN', (0,0), (-1,-1), 'RIGHT'), # Align numbers right
            ('ALIGN', (0,0), (0,-1), 'LEFT'), # Align Year column left
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0,0), (-1,0), 12),
            ('BACKGROUND', (0,1), (-1,-1), colors.white),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('FONTSIZE', (0,0), (-1,-1), 8), # Smaller font for year-wise data
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('LEFTPADDING', (0,0), (-1,-1), 2),
            ('RIGHTPADDING', (0,0), (-1,-1), 2),
            ('TOPPADDING', (0,0), (-1,-1), 2),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ])
        year_wise_table = Table(year_wise_data_str, colWidths=[0.5*inch] + [0.9*inch]*12) # Adjust column widths
        year_wise_table.setStyle(year_wise_table_style)
        story.append(year_wise_table)


        doc.build(story)
        buffer.seek(0)
        return buffer.getvalue()


    def export_to_csv(self, df, filename='analysis_data.csv', label="Download Year-wise CSV"):
        """Export DataFrame to CSV and provide a download button"""
        csv_output = df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label=label,
            data=csv_output,
            file_name=filename,
            mime="text/csv",
            key=f"csv_download_{filename.replace('.', '_')}" # Ensure unique key
        )

    def export_options_section(self, inputs, results, comparison_df, year_wise_df):
        """Creates a dedicated section for export options."""
        st.markdown("---")
        st.subheader("📥 Export Options")
        
        col_export_1, col_export_2, col_export_3 = st.columns(3)

        with col_export_1:
            self.export_to_excel(inputs, results, year_wise_df, filename='Project_Financing_Analysis.xlsx')
        
        with col_export_2:
            # CSV for Comparison Table
            csv_comparison_output = comparison_df.to_csv().encode('utf-8') # Use .to_csv() directly on comparison_df
            st.download_button(
                label="Download Comparison CSV",
                data=csv_comparison_output,
                file_name="Financing_Comparison.csv",
                mime="text/csv",
                key="csv_download_comparison"
            )
            # CSV for Year-wise Data
            self.export_to_csv(year_wise_df, filename='Year_Wise_Analysis.csv', label="Download Year-wise CSV")

        with col_export_3:
            pdf_data = self.generate_pdf_report(inputs, results, comparison_df, year_wise_df)
            st.download_button(
                label="Download Report as PDF",
                data=pdf_data,
                file_name="Project_Financing_Report.pdf",
                mime="application/pdf",
                key="pdf_download"
            )
        
        st.info("Note: Direct PowerPoint (PPT) export is a complex feature and is not currently supported. You can use the PDF report for presentation purposes.")


def main():
    """Main function to run the calculator with Streamlit UI"""
    st.set_page_config(layout="wide", page_title="Project Financing Calculator")
    st.title("🏦 Project Financing Calculator")
    st.markdown("""
        This tool helps you compare three project financing strategies, incorporating key financial and strategic considerations:
        1.  **Option 1:** Use all your own capital first, then take a loan for the remaining project cost.
        2.  **Option 2:** Invest all your own capital and take a full loan for the entire project cost.
        3.  **Option 3:** You specify a custom amount of your own capital to use directly, and take a loan for the rest. Any remaining own capital is invested.
        
        Enter your project details below to see a detailed comparison and recommendation, along with advanced insights for a CFO.
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
        # NEW INPUTS FOR ADVANCED ANALYSIS
        loan_type = st.selectbox(
            "Loan Interest Rate Type:",
            options=['Fixed', 'Floating'],
            index=0, # Default to Fixed
            help="Is the loan interest rate fixed or floating? Important for interest rate risk."
        )
        loan_interest_deductible = st.checkbox(
            "Is loan interest tax deductible as a business expense?",
            value=True, # Default to True for typical business loans
            help="Check if the interest paid on the loan can be deducted from taxable income."
        )
        prepayment_penalty_pct = st.number_input(
            "Prepayment Penalty (%) on principal:",
            min_value=0.0, value=2.0, step=0.1, format="%.2f",
            help="Percentage penalty if the loan is repaid early (e.g., 2 for 2%)."
        )
        min_liquidity_target = st.number_input(
            "Minimum Liquidity Target (₹ lakh):",
            min_value=0.0, value=20.0, step=5.0, format="%.1f",
            help="Target amount of cash/liquid assets to maintain for contingencies."
        )


        st.markdown("---")
        st.subheader("Option 3: Custom Capital Contribution")
        
        custom_capital_input_type = st.radio(
            "How to specify Custom Capital Contribution?",
            ('Value', 'Percentage of Own Capital'),
            key="custom_capital_input_type_radio"
        )

        custom_capital_contribution = 0.0
        custom_capital_percentage = 0.0

        if custom_capital_input_type == 'Value':
            custom_capital_contribution = st.number_input(
                "Amount of your capital to use directly for project (₹ lakh):",
                min_value=0.0,
                max_value=min(project_cost, own_capital),
                value=min(project_cost, own_capital),
                step=5.0,
                format="%.1f",
                help="Specify how much of your own capital you want to use directly for the project. The rest will be loaned."
            )
        else: # Percentage of Own Capital
            custom_capital_percentage = st.slider(
                "Percentage of your Own Capital to use directly for project:",
                min_value=0, max_value=100, value=int((min(project_cost, own_capital) / own_capital) * 100) if own_capital > 0 else 0,
                step=5, format="%d%%",
                help="Specify what percentage of your available own capital to use directly. The rest will be loaned."
            )
            custom_capital_contribution = (custom_capital_percentage / 100) * own_capital
            # Ensure it doesn't exceed project cost
            custom_capital_contribution = min(custom_capital_contribution, project_cost)


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
        'loan_type': loan_type,
        'loan_interest_deductible': loan_interest_deductible,
        'prepayment_penalty_pct': float(prepayment_penalty_pct),
        'min_liquidity_target': float(min_liquidity_target),
        'investment_type': investment_type,
        'investment_return': float(investment_return),
        'tax_rate': float(tax_rate),
        'custom_capital_contribution': float(custom_capital_contribution),
        'custom_capital_input_type': custom_capital_input_type,
        'custom_capital_percentage': float(custom_capital_percentage)
    }

    st.markdown("---")

    # Add a button to trigger calculation
    if st.button("Calculate Financing Options", type="primary"):
        # Perform calculations
        results = calc.calculate_comparison(inputs)

        # Display detailed report
        calc.print_detailed_report(inputs, results)

        # Create and display visualizations
        df = calc.create_visualizations(inputs, results)

        # Prepare comparison_df for export
        comparison_data_for_export = {
            'Metric': [
                'Capital Used Directly for Project',
                'Own Capital Invested',
                'Loan Amount',
                'Monthly EMI',
                'Gross Total Interest',
                'Effective Total Interest (After Tax)',
                'Investment Maturity Value (from invested capital)',
                'Gross Investment Gain (from invested capital)',
                'Post-Tax Investment Gain (from invested capital)',
                'Total Investment (Gross Cash Outflow)',
                'NET COST'
            ],
            'Option 1: Own Capital + Small Loan': [
                f"₹{results['option1']['capital_used']:.1f} lakh",
                "₹0.0 lakh",
                f"₹{results['option1']['loan_amount']:.1f} lakh",
                f"₹{results['option1']['emi']:,.0f}",
                f"₹{results['option1']['total_interest']:.1f} lakh",
                f"₹{results['option1']['effective_interest']:.1f} lakh",
                "₹0.0 lakh",
                "₹0.0 lakh",
                "₹0.0 lakh",
                f"₹{results['option1']['total_investment_gross']:.2f} lakh",
                f"₹{results['option1']['net_outflow']:.2f} lakh"
            ],
            'Option 2: Invest Capital + Full Loan': [
                "₹0.0 lakh",
                f"₹{results['option2']['capital_invested']:.1f} lakh",
                f"₹{results['option2']['loan_amount']:.1f} lakh",
                f"₹{results['option2']['emi']:,.0f}",
                f"₹{results['option2']['total_interest']:.1f} lakh",
                f"₹{results['option2']['effective_interest']:.1f} lakh",
                f"₹{results['option2']['investment_maturity']:.2f} lakh",
                f"₹{results['option2']['investment_gain']:.2f} lakh",
                f"₹{results['option2']['post_tax_gain']:.2f} lakh",
                f"₹{results['option2']['total_investment_gross']:.2f} lakh",
                f"₹{results['option2']['net_outflow']:.2f} lakh"
            ],
            'Option 3: Custom Capital + Loan': [
                f"₹{results['option3']['capital_used']:.1f} lakh",
                f"₹{results['option3']['remaining_own_capital_invested']:.1f} lakh",
                f"₹{results['option3']['loan_amount']:.1f} lakh",
                f"₹{results['option3']['emi']:,.0f}",
                f"₹{results['option3']['total_interest']:.1f} lakh",
                f"₹{results['option3']['effective_interest']:.1f} lakh",
                f"₹{results['option3']['investment_maturity']:.2f} lakh" if results['option3']['remaining_own_capital_invested'] > 0 else "₹0.0 lakh",
                f"₹{results['option3']['investment_gain']:.2f} lakh" if results['option3']['remaining_own_capital_invested'] > 0 else "₹0.0 lakh",
                f"₹{results['option3']['post_tax_gain']:.2f} lakh" if results['option3']['remaining_own_capital_invested'] > 0 else "₹0.0 lakh",
                f"₹{results['option3']['total_investment_gross']:.2f} lakh",
                f"₹{results['option3']['net_outflow']:.2f} lakh"
            ]
        }
        comparison_df_for_export = pd.DataFrame(comparison_data_for_export).set_index('Metric')


        # Provide Export Options
        calc.export_options_section(inputs, results, comparison_df_for_export, df)

        st.success("Calculations complete! Review the detailed analysis and visualizations below.")

if __name__ == "__main__":
    main()
