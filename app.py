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
        # Basic input validation
        if principal < 0 or rate < 0 or tenure_years < 0:
            st.error("Principal, rate, and tenure must be non-negative.")
            return 0

        if principal == 0:
            return 0

        monthly_rate = rate / (12 * 100)
        months = tenure_years * 12

        if rate == 0:
            # If rate is 0, EMI is simply principal / months. Handle division by zero for months.
            return principal / months if months > 0 else principal

        try:
            emi = (principal * monthly_rate * math.pow(1 + monthly_rate, months)) / \
                  (math.pow(1 + monthly_rate, months) - 1)
        except OverflowError:
            st.error("EMI calculation resulted in an overflow. Check your inputs (e.g., extremely high rate or tenure).")
            return float('inf') # Indicate a very large EMI
        except ZeroDivisionError:
            st.error("EMI calculation resulted in division by zero. Check your inputs (e.g., monthly rate leading to zero denominator).")
            return float('inf')
            
        return emi

    def calculate_investment_growth(self, principal, annual_rate, years, compounding='quarterly'):
        """Calculate investment growth with different compounding frequencies"""
        # Basic input validation
        if principal < 0 or annual_rate < 0 or years < 0:
            st.error("Principal, annual rate, and years for investment must be non-negative.")
            return 0

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
            'Scenario1_Interest_Paid_Cumulative (Lakh)',
            'Scenario2_Interest_Paid_Cumulative (Lakh)',
            'Scenario3_Interest_Paid_Cumulative (Lakh)',
            'Investment_Value_Scenario2 (Lakh)',
            'Investment_Value_Scenario3 (Lakh)',
            'Investment_Gain_Scenario2 (Lakh)',
            'Investment_Gain_Scenario3 (Lakh)',
            'Post_Tax_Gain_Scenario2 (Lakh)',
            'Post_Tax_Gain_Scenario3 (Lakh)',
            'Scenario1_Net_Effective_Cost_Cumulative (Lakh)',
            'Scenario2_Net_Effective_Cost_Cumulative (Lakh)',
            'Scenario3_Net_Effective_Cost_Cumulative (Lakh)'
        ]
        data = []

        # Ensure project cost is consistently used for calculations (in actual rupees for EMI)
        project_cost_actual = inputs['project_cost'] * 100000
        own_capital_lakh = inputs['own_capital'] # Keep in Lakhs for investment growth
        
        # Calculate EMIs once for the full tenure to avoid recalculating in loop
        # Scenario 1 Loan
        scenario1_loan_amount_actual = max(0, inputs['project_cost'] - own_capital_lakh) * 100000
        scenario1_emi = self.calculate_emi(scenario1_loan_amount_actual, inputs['loan_rate'], inputs['loan_tenure'])

        # Scenario 2 Loan
        scenario2_loan_amount_actual = inputs['project_cost'] * 100000
        scenario2_emi = self.calculate_emi(scenario2_loan_amount_actual, inputs['loan_rate'], inputs['loan_tenure'])

        # Scenario 3 Loan
        scenario3_capital_used_lakh = inputs['custom_capital_contribution']
        scenario3_loan_amount_actual = max(0, inputs['project_cost'] - scenario3_capital_used_lakh) * 100000
        scenario3_emi = self.calculate_emi(scenario3_loan_amount_actual, inputs['loan_rate'], inputs['loan_tenure'])
        
        scenario3_remaining_own_capital_lakh = own_capital_lakh - scenario3_capital_used_lakh


        for year in range(inputs['loan_tenure'] + 1):
            # Scenario 1: Maximum Own Funding
            if year == 0:
                s1_interest_paid_cumulative = 0
                s1_net_effective_cost_cumulative = inputs['project_cost'] # Only project cost at year 0
            else:
                s1_total_paid_cumulative = scenario1_emi * 12 * year
                s1_principal_paid_cumulative = min(s1_total_paid_cumulative, scenario1_loan_amount_actual)
                s1_interest_paid_cumulative = max(0, s1_total_paid_cumulative - s1_principal_paid_cumulative) / 100000 # Convert to lakh
                
                s1_effective_interest_cumulative = s1_interest_paid_cumulative
                if inputs['loan_interest_deductible']:
                    s1_effective_interest_cumulative *= (1 - inputs['tax_rate']/100)
                s1_net_effective_cost_cumulative = inputs['project_cost'] + s1_effective_interest_cumulative


            # Scenario 2: Maximum Leverage
            if year == 0:
                s2_interest_paid_cumulative = 0
                s2_investment_value = own_capital_lakh # Initial investment value
                s2_investment_gain = 0
                s2_post_tax_gain = 0
                s2_net_effective_cost_cumulative = inputs['project_cost'] # Only project cost at year 0
            else:
                s2_total_paid_cumulative = scenario2_emi * 12 * year
                s2_principal_paid_cumulative = min(s2_total_paid_cumulative, scenario2_loan_amount_actual)
                s2_interest_paid_cumulative = max(0, s2_total_paid_cumulative - s2_principal_paid_cumulative) / 100000 # Convert to lakh

                s2_effective_interest_cumulative = s2_interest_paid_cumulative
                if inputs['loan_interest_deductible']:
                    s2_effective_interest_cumulative *= (1 - inputs['tax_rate']/100)
                
                s2_investment_value = self.calculate_investment_growth(
                    own_capital_lakh,
                    inputs['investment_return'],
                    year,
                    self.investment_options[inputs['investment_type']]['compounding']
                )
                s2_investment_gain = s2_investment_value - own_capital_lakh
                s2_post_tax_gain = s2_investment_gain * (1 - inputs['tax_rate']/100)
                
                s2_net_effective_cost_cumulative = inputs['project_cost'] + s2_effective_interest_cumulative - s2_post_tax_gain


            # Scenario 3: Balanced Approach
            if year == 0:
                s3_interest_paid_cumulative = 0
                s3_investment_value = scenario3_remaining_own_capital_lakh if scenario3_remaining_own_capital_lakh > 0 else 0
                s3_investment_gain = 0
                s3_post_tax_gain = 0
                s3_net_effective_cost_cumulative = inputs['project_cost'] # Only project cost at year 0
            else:
                s3_total_paid_cumulative = scenario3_emi * 12 * year
                s3_principal_paid_cumulative = min(s3_total_paid_cumulative, scenario3_loan_amount_actual)
                s3_interest_paid_cumulative = max(0, s3_total_paid_cumulative - s3_principal_paid_cumulative) / 100000 # Convert to lakh

                s3_effective_interest_cumulative = s3_interest_paid_cumulative
                if inputs['loan_interest_deductible']:
                    s3_effective_interest_cumulative *= (1 - inputs['tax_rate']/100)

                s3_investment_value = scenario3_remaining_own_capital_lakh # Initial value if no growth
                if scenario3_remaining_own_capital_lakh > 0:
                    s3_investment_value = self.calculate_investment_growth(
                        scenario3_remaining_own_capital_lakh,
                        inputs['investment_return'],
                        year,
                        self.investment_options[inputs['investment_type']]['compounding']
                    )
                s3_investment_gain = s3_investment_value - scenario3_remaining_own_capital_lakh
                s3_post_tax_gain = s3_investment_gain * (1 - inputs['tax_rate']/100)
                
                s3_net_effective_cost_cumulative = inputs['project_cost'] + s3_effective_interest_cumulative - s3_post_tax_gain


            data.append([
                year,
                s1_interest_paid_cumulative,
                s2_interest_paid_cumulative,
                s3_interest_paid_cumulative,
                s2_investment_value,
                s3_investment_value,
                s2_investment_gain,
                s3_investment_gain,
                s2_post_tax_gain,
                s3_post_tax_gain,
                s1_net_effective_cost_cumulative,
                s2_net_effective_cost_cumulative,
                s3_net_effective_cost_cumulative
            ])

        # Create DataFrame with explicit column names
        return pd.DataFrame(data, columns=column_names)

    def calculate_comparison(self, inputs):
        """Main calculation function to compare the three financing scenarios."""
        # Input validation for key parameters
        if not all(k in inputs for k in ['project_cost', 'own_capital', 'loan_rate', 'loan_tenure', 'tax_rate', 'investment_return', 'custom_capital_contribution']):
            st.error("Missing required input parameters for calculation.")
            return {} # Return empty dict or raise error
        
        if inputs['project_cost'] <= 0:
            st.error("Project cost must be greater than zero.")
            return {}
        if inputs['loan_tenure'] <= 0:
            st.error("Loan tenure must be greater than zero.")
            return {}
        if inputs['own_capital'] < 0 or inputs['loan_rate'] < 0 or inputs['investment_return'] < 0 or inputs['tax_rate'] < 0:
            st.error("Financial rates and capital cannot be negative.")
            return {}
        if inputs['custom_capital_contribution'] < 0 or inputs['custom_capital_contribution'] > inputs['own_capital']:
            st.error("Custom capital contribution must be non-negative and not exceed total own capital.")
            return {}

        # --- Scenario 1: Maximum Own Funding ---
        # Description: Use your money first to minimize loan.
        s1_capital_used_directly_lakh = min(inputs['project_cost'], inputs['own_capital'])
        s1_loan_amount_lakh = max(0, inputs['project_cost'] - s1_capital_used_directly_lakh)
        s1_loan_amount_actual = s1_loan_amount_lakh * 100000 # Convert to actual amount for EMI calc

        s1_emi = self.calculate_emi(s1_loan_amount_actual, inputs['loan_rate'], inputs['loan_tenure'])
        s1_total_loan_payment = s1_emi * 12 * inputs['loan_tenure']
        s1_gross_total_interest = s1_total_loan_payment - s1_loan_amount_actual
        
        # Apply tax deductibility for loan interest
        s1_effective_total_interest = s1_gross_total_interest
        if inputs['loan_interest_deductible']:
            s1_effective_total_interest *= (1 - inputs['tax_rate']/100)
        
        s1_gross_total_interest_lakh = s1_gross_total_interest / 100000
        s1_effective_total_interest_lakh = s1_effective_total_interest / 100000

        # Total Project Outlay (Gross Cash Outflow): Capital directly used + Gross loan payments
        s1_total_project_outlay = s1_capital_used_directly_lakh + (s1_total_loan_payment / 100000)

        # Net Effective Cost: Project Cost + Effective Total Interest (no investment gains here)
        s1_net_effective_cost = inputs['project_cost'] + s1_effective_total_interest_lakh

        # --- Scenario 2: Maximum Leverage ---
        # Description: Take a loan for the entire project cost and invest all your available own capital.
        s2_capital_used_directly_lakh = 0.0 # No own capital used directly for project
        s2_capital_invested_lakh = inputs['own_capital']
        s2_loan_amount_lakh = inputs['project_cost']
        s2_loan_amount_actual = s2_loan_amount_lakh * 100000 # Convert to actual amount for EMI calc

        s2_emi = self.calculate_emi(s2_loan_amount_actual, inputs['loan_rate'], inputs['loan_tenure'])
        s2_total_loan_payment = s2_emi * 12 * inputs['loan_tenure']
        s2_gross_total_interest = s2_total_loan_payment - s2_loan_amount_actual
        
        # Apply tax deductibility for loan interest
        s2_effective_total_interest = s2_gross_total_interest
        if inputs['loan_interest_deductible']:
            s2_effective_total_interest *= (1 - inputs['tax_rate']/100)
        
        s2_gross_total_interest_lakh = s2_gross_total_interest / 100000
        s2_effective_total_interest_lakh = s2_effective_total_interest / 100000

        # Investment calculations for Scenario 2
        s2_investment_maturity_value_lakh = self.calculate_investment_growth(
            s2_capital_invested_lakh,
            inputs['investment_return'],
            inputs['loan_tenure'],
            self.investment_options[inputs['investment_type']]['compounding']
        )
        s2_investment_gain_lakh = s2_investment_maturity_value_lakh - s2_capital_invested_lakh
        s2_post_tax_gain_lakh = s2_investment_gain_lakh * (1 - inputs['tax_rate']/100)

        # Total Project Outlay (Gross Cash Outflow): Gross loan payments + Own capital invested (initial outflow)
        s2_total_project_outlay = (s2_total_loan_payment / 100000) + s2_capital_invested_lakh

        # Net Effective Cost: Project Cost + Effective Total Interest - Post-Tax Investment Gains
        s2_net_effective_cost = inputs['project_cost'] + s2_effective_total_interest_lakh - s2_post_tax_gain_lakh

        # --- Scenario 3: Balanced Approach ---
        # Description: Contribute a custom amount of own capital directly, and take a loan for the rest. Invest any remaining own capital.
        s3_capital_used_directly_lakh = inputs['custom_capital_contribution']
        s3_remaining_own_capital_invested_lakh = max(0, inputs['own_capital'] - s3_capital_used_directly_lakh)
        s3_loan_amount_lakh = max(0, inputs['project_cost'] - s3_capital_used_directly_lakh)
        s3_loan_amount_actual = s3_loan_amount_lakh * 100000

        s3_emi = self.calculate_emi(s3_loan_amount_actual, inputs['loan_rate'], inputs['loan_tenure'])
        s3_total_loan_payment = s3_emi * 12 * inputs['loan_tenure']
        s3_gross_total_interest = s3_total_loan_payment - s3_loan_amount_actual
        
        # Apply tax deductibility for loan interest
        s3_effective_total_interest = s3_gross_total_interest
        if inputs['loan_interest_deductible']:
            s3_effective_total_interest *= (1 - inputs['tax_rate']/100)
        
        s3_gross_total_interest_lakh = s3_gross_total_interest / 100000
        s3_effective_total_interest_lakh = s3_effective_total_interest / 100000

        # Calculate investment for remaining own capital (if any) for Scenario 3
        s3_investment_maturity_value_lakh = 0
        s3_investment_gain_lakh = 0
        s3_post_tax_gain_lakh = 0

        if s3_remaining_own_capital_invested_lakh > 0:
            s3_investment_maturity_value_lakh = self.calculate_investment_growth(
                s3_remaining_own_capital_invested_lakh,
                inputs['investment_return'],
                inputs['loan_tenure'],
                self.investment_options[inputs['investment_type']]['compounding']
            )
            s3_investment_gain_lakh = s3_investment_maturity_value_lakh - s3_remaining_own_capital_invested_lakh
            s3_post_tax_gain_lakh = s3_investment_gain_lakh * (1 - inputs['tax_rate']/100)

        # Total Project Outlay (Gross Cash Outflow): Custom capital used + Gross loan payments + Remaining capital invested (initial outflow)
        s3_total_project_outlay = s3_capital_used_directly_lakh + (s3_total_loan_payment / 100000) + s3_remaining_own_capital_invested_lakh

        # Net Effective Cost: Project Cost + Effective Total Interest - Post-Tax Investment Gains
        s3_net_effective_cost = inputs['project_cost'] + s3_effective_total_interest_lakh - s3_post_tax_gain_lakh

        # --- Overall Metrics & Recommendation ---
        # Calculate annualized effective rates for display
        effective_loan_rate_annual = inputs['loan_rate']
        if inputs['loan_interest_deductible']:
            effective_loan_rate_annual = inputs['loan_rate'] * (1 - inputs['tax_rate']/100)

        effective_investment_return_annual = 0
        if inputs['own_capital'] > 0 and inputs['loan_tenure'] > 0:
            # Calculate CAGR for Scenario 2's investment (using post-tax maturity value)
            if s2_investment_maturity_value_lakh > 0 and s2_capital_invested_lakh > 0:
                effective_investment_return_annual = ( (s2_investment_maturity_value_lakh / s2_capital_invested_lakh)**(1/inputs['loan_tenure']) - 1 ) * 100

        # Determine recommendation based on Net Effective Cost
        all_net_effective_costs = {
            'scenario1': s1_net_effective_cost,
            'scenario2': s2_net_effective_cost,
            'scenario3': s3_net_effective_cost
        }
        
        min_net_effective_cost_value = min(all_net_effective_costs.values())
        
        recommendation = ''
        for key, value in all_net_effective_costs.items():
            if value == min_net_effective_cost_value:
                recommendation = key
                break # Found the first matching scenario

        max_net_effective_cost_value = max(all_net_effective_costs.values())
        savings_against_worst = max_net_effective_cost_value - min_net_effective_cost_value

        # Calculate prepayment penalty (assuming on initial loan amount of Scenario 2 for illustration)
        prepayment_cost = 0
        if inputs['prepayment_penalty_pct'] > 0 and inputs['loan_tenure'] > 0:
            prepayment_cost = (inputs['project_cost'] * inputs['prepayment_penalty_pct'] / 100)


        # Compile results dictionary
        results = {
            'scenario1': {
                'description': 'Maximum Own Funding: Use your money first to minimize loan.',
                'capital_used_directly': s1_capital_used_directly_lakh,
                'loan_amount': s1_loan_amount_lakh,
                'emi': s1_emi,
                'total_loan_payment': s1_total_loan_payment,
                'gross_total_interest': s1_gross_total_interest_lakh,
                'effective_total_interest': s1_effective_total_interest_lakh,
                'total_project_outlay': s1_total_project_outlay,
                'net_effective_cost': s1_net_effective_cost
            },
            'scenario2': {
                'description': 'Maximum Leverage: Take a loan for the entire project cost and invest all your available own capital.',
                'capital_used_directly': s2_capital_used_directly_lakh,
                'capital_invested': s2_capital_invested_lakh,
                'loan_amount': s2_loan_amount_lakh,
                'emi': s2_emi,
                'total_loan_payment': s2_total_loan_payment,
                'gross_total_interest': s2_gross_total_interest_lakh,
                'effective_total_interest': s2_effective_total_interest_lakh,
                'investment_maturity': s2_investment_maturity_value_lakh,
                'investment_gain': s2_investment_gain_lakh,
                'post_tax_gain': s2_post_tax_gain_lakh,
                'total_project_outlay': s2_total_project_outlay,
                'net_effective_cost': s2_net_effective_cost
            },
            'scenario3': {
                'description': 'Balanced Approach: Contribute a custom amount of own capital directly, and take a loan for the rest. Invest any remaining own capital.',
                'capital_used_directly': s3_capital_used_directly_lakh,
                'remaining_own_capital_invested': s3_remaining_own_capital_invested_lakh,
                'loan_amount': s3_loan_amount_lakh,
                'emi': s3_emi,
                'total_loan_payment': s3_total_loan_payment,
                'gross_total_interest': s3_gross_total_interest_lakh,
                'effective_total_interest': s3_effective_total_interest_lakh,
                'investment_maturity': s3_investment_maturity_value_lakh,
                'investment_gain': s3_investment_gain_lakh,
                'post_tax_gain': s3_post_tax_gain_lakh,
                'total_project_outlay': s3_total_project_outlay,
                'net_effective_cost': s3_net_effective_cost
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
        """Generate recommendation text considering 3 scenarios"""
        interest_spread = results['interest_spread']
        recommendation_scenario = results['recommendation']

        base_text = ""
        if recommendation_scenario == 'scenario1':
            base_text = "💡 **Scenario 1 (Maximum Own Funding)** is recommended."
            if results['scenario1']['loan_amount'] == 0:
                base_text += " You have enough own capital to cover the entire project cost, eliminating the need for a loan."
            else:
                base_text += " This approach minimizes your loan burden by utilizing your own capital first."
            if interest_spread > 3:
                base_text += " The loan interest rate is significantly higher than potential investment returns, making direct funding more economical."
            else:
                base_text += " It offers better capital preservation with the lowest overall effective cost."

        elif recommendation_scenario == 'scenario2':
            base_text = "💡 **Scenario 2 (Maximum Leverage)** is recommended."
            base_text += " This strategy allows you to keep your own capital liquid and invested, potentially generating significant returns."
            if interest_spread < -2:
                base_text += " Your investment returns are substantially higher than your loan costs, leading to a net positive arbitrage."
            else:
                base_text += " It helps maintain maximum liquidity while still proving to be the most cost-effective option."

        else: # recommendation_scenario == 'scenario3'
            base_text = "💡 **Scenario 3 (Balanced Approach)** is recommended."
            base_text += f" This option proposes using ₹{results['scenario3']['capital_used_directly']:.1f}L of your capital directly for the project, and investing the remaining ₹{results['scenario3']['remaining_own_capital_invested']:.1f}L."
            if results['scenario3']['loan_amount'] == 0:
                 base_text += " You can fully fund the project with your custom contribution, eliminating the loan, and still invest your remaining capital."
            else:
                base_text += " It offers the lowest effective cost by striking a balance between direct capital use and strategic investment of remaining funds."
        
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
        st.write(f"**Potential Savings (compared to worst scenario):** ₹{results['savings']:.2f} lakh")
        st.markdown("---")

        # Tabular Comparison of Scenarios
        st.markdown("#### 📈 Financing Scenarios Comparison")
        comparison_data = {
            'Metric': [
                'Scenario Description',
                'Capital Used Directly for Project',
                'Own Capital Invested (Initially)',
                'Loan Amount Taken',
                'Monthly EMI',
                'Gross Total Interest Paid',
                'Effective Total Interest (After Tax Benefits)',
                'Investment Maturity Value (from invested capital)',
                'Gross Investment Gain (from invested capital)',
                'Post-Tax Investment Gain (from invested capital)',
                '**Total Project Outlay (Gross Cash Outflow)**', 
                '**NET EFFECTIVE COST**'
            ],
            'Scenario 1: Maximum Own Funding': [
                results['scenario1']['description'],
                f"₹{results['scenario1']['capital_used_directly']:.1f} lakh",
                "₹0.0 lakh", # No own capital is explicitly "invested" separately
                f"₹{results['scenario1']['loan_amount']:.1f} lakh",
                f"₹{results['scenario1']['emi']:,.0f}",
                f"₹{results['scenario1']['gross_total_interest']:.1f} lakh",
                f"₹{results['scenario1']['effective_total_interest']:.1f} lakh",
                "₹0.0 lakh",
                "₹0.0 lakh",
                "₹0.0 lakh",
                f"₹{results['scenario1']['total_project_outlay']:.2f} lakh", 
                f"₹{results['scenario1']['net_effective_cost']:.2f} lakh"
            ],
            'Scenario 2: Maximum Leverage': [
                results['scenario2']['description'],
                f"₹{results['scenario2']['capital_used_directly']:.1f} lakh",
                f"₹{results['scenario2']['capital_invested']:.1f} lakh",
                f"₹{results['scenario2']['loan_amount']:.1f} lakh",
                f"₹{results['scenario2']['emi']:,.0f}",
                f"₹{results['scenario2']['gross_total_interest']:.1f} lakh",
                f"₹{results['scenario2']['effective_total_interest']:.1f} lakh",
                f"₹{results['scenario2']['investment_maturity']:.2f} lakh",
                f"₹{results['scenario2']['investment_gain']:.2f} lakh",
                f"₹{results['scenario2']['post_tax_gain']:.2f} lakh",
                f"₹{results['scenario2']['total_project_outlay']:.2f} lakh", 
                f"₹{results['scenario2']['net_effective_cost']:.2f} lakh"
            ],
            'Scenario 3: Balanced Approach': [
                results['scenario3']['description'],
                f"₹{results['scenario3']['capital_used_directly']:.1f} lakh",
                f"₹{results['scenario3']['remaining_own_capital_invested']:.1f} lakh",
                f"₹{results['scenario3']['loan_amount']:.1f} lakh",
                f"₹{results['scenario3']['emi']:,.0f}",
                f"₹{results['scenario3']['gross_total_interest']:.1f} lakh",
                f"₹{results['scenario3']['effective_total_interest']:.1f} lakh",
                f"₹{results['scenario3']['investment_maturity']:.2f} lakh" if results['scenario3']['remaining_own_capital_invested'] > 0 else "₹0.0 lakh",
                f"₹{results['scenario3']['investment_gain']:.2f} lakh" if results['scenario3']['remaining_own_capital_invested'] > 0 else "₹0.0 lakh",
                f"₹{results['scenario3']['post_tax_gain']:.2f} lakh" if results['scenario3']['remaining_own_capital_invested'] > 0 else "₹0.0 lakh",
                f"₹{results['scenario3']['total_project_outlay']:.2f} lakh",
                f"₹{results['scenario3']['net_effective_cost']:.2f} lakh"
            ]
        }
        comparison_df = pd.DataFrame(comparison_data).set_index('Metric')
        st.dataframe(comparison_df)

        # Definitions Expander
        with st.expander("❓ Understanding the Key Metrics: Total Project Outlay & Net Effective Cost"): 
            st.markdown("""
            Here's a breakdown of the core financial metrics used in this analysis:
            
            **1. Total Project Outlay (Gross Cash Outflow):**
            This represents the **total amount of money that leaves your hands** (or your business's accounts) throughout the project's financing period. It's the sum of:
            * **Own Capital Used Directly for Project:** Any of your own funds that you directly put into the project to reduce the loan amount.
            * **Total Loan Payments (Principal + Gross Interest):** All EMI payments made to the bank over the loan tenure.
            * **Own Capital Invested (Initially):** If you choose to invest your capital rather than use it directly for the project, this represents the initial lump sum investment.
            
            Think of it as the sum of all 'debit' entries related to acquiring the project and managing your capital. It does *not* account for tax benefits or investment gains that might offset these outflows.
            
            ---
            
            **2. Net Effective Cost (True Economic Cost):**
            This is the **ultimate financial burden** of funding the project, reflecting what the project *truly costs you* after accounting for all direct expenses, tax advantages, and investment benefits. It's calculated as:
            `Net Effective Cost = Total Project Cost + Effective Total Loan Interest (After Tax) - Post-Tax Investment Gains`
            
            Let's break that down:
            * **Total Project Cost:** The fundamental purchase or development cost of the project itself.
            * **Effective Total Loan Interest (After Tax):** This is the total gross interest you pay on your loan, reduced by any tax savings if the loan interest is tax-deductible (based on your input Tax Rate). This provides the *actual* cost of borrowing.
            * **Post-Tax Investment Gains:** If you invest your own capital, any profits earned from that investment (after taxes) are subtracted from your total cost. These gains effectively reduce the overall financial impact of the project on your net worth.
            
            This metric provides the most accurate "bottom-line" figure for each financing scenario, allowing for a direct and comparable evaluation of the real economic impact.
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
            st.write(f"**Custom Capital Contribution (Scenario 3):** ₹{inputs['custom_capital_contribution']:.1f} lakh")
        else:
            st.write(f"**Custom Capital Contribution (Scenario 3):** {inputs['custom_capital_percentage']:.1f}% of Own Capital (₹{inputs['custom_capital_contribution']:.1f} lakh)")


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
            st.write(f" - **Prepayment Cost:** A {inputs['prepayment_penalty_pct']:.2f}% penalty on a full loan (Scenario 2) would be **₹{results['prepayment_cost']:.2f} lakh**. Factor this into early exit scenarios and loan terms.")
            st.write(" - **Investment Volatility:** Assumed investment returns are estimates. Stress test your plan with $\pm 1-2\%$ returns to understand the impact on net effective cost and liquidity.")
            st.write(f" - **Liquidity Buffer:** Your target minimum liquidity is **₹{inputs['min_liquidity_target']:.1f} lakh**. Ensure the chosen option consistently maintains this, especially under stressed cash flow scenarios (e.g., delayed project revenue, unexpected expenses).")
        
        with st.expander("Nuanced Strategic Considerations"):
            if inputs['loan_interest_deductible']:
                st.write(" - **Tax Planning:** Loan interest is considered tax-deductible, significantly reducing the effective cost of borrowing. Ensure proper documentation for claiming this deduction.")
            else:
                st.write(" - **Tax Planning:** Loan interest is NOT considered tax-deductible, meaning the gross interest is the effective cost. Explore other tax optimization strategies.")
            st.write(" - **Credit Profile & Leverage:** Assess the impact of increased debt on your company's debt-to-equity ratio, credit rating, and future borrowing capacity. Excessive leverage can affect banking relationships and future funding flexibility.")
            st.write(" - **Regulatory Compliance:** Confirm all documentation (project reports, audited financials) are ready for smooth loan disbursal and subsequent annual reviews to avoid penalties or delays.")
            st.write(" - **Investment Liquidity/Market Risk:** Even 'liquid' investments carry some market risk (e.g., temporary impairment during market freezes). Maintain an emergency buffer in a highly liquid bank account (beyond investment) for immediate needs.")
            st.write(" - **Optimization Moves & Additional Value Levers:** Consider phased loan drawdown or capital deployment strategies to minimize idle funds and optimize interest costs or investment gains.")
        st.markdown("---")

    def generate_pdf_report(self, inputs, results, year_wise_df):
        """Generates a PDF report using ReportLab."""
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=letter)
        styles = getSampleStyleSheet()
        story = []

        # Title
        story.append(Paragraph("Project Financing Analysis Report", styles['h1']))
        story.append(Spacer(1, 0.2 * inch))

        # Date of Report
        story.append(Paragraph(f"Date: {pd.Timestamp.now().strftime('%Y-%m-%d')}", styles['Normal']))
        story.append(Spacer(1, 0.2 * inch))

        # Recommendation
        story.append(Paragraph("<b>Recommendation:</b>", styles['h2']))
        recommendation_text = self.get_recommendation_text(results, inputs)
        story.append(Paragraph(recommendation_text, styles['Normal']))
        story.append(Paragraph(f"Potential Savings (compared to worst scenario): ₹{results['savings']:.2f} lakh", styles['Normal']))
        story.append(Spacer(1, 0.2 * inch))

        # Input Parameters
        story.append(Paragraph("<b>Input Parameters Used:</b>", styles['h2']))
        input_data = [
            ['Metric', 'Value'],
            ['Total Project Cost', f"₹{inputs['project_cost']:.1f} lakh"],
            ['Own Capital Available', f"₹{inputs['own_capital']:.1f} lakh"],
            ['Bank Loan Interest Rate', f"{inputs['loan_rate']:.2f}% p.a. ({inputs['loan_type']})"],
            ['Loan Tenure', f"{inputs['loan_tenure']} years"],
            ['Loan Interest Tax Deductible', 'Yes' if inputs['loan_interest_deductible'] else 'No'],
            ['Prepayment Penalty', f"{inputs['prepayment_penalty_pct']:.2f}%"],
            ['Minimum Liquidity Target', f"₹{inputs['min_liquidity_target']:.1f} lakh"],
            ['Investment Type', inputs['investment_type']],
            ['Investment Return', f"{inputs['investment_return']:.2f}% p.a."],
            ['Tax Rate', f"{inputs['tax_rate']:.0f}%"]
        ]
        if inputs['custom_capital_input_type'] == 'Value':
            input_data.append(['Custom Capital Contribution (Scenario 3)', f"₹{inputs['custom_capital_contribution']:.1f} lakh"])
        else:
            input_data.append(['Custom Capital Contribution (Scenario 3)', f"{inputs['custom_capital_percentage']:.1f}% of Own Capital (₹{inputs['custom_capital_contribution']:.1f} lakh)"])
        
        input_table = Table(input_data, colWidths=[2.5*inch, 3*inch])
        input_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#D3D3D3')),
            ('GRID', (0,0), (-1,-1), 1, colors.black),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0,0), (-1,0), 6),
            ('BACKGROUND', (0,1), (-1,-1), colors.white),
            ('FONTSIZE', (0,0), (-1,-1), 10)
        ]))
        story.append(input_table)
        story.append(Spacer(1, 0.2 * inch))

        # Investment Details
        investment_details = self.investment_options[inputs['investment_type']]
        story.append(Paragraph("<b>Investment Details:</b>", styles['h3']))
        story.append(Paragraph(f" - <b>Liquidity:</b> {investment_details['liquidity']}", styles['Normal']))
        story.append(Paragraph(f" - <b>Tax Efficiency:</b> {investment_details['tax_efficiency']}", styles['Normal']))
        story.append(Paragraph(f" - <b>Compounding:</b> {investment_details['compounding']}", styles['Normal']))
        story.append(Paragraph(f" - <b>Notes:</b> {investment_details['notes']}", styles['Normal']))
        story.append(Spacer(1, 0.2 * inch))

        # Financing Scenarios Comparison
        story.append(Paragraph("<b>Financing Scenarios Comparison:</b>", styles['h2']))
        
        # Prepare data for PDF comparison table, ensuring consistent formatting
        pdf_comparison_values = {
            'Scenario 1: Maximum Own Funding': [
                results['scenario1']['description'],
                f"₹{results['scenario1']['capital_used_directly']:.1f} lakh",
                "₹0.0 lakh", 
                f"₹{results['scenario1']['loan_amount']:.1f} lakh",
                f"₹{results['scenario1']['emi']:,.0f}",
                f"₹{results['scenario1']['gross_total_interest']:.1f} lakh",
                f"₹{results['scenario1']['effective_total_interest']:.1f} lakh",
                "₹0.0 lakh",
                "₹0.0 lakh",
                "₹0.0 lakh",
                f"₹{results['scenario1']['total_project_outlay']:.2f} lakh", 
                f"₹{results['scenario1']['net_effective_cost']:.2f} lakh"
            ],
            'Scenario 2: Maximum Leverage': [
                results['scenario2']['description'],
                f"₹{results['scenario2']['capital_used_directly']:.1f} lakh",
                f"₹{results['scenario2']['capital_invested']:.1f} lakh",
                f"₹{results['scenario2']['loan_amount']:.1f} lakh",
                f"₹{results['scenario2']['emi']:,.0f}",
                f"₹{results['scenario2']['gross_total_interest']:.1f} lakh",
                f"₹{results['scenario2']['effective_total_interest']:.1f} lakh",
                f"₹{results['scenario2']['investment_maturity']:.2f} lakh",
                f"₹{results['scenario2']['investment_gain']:.2f} lakh",
                f"₹{results['scenario2']['post_tax_gain']:.2f} lakh",
                f"₹{results['scenario2']['total_project_outlay']:.2f} lakh", 
                f"₹{results['scenario2']['net_effective_cost']:.2f} lakh"
            ],
            'Scenario 3: Balanced Approach': [
                results['scenario3']['description'],
                f"₹{results['scenario3']['capital_used_directly']:.1f} lakh",
                f"₹{results['scenario3']['remaining_own_capital_invested']:.1f} lakh",
                f"₹{results['scenario3']['loan_amount']:.1f} lakh",
                f"₹{results['scenario3']['emi']:,.0f}",
                f"₹{results['scenario3']['gross_total_interest']:.1f} lakh",
                f"₹{results['scenario3']['effective_total_interest']:.1f} lakh",
                f"₹{results['scenario3']['investment_maturity']:.2f} lakh" if results['scenario3']['remaining_own_capital_invested'] > 0 else "₹0.0 lakh",
                f"₹{results['scenario3']['investment_gain']:.2f} lakh" if results['scenario3']['remaining_own_capital_invested'] > 0 else "₹0.0 lakh",
                f"₹{results['scenario3']['post_tax_gain']:.2f} lakh" if results['scenario3']['remaining_own_capital_invested'] > 0 else "₹0.0 lakh",
                f"₹{results['scenario3']['total_project_outlay']:.2f} lakh",
                f"₹{results['scenario3']['net_effective_cost']:.2f} lakh"
            ]
        }
        
        # Build the comparison table for PDF
        comparison_data_for_pdf = [
            ['Metric', 'Scenario 1: Maximum Own Funding', 'Scenario 2: Maximum Leverage', 'Scenario 3: Balanced Approach']
        ]
        
        # Match metrics to the dynamically generated values
        metrics_list = [
            'Scenario Description',
            'Capital Used Directly for Project',
            'Own Capital Invested (Initially)',
            'Loan Amount Taken',
            'Monthly EMI',
            'Gross Total Interest Paid',
            'Effective Total Interest (After Tax Benefits)',
            'Investment Maturity Value (from invested capital)',
            'Gross Investment Gain (from invested capital)',
            'Post-Tax Investment Gain (from invested capital)',
            'Total Project Outlay (Gross Cash Outflow)', 
            'NET EFFECTIVE COST'
        ]

        for i, metric in enumerate(metrics_list):
            row = [metric]
            row.append(pdf_comparison_values['Scenario 1: Maximum Own Funding'][i])
            row.append(pdf_comparison_values['Scenario 2: Maximum Leverage'][i])
            row.append(pdf_comparison_values['Scenario 3: Balanced Approach'][i])
            comparison_data_for_pdf.append(row)

        comparison_table = Table(comparison_data_for_pdf, colWidths=[2.1*inch, 1.8*inch, 1.8*inch, 1.8*inch])
        comparison_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#D3D3D3')),
            ('GRID', (0,0), (-1,-1), 1, colors.black),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('ALIGN', (0,0), (0,-1), 'LEFT'), # Left align metric column
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0,0), (-1,0), 6),
            ('BACKGROUND', (0,1), (-1,-1), colors.white),
            ('FONTSIZE', (0,0), (-1,-1), 7) # Smaller font for more compact table
        ]))
        story.append(comparison_table)
        story.append(Spacer(1, 0.2 * inch))
        story.append(Paragraph("<i>Detailed explanations of 'Total Project Outlay' and 'Net Effective Cost' are available in the web application.</i>", styles['Italic']))
        story.append(PageBreak()) # Start Year-wise data on a new page


        # Year-wise Data
        story.append(Paragraph("<b>Year-wise Financial Data:</b>", styles['h2']))
        story.append(Spacer(1, 0.1 * inch))

        # Convert DataFrame to list of lists for ReportLab table, including headers
        year_wise_data_for_pdf = [year_wise_df.columns.tolist()] + year_wise_df.values.tolist()
        
        # Apply formatting to numeric columns in year_wise_data_for_pdf
        for r_idx, row in enumerate(year_wise_data_for_pdf):
            if r_idx == 0: # Skip header row
                continue
            for c_idx, value in enumerate(row):
                if c_idx > 0: # All columns except 'Year' are numeric Lakhs
                    year_wise_data_for_pdf[r_idx][c_idx] = f"{value:.2f}" # Format to 2 decimal places

        # Set column widths to try and fit on page
        num_cols = len(year_wise_data_for_pdf[0])
        # Distribute widths based on the number of columns and page size
        col_widths = [letter[0] / num_cols - (0.5 * inch / num_cols)] * num_cols # Basic distribution
        # Adjust specific columns if known to be wider/narrower
        col_widths[0] = 0.5 * inch # Year column
        
        try:
            year_wise_table = Table(year_wise_data_for_pdf, colWidths=col_widths)
            year_wise_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#D3D3D3')),
                ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0,0), (-1,0), 4),
                ('BACKGROUND', (0,1), (-1,-1), colors.white),
                ('FONTSIZE', (0,0), (-1,-1), 6) # Very small font to fit many columns
            ]))
            story.append(year_wise_table)
        except Exception as e:
            story.append(Paragraph(f"<i>Could not generate Year-wise table due to: {e}. It might be too wide for the page.</i>", styles['Italic']))


        doc.build(story)
        buffer.seek(0)
        return buffer
