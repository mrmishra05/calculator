import math
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import streamlit as st
import io

# For PDF generation
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT

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
            'S1_Cumulative_Gross_Interest (Lakh)',
            'S2_Cumulative_Gross_Interest (Lakh)',
            'S3_Cumulative_Gross_Interest (Lakh)',
            'S2_Investment_Value (Lakh)',
            'S3_Investment_Value (Lakh)',
            'S2_Investment_Gain (Lakh)',
            'S3_Investment_Gain (Lakh)',
            'S2_Post_Tax_Investment_Gain (Lakh)',
            'S3_Post_Tax_Investment_Gain (Lakh)',
            'S1_Cumulative_Net_Effective_Cost (Lakh)',
            'S2_Cumulative_Net_Effective_Cost (Lakh)',
            'S3_Cumulative_Net_Effective_Cost (Lakh)'
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
                s1_gross_interest_cumulative = 0
                s1_net_effective_cost_cumulative = inputs['project_cost'] # Only project cost at year 0
            else:
                s1_total_paid_cumulative = scenario1_emi * 12 * year
                s1_principal_paid_cumulative = min(s1_total_paid_cumulative, scenario1_loan_amount_actual)
                s1_gross_interest_cumulative = max(0, s1_total_paid_cumulative - s1_principal_paid_cumulative) / 100000 # Convert to lakh
                
                s1_effective_interest_cumulative = s1_gross_interest_cumulative
                if inputs['loan_interest_deductible']:
                    s1_effective_interest_cumulative *= (1 - inputs['tax_rate']/100)
                s1_net_effective_cost_cumulative = inputs['project_cost'] + s1_effective_interest_cumulative


            # Scenario 2: Maximum Leverage
            if year == 0:
                s2_gross_interest_cumulative = 0
                s2_investment_value = own_capital_lakh # Initial investment value
                s2_investment_gain = 0
                s2_post_tax_gain = 0
                s2_net_effective_cost_cumulative = inputs['project_cost'] # Only project cost at year 0
            else:
                s2_total_paid_cumulative = scenario2_emi * 12 * year
                s2_principal_paid_cumulative = min(s2_total_paid_cumulative, scenario2_loan_amount_actual)
                s2_gross_interest_cumulative = max(0, s2_total_paid_cumulative - s2_principal_paid_cumulative) / 100000 # Convert to lakh

                s2_effective_interest_cumulative = s2_gross_interest_cumulative
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
                s3_gross_interest_cumulative = 0
                s3_investment_value = scenario3_remaining_own_capital_lakh if scenario3_remaining_own_capital_lakh > 0 else 0
                s3_investment_gain = 0
                s3_post_tax_gain = 0
                s3_net_effective_cost_cumulative = inputs['project_cost'] # Only project cost at year 0
            else:
                s3_total_paid_cumulative = scenario3_emi * 12 * year
                s3_principal_paid_cumulative = min(s3_total_paid_cumulative, scenario3_loan_amount_actual)
                s3_gross_interest_cumulative = max(0, s3_total_paid_cumulative - s3_principal_paid_cumulative) / 100000 # Convert to lakh

                s3_effective_interest_cumulative = s3_gross_interest_cumulative
                if inputs['loan_interest_deductible']:
                    s3_effective_interest_cumulative *= (1 - inputs['tax_rate']/100)

                s3_investment_value = scenario3_remaining_own_capital_lakh # Initial value if no growth
                if scenario3_remaining_own_capital_invested_lakh > 0: # Ensure there was capital to invest
                    s3_investment_value = self.calculate_investment_growth(
                        scenario3_remaining_own_capital_invested_lakh,
                        inputs['investment_return'],
                        year,
                        self.investment_options[inputs['investment_type']]['compounding']
                    )
                s3_investment_gain = s3_investment_value - scenario3_remaining_own_capital_lakh
                s3_post_tax_gain = s3_investment_gain * (1 - inputs['tax_rate']/100)
                
                s3_net_effective_cost_cumulative = inputs['project_cost'] + s3_effective_interest_cumulative - s3_post_tax_gain


            data.append([
                year,
                s1_gross_interest_cumulative,
                s2_gross_interest_cumulative,
                s3_gross_interest_cumulative,
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
        # If the loan is fully prepaid immediately, it's on the full loan amount
        # For a more accurate calculation, one would need to know the outstanding principal at time of prepayment
        # For simplicity here, we'll calculate based on the initial project cost
        if inputs['prepayment_penalty_pct'] > 0:
            prepayment_cost = (inputs['project_cost'] * inputs['prepayment_penalty_pct'] / 100)


        # Compile results dictionary
        results = {
            'scenario1': {
                'description': 'Maximum Own Funding: Use your money first to minimize loan.',
                'capital_used_directly': s1_capital_used_directly_lakh,
                'loan_amount': s1_loan_amount_lakh,
                'emi': s1_emi,
                'total_gross_loan_payments': s1_total_loan_payment, # Keep this for internal calculation if needed, but not displayed as a primary metric
                'gross_total_interest': s1_gross_total_interest_lakh,
                'effective_total_interest': s1_effective_total_interest_lakh,
                'net_effective_cost': s1_net_effective_cost
            },
            'scenario2': {
                'description': 'Maximum Leverage: Take a loan for the entire project cost and invest all your available own capital.',
                'capital_used_directly': s2_capital_used_directly_lakh,
                'capital_invested': s2_capital_invested_lakh,
                'loan_amount': s2_loan_amount_lakh,
                'emi': s2_emi,
                'total_gross_loan_payments': s2_total_loan_payment,
                'gross_total_interest': s2_gross_total_interest_lakh,
                'effective_total_interest': s2_effective_total_interest_lakh,
                'investment_maturity': s2_investment_maturity_value_lakh,
                'investment_gain': s2_investment_gain_lakh,
                'post_tax_gain': s2_post_tax_gain_lakh,
                'net_effective_cost': s2_net_effective_cost
            },
            'scenario3': {
                'description': 'Balanced Approach: Contribute a custom amount of own capital directly, and take a loan for the rest. Invest any remaining own capital.',
                'capital_used_directly': s3_capital_used_directly_lakh,
                'remaining_own_capital_invested': s3_remaining_own_capital_invested_lakh,
                'loan_amount': s3_loan_amount_lakh,
                'emi': s3_emi,
                'total_gross_loan_payments': s3_total_loan_payment,
                'gross_total_interest': s3_gross_total_interest_lakh,
                'effective_total_interest': s3_effective_total_interest_lakh,
                'investment_maturity': s3_investment_maturity_value_lakh,
                'investment_gain': s3_investment_gain_lakh,
                'post_tax_gain': s3_post_tax_gain_lakh,
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
                'Own Capital Used for Project',
                'Own Capital Available & Invested',
                'Loan Amount Required',
                'Monthly EMI',
                'Gross Total Interest Paid (Loan)',
                'Effective Total Interest (After Tax Benefits)',
                'Total Investment Maturity Value',
                'Total Post-Tax Investment Gain',
                '**NET EFFECTIVE COST**'
            ],
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
                f"₹{results['scenario2']['post_tax_gain']:.2f} lakh",
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
                f"₹{results['scenario3']['post_tax_gain']:.2f} lakh" if results['scenario3']['remaining_own_capital_invested'] > 0 else "₹0.0 lakh",
                f"₹{results['scenario3']['net_effective_cost']:.2f} lakh"
            ]
        }
        comparison_df = pd.DataFrame(comparison_data).set_index('Metric')
        st.dataframe(comparison_df)

        # Definitions Expander (Updated)
        with st.expander("❓ Understanding the Key Metric: Net Effective Cost"): 
            st.markdown("""
            Here's a breakdown of the core financial metric used in this analysis:
            
            **Net Effective Cost (True Economic Cost):**
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
        if inputs['custom_capital_input_type'] == 'Value (Lakhs)': # Match the radio button label exactly
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
        doc = SimpleDocTemplate(buffer, pagesize=A4,
                                rightMargin=0.5*inch, leftMargin=0.5*inch,
                                topMargin=0.5*inch, bottomMargin=0.5*inch)
        styles = getSampleStyleSheet()
        # Custom style for small font tables
        styles.add(ParagraphStyle(name='TableCaption', fontSize=8, alignment=TA_CENTER))
        styles.add(ParagraphStyle(name='SmallTableText', fontSize=6, alignment=TA_CENTER))
        styles.add(ParagraphStyle(name='SmallTableTextLeft', fontSize=6, alignment=TA_LEFT))

        story = []

        # ... (rest of the PDF generation code) ...

        # This part needs to be correctly indented:
        try:
            year_wise_table = Table(year_wise_data_for_pdf, colWidths=col_widths_year_wise)
            year_wise_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#D3D3D3')),
                ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('VALIGN', (0,0), (-1,-1), '-1,-1', 'MIDDLE'),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0,0), (-1,0), 4),
                ('BACKGROUND', (0,1), (-1,-1), colors.white),
                ('FONTSIZE', (0,0), (-1,-1), 5) # VERY small font to try and fit
            ]))
            story.append(year_wise_table)
        except Exception as e:
            # Fallback for tables too wide to fit
            story.append(Paragraph(f"<font color='red'><i>Could not generate Year-wise table in PDF due to layout constraints: {e}. "
                                   f"It contains {num_cols} columns and might be too wide for the page. "
                                   f"Please download the Excel/CSV for full year-wise data.</i></font>", styles['Normal']))

        # Ensure these two lines are correctly indented within the method
        doc.build(story)
        buffer.seek(0)
        return buffer # <--- THIS LINE MUST BE INDENTED TO BE PART OF THE METHOD

def main():
    st.set_page_config(layout="wide", page_title="Project Financing Calculator")
    calculator = ProjectFinancingCalculator()
    styles = getSampleStyleSheet() # Needed for custom style below
    
    # Custom ParagraphStyle for PDF
    from reportlab.lib.styles import ParagraphStyle # Import ParagraphStyle
    styles.add(ParagraphStyle(name='Italic', fontName='Helvetica-Oblique', fontSize=9, alignment=TA_LEFT))


    st.title("💰 Project Financing Strategy Calculator")
    st.markdown("""
        Understand the true cost of funding your project by comparing different financing strategies:
        * **Scenario 1: Maximum Own Funding** - Use your available capital first, minimizing the loan.
        * **Scenario 2: Maximum Leverage** - Take a full loan and invest all your own capital elsewhere.
        * **Scenario 3: Balanced Approach** - A custom mix of own capital contribution and loan.
    """)
    st.markdown("---")

    # --- Input Section ---
    st.header("1. Input Project & Financial Details")

    col1, col2 = st.columns(2)
    with col1:
        project_cost = st.number_input("Total Project Cost (in Lakhs ₹)", min_value=1.0, value=100.0, step=5.0)
        own_capital = st.number_input("Your Own Capital Available (in Lakhs ₹)", min_value=0.0, value=50.0, step=5.0)
        
        # Custom Capital Contribution for Scenario 3
        st.subheader("Scenario 3: Custom Capital Contribution")
        custom_capital_input_type = st.radio("How would you like to define custom capital for Scenario 3?", 
                                             ('Value (Lakhs)', 'Percentage of Own Capital (%)'), 
                                             index=0)
        
        custom_capital_contribution = 0.0
        if custom_capital_input_type == 'Value (Lakhs)':
            custom_capital_contribution = st.number_input("Custom Capital Contribution (in Lakhs ₹)", 
                                                            min_value=0.0, 
                                                            max_value=own_capital, 
                                                            value=min(25.0, own_capital), 
                                                            step=1.0)
        else: # Percentage
            custom_capital_percentage = st.slider("Custom Capital Contribution (% of Own Capital)", 
                                                  min_value=0, max_value=100, value=50, step=1)
            custom_capital_contribution = (own_capital * custom_capital_percentage) / 100.0
            st.info(f"This translates to: ₹{custom_capital_contribution:.1f} Lakhs")

    with col2:
        loan_rate = st.number_input("Bank Loan Interest Rate (% p.a.)", min_value=0.1, value=9.0, step=0.1)
        loan_type = st.radio("Loan Type", ('Floating', 'Fixed'))
        loan_tenure = st.number_input("Loan Tenure (Years)", min_value=1, value=10, step=1)
        loan_interest_deductible = st.checkbox("Loan Interest Tax-Deductible (for business expenses)", value=True)
        prepayment_penalty_pct = st.number_input("Prepayment Penalty (% of outstanding loan)", min_value=0.0, value=0.0, step=0.1)
        min_liquidity_target = st.number_input("Minimum Liquidity Target (in Lakhs ₹)", min_value=0.0, value=10.0, step=1.0)

    st.markdown("---")
    st.header("2. Input Investment Details")
    col3, col4 = st.columns(2)
    with col3:
        investment_types = list(calculator.investment_options.keys())
        default_investment_index = investment_types.index('Liquid Funds') if 'Liquid Funds' in investment_types else 0
        investment_type = st.selectbox("Select Investment Type for Own Capital", investment_types, index=default_investment_index)
        
        # Display default return for selected type, allow override
        default_inv_return = calculator.investment_options[investment_type]['default_return']
        investment_return = st.number_input(f"Expected Investment Return (% p.a. - for {investment_type})", min_value=0.0, value=default_inv_return, step=0.1)
        
    with col4:
        tax_rate = st.slider("Your Applicable Tax Rate (%)", min_value=0, max_value=50, value=30, step=1)

    st.markdown("---")

    # Prepare inputs dictionary
    inputs = {
        'project_cost': project_cost,
        'own_capital': own_capital,
        'loan_rate': loan_rate,
        'loan_type': loan_type,
        'loan_tenure': loan_tenure,
        'loan_interest_deductible': loan_interest_deductible,
        'prepayment_penalty_pct': prepayment_penalty_pct,
        'min_liquidity_target': min_liquidity_target,
        'investment_type': investment_type,
        'investment_return': investment_return,
        'tax_rate': tax_rate,
        'custom_capital_input_type': custom_capital_input_type,
        'custom_capital_percentage': custom_capital_percentage if custom_capital_input_type == 'Percentage of Own Capital (%)' else 0,
        'custom_capital_contribution': custom_capital_contribution
    }

    # --- Calculation & Display Section ---
    st.header("3. Results")
    
    if st.button("Calculate Financing Scenarios"):
        # Perform calculations
        results = calculator.calculate_comparison(inputs)

        if results: # Only proceed if calculations were successful
            # Print detailed report
            calculator.print_detailed_report(inputs, results)
            
            st.markdown("---")
            st.header("4. Visualizations")

            # --- Plotting ---
            year_wise_df = calculator.generate_year_wise_data(inputs, results)
            
            col_plot1, col_plot2 = st.columns(2)

            with col_plot1:
                st.subheader("Net Effective Cost Over Years")
                fig_cost, ax_cost = plt.subplots(figsize=(10, 6))
                ax_cost.plot(year_wise_df['Year'], year_wise_df['S1_Cumulative_Net_Effective_Cost (Lakh)'], label='Scenario 1', marker='o')
                ax_cost.plot(year_wise_df['Year'], year_wise_df['S2_Cumulative_Net_Effective_Cost (Lakh)'], label='Scenario 2', marker='x')
                ax_cost.plot(year_wise_df['Year'], year_wise_df['S3_Cumulative_Net_Effective_Cost (Lakh)'], label='Scenario 3', marker='s')
                ax_cost.set_xlabel("Year")
                ax_cost.set_ylabel("Cumulative Net Effective Cost (Lakh ₹)")
                ax_cost.set_title("Cumulative Net Effective Cost for Each Scenario")
                ax_cost.legend()
                ax_cost.grid(True)
                st.pyplot(fig_cost)

            with col_plot2:
                st.subheader("Cumulative Gross Interest Paid Over Years")
                fig_interest, ax_interest = plt.subplots(figsize=(10, 6))
                ax_interest.plot(year_wise_df['Year'], year_wise_df['S1_Cumulative_Gross_Interest (Lakh)'], label='Scenario 1', marker='o')
                ax_interest.plot(year_wise_df['Year'], year_wise_df['S2_Cumulative_Gross_Interest (Lakh)'], label='Scenario 2', marker='x')
                ax_interest.plot(year_wise_df['Year'], year_wise_df['S3_Cumulative_Gross_Interest (Lakh)'], label='Scenario 3', marker='s')
                ax_interest.set_xlabel("Year")
                ax_interest.set_ylabel("Cumulative Gross Interest Paid (Lakh ₹)")
                ax_interest.set_title("Cumulative Gross Interest Paid for Each Scenario")
                ax_interest.legend()
                ax_interest.grid(True)
                st.pyplot(fig_interest)

            # Bar chart for final Net Effective Cost
            st.subheader("Final Net Effective Cost Comparison")
            final_costs = {
                'Scenario 1': results['scenario1']['net_effective_cost'],
                'Scenario 2': results['scenario2']['net_effective_cost'],
                'Scenario 3': results['scenario3']['net_effective_cost']
            }
            cost_df = pd.DataFrame(final_costs.items(), columns=['Scenario', 'Net Effective Cost (Lakh ₹)'])
            
            fig_bar, ax_bar = plt.subplots(figsize=(10, 6))
            sns.barplot(x='Scenario', y='Net Effective Cost (Lakh ₹)', data=cost_df, ax=ax_bar, palette='viridis')
            ax_bar.set_title("Final Net Effective Cost by Scenario")
            ax_bar.set_ylabel("Net Effective Cost (Lakh ₹)")
            st.pyplot(fig_bar)

            st.markdown("---")
            st.header("5. Download Data")
            
            # Download DataFrame as CSV or Excel
            csv_buffer = io.StringIO()
            year_wise_df.to_csv(csv_buffer, index=False)
            st.download_button(
                label="Download Year-wise Data as CSV",
                data=csv_buffer.getvalue(),
                file_name="project_financing_year_wise_data.csv",
                mime="text/csv",
                key="download_csv"
            )

            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                year_wise_df.to_excel(writer, sheet_name='Year-wise Data', index=False)
                # You can add the summary results to another sheet
                summary_data_for_excel = {
                    "Scenario 1": results["scenario1"],
                    "Scenario 2": results["scenario2"],
                    "Scenario 3": results["scenario3"]
                }
                summary_df = pd.DataFrame.from_dict(summary_data_for_excel, orient='index')
                # Filter out 'description' and 'total_gross_loan_payments' as they are not needed in this summary view
                summary_df = summary_df.drop(columns=['description', 'total_gross_loan_payments'], errors='ignore') 
                summary_df.to_excel(writer, sheet_name='Summary Results')
            
            st.download_button(
                label="Download Full Report as Excel",
                data=excel_buffer.getvalue(),
                file_name="project_financing_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_excel"
            )

            # Download PDF Report
            pdf_buffer = calculator.generate_pdf_report(inputs, results, year_wise_df)
            st.download_button(
                label="Download Summary Report as PDF",
                data=pdf_buffer,
                file_name="project_financing_summary_report.pdf",
                mime="application/pdf",
                key="download_pdf"
            )


if __name__ == "__main__":
    main()
