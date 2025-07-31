import math
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import numpy as np

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
            option1_paid = results['option1']['emi'] * 12 * year if year > 0 else 0
            option1_principal_paid = min(option1_paid, results['option1']['loan_amount'])
            option1_interest_paid = max(0, option1_paid - option1_principal_paid)
            
            # Option 2 calculations
            option2_paid = results['option2']['emi'] * 12 * year if year > 0 else 0
            option2_principal_paid = min(option2_paid, inputs['project_cost'])
            option2_interest_paid = max(0, option2_paid - option2_principal_paid)
            
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
        option1_loan_amount = max(0, inputs['project_cost'] - inputs['own_capital'])
        option1_emi = self.calculate_emi(option1_loan_amount, inputs['loan_rate'], inputs['loan_tenure'])
        option1_total_payment = option1_emi * 12 * inputs['loan_tenure']
        option1_total_interest = option1_total_payment - option1_loan_amount
        option1_net_outflow = inputs['own_capital'] + option1_total_interest
        
        # Option 2: Invest own capital + take full loan
        option2_loan_amount = inputs['project_cost']
        option2_emi = self.calculate_emi(option2_loan_amount, inputs['loan_rate'], inputs['loan_tenure'])
        option2_total_payment = option2_emi * 12 * inputs['loan_tenure']
        option2_total_interest = option2_total_payment - option2_loan_amount
        
        # Investment calculations
        investment_maturity_value = self.calculate_investment_growth(
            inputs['own_capital'],
            inputs['investment_return'],
            inputs['loan_tenure'],
            self.investment_options[inputs['investment_type']]['compounding']
        )
        
        investment_gain = investment_maturity_value - inputs['own_capital']
        post_tax_gain = investment_gain * (1 - inputs['tax_rate']/100)
        option2_net_outflow = option2_total_interest - post_tax_gain
        
        # Determine recommendation
        recommendation = 'option1' if option1_net_outflow < option2_net_outflow else 'option2'
        savings = abs(option1_net_outflow - option2_net_outflow)
        
        results = {
            'option1': {
                'loan_amount': option1_loan_amount,
                'emi': option1_emi,
                'total_payment': option1_total_payment,
                'total_interest': option1_total_interest,
                'net_outflow': option1_net_outflow,
                'capital_used': inputs['own_capital']
            },
            'option2': {
                'loan_amount': option2_loan_amount,
                'emi': option2_emi,
                'total_payment': option2_total_payment,
                'total_interest': option2_total_interest,
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
        """Print comprehensive analysis report"""
        print("=" * 80)
        print("PROJECT FINANCING CALCULATOR - DETAILED ANALYSIS")
        print("=" * 80)
        
        # Input Summary
        print("\n📋 INPUT PARAMETERS:")
        print(f"   Total Project Cost: ₹{inputs['project_cost']:.1f} lakh")
        print(f"   Own Capital Available: ₹{inputs['own_capital']:.1f} lakh")
        print(f"   Bank Loan Interest Rate: {inputs['loan_rate']:.2f}% p.a.")
        print(f"   Loan Tenure: {inputs['loan_tenure']} years")
        print(f"   Investment Type: {inputs['investment_type']}")
        print(f"   Investment Return: {inputs['investment_return']:.2f}% p.a.")
        print(f"   Tax Rate: {inputs['tax_rate']:.0f}%")
        
        investment_details = self.investment_options[inputs['investment_type']]
        print(f"   Investment Details:")
        print(f"     - Liquidity: {investment_details['liquidity']}")
        print(f"     - Tax Efficiency: {investment_details['tax_efficiency']}")
        print(f"     - Compounding: {investment_details['compounding']}")
        
        print("\n" + "="*50)
        
        # Option 1 Analysis
        print(f"\n📊 OPTION 1: Use Own Capital (₹{inputs['own_capital']:.1f}L) + Small Loan")
        print(f"   Loan Amount: ₹{results['option1']['loan_amount']:.1f} lakh")
        if results['option1']['loan_amount'] > 0:
            print(f"   Monthly EMI: ₹{results['option1']['emi']:,.0f}")
            print(f"   Total Payment: ₹{results['option1']['total_payment']/100000:.2f} lakh")
            print(f"   Total Interest: ₹{results['option1']['total_interest']:.1f} lakh")
        print(f"   Capital Used: ₹{results['option1']['capital_used']:.1f} lakh")
        print(f"   NET COST: ₹{results['option1']['net_outflow']:.2f} lakh")
        
        print("\n" + "-"*50)
        
        # Option 2 Analysis
        print(f"\n💰 OPTION 2: Invest Capital + Full Loan (₹{inputs['project_cost']:.1f}L)")
        print(f"   Loan Amount: ₹{results['option2']['loan_amount']:.1f} lakh")
        print(f"   Monthly EMI: ₹{results['option2']['emi']:,.0f}")
        print(f"   Total Payment: ₹{results['option2']['total_payment']/100000:.2f} lakh")
        print(f"   Total Interest: ₹{results['option2']['total_interest']:.1f} lakh")
        print(f"\n   Investment Analysis:")
        print(f"     Initial Investment: ₹{inputs['own_capital']:.1f} lakh")
        print(f"     Maturity Value: ₹{results['option2']['investment_maturity']:.2f} lakh")
        print(f"     Gross Gain: ₹{results['option2']['investment_gain']:.2f} lakh")
        print(f"     Post-Tax Gain: ₹{results['option2']['post_tax_gain']:.2f} lakh")
        print(f"   NET COST: ₹{results['option2']['net_outflow']:.2f} lakh")
        
        print("\n" + "="*50)
        
        # Final Recommendation
        print(f"\n✅ RECOMMENDATION:")
        recommendation_text = self.get_recommendation_text(results, inputs)
        print(f"   {recommendation_text}")
        print(f"   Potential Savings: ₹{results['savings']:.2f} lakh")
        
        # Additional Insights
        print(f"\n🔍 KEY INSIGHTS:")
        print(f"   Interest Rate Spread: {results['interest_spread']:.2f}% (Loan - Investment)")
        
        if results['recommendation'] == 'option1':
            print(f"   ✓ Lower total cost")
            print(f"   ⚠ Capital gets locked in project")
            print(f"   ⚠ Reduced liquidity")
        else:
            print(f"   ✓ Maintains full liquidity of ₹{inputs['own_capital']:.1f}L")
            print(f"   ✓ Emergency funds available")
            print(f"   ✓ Opportunity for better investments")
            print(f"   ⚠ Higher total interest outgo")
        
        print("\n" + "="*80)
    
    def create_visualizations(self, inputs, results):
        """Create comprehensive visualizations"""
        # Set up the plotting style
        plt.style.use('seaborn-v0_8')
        fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(15, 12))
        
        # 1. Net Cost Comparison
        options = ['Option 1\n(Own Capital)', 'Option 2\n(Invest + Loan)']
        costs = [results['option1']['net_outflow'], results['option2']['net_outflow']]
        colors = ['#2E8B57' if results['recommendation'] == 'option1' else '#FF6B6B',
                  '#2E8B57' if results['recommendation'] == 'option2' else '#FF6B6B']
        
        bars1 = ax1.bar(options, costs, color=colors, alpha=0.7, edgecolor='black')
        ax1.set_title('Net Cost Comparison', fontsize=14, fontweight='bold')
        ax1.set_ylabel('Net Cost (₹ lakh)')
        
        # Add value labels on bars
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
        df = self.generate_year_wise_data(inputs, results)
        
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
        plt.savefig('project_financing_analysis.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        return df
    
    def export_to_excel(self, inputs, results, df, filename='project_financing_analysis.xlsx'):
        """Export detailed analysis to Excel"""
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
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
        
        print(f"\n📁 Analysis exported to: {filename}")

def main():
    """Main function to run the calculator"""
    calc = ProjectFinancingCalculator()
    
    print("🏦 PROJECT FINANCING CALCULATOR")
    print("=" * 50)
    
    # Example from your case study
    inputs = {
        'project_cost': 150,      # ₹1.5 crore
        'own_capital': 100,       # ₹1 crore
        'loan_rate': 10.5,        # 10.5% p.a.
        'loan_tenure': 7,         # 7 years
        'investment_type': 'FD',  # Fixed Deposit
        'investment_return': 7.0, # 7% p.a.
        'tax_rate': 30            # 30%
    }
    
    # For interactive input, uncomment below:
    """
    inputs = {
        'project_cost': float(input("Enter project cost (₹ lakh): ")),
        'own_capital': float(input("Enter own capital available (₹ lakh): ")),
        'loan_rate': float(input("Enter loan interest rate (%): ")),
        'loan_tenure': int(input("Enter loan tenure (years): ")),
        'investment_type': input("Enter investment type (FD/Liquid Funds/SGBs/Arbitrage Fund/Debt Funds): "),
        'investment_return': float(input("Enter expected investment return (%): ")),
        'tax_rate': float(input("Enter tax rate (%): "))
    }
    """
    
    # Calculate comparison
    results = calc.calculate_comparison(inputs)
    
    # Print detailed report
    calc.print_detailed_report(inputs, results)
    
    # Create visualizations
    df = calc.create_visualizations(inputs, results)
    
    # Export to Excel
    calc.export_to_excel(inputs, results, df)
    
    return results, df

if __name__ == "__main__":
    results, analysis_df = main()
