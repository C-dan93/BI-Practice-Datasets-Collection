import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta

# Set random seed for reproducibility
np.random.seed(42)
random.seed(42)

# Create messy data that requires all the transformations from your guide
def create_practice_dataset():
    
    # Sample data with intentional issues
    products = ['Laptop', 'Mouse', 'Keyboard', 'Monitor', 'Headphones', 'Webcam', 'Tablet', 'Phone']
    regions = ['North', 'South', 'East', 'West', 'Central']
    sales_reps = ['John Smith', 'Sarah Johnson', 'Mike Davis', 'Lisa Chen', 'David Wilson', 'Emma Brown']
    
    # Generate base data
    n_records = 500
    
    data = []
    
    for i in range(n_records):
        # Create intentionally messy data
        record = {
            'ORDER_ID': f'ORD-{1000 + i}',
            'Date': (datetime(2024, 1, 1) + timedelta(days=random.randint(0, 365))).strftime('%m/%d/%Y'),
            'Product': random.choice(products),
            'Quantity': str(random.randint(1, 10)),  # Numbers as text
            'Unit_Price': f'${random.randint(50, 2000)}.{random.randint(0, 99):02d}',  # Price with $ symbol
            'Sales_Rep': random.choice(sales_reps),
            'Region': random.choice(regions),
            'Customer_Type': random.choice(['Business', 'Individual', 'Government']),
            'Discount_Percent': f'{random.randint(0, 25)}%',  # Percentage as text
            'Status': random.choice(['Completed', 'Pending', 'Cancelled', 'Shipped'])
        }
        data.append(record)
    
    # Add some problematic records
    # Records with missing values
    for i in range(10):
        record = data[i].copy()
        record['Unit_Price'] = ''  # Empty price
        record['Region'] = None   # Null region
        data.append(record)
    
    # Records with wrong data types
    for i in range(5):
        record = data[i].copy()
        record['Quantity'] = 'N/A'  # Text in number field
        record['Date'] = 'Invalid Date'
        data.append(record)
    
    # Duplicate records
    for i in range(3):
        data.append(data[i].copy())
    
    df = pd.DataFrame(data)
    
    # Create the main messy dataset
    with pd.ExcelWriter('ecommerce_sales_practice.xlsx', engine='openpyxl') as writer:
        
        # Sheet 1: Main data with header issues
        messy_df = df.copy()
        
        # Add unwanted header rows
        header_rows = pd.DataFrame([
            ['E-Commerce Sales Report', '', '', '', '', '', '', '', '', ''],
            ['Generated on: 2024-07-27', '', '', '', '', '', '', '', '', ''],
            ['Confidential Data', '', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', '', '']
        ])
        
        # Combine header rows with actual data
        final_df = pd.concat([header_rows, messy_df], ignore_index=True)
        final_df.to_excel(writer, sheet_name='Sales_Data', index=False, header=False)
        
        # Sheet 2: Pivot-style data that needs unpivoting
        pivot_data = {
            'Product': ['Laptop', 'Mouse', 'Keyboard', 'Monitor', 'Headphones'],
            'Q1_2024': [150000, 25000, 15000, 75000, 30000],
            'Q2_2024': [180000, 30000, 18000, 80000, 35000],
            'Q3_2024': [165000, 28000, 16000, 78000, 32000],
            'Q4_2024': [200000, 35000, 20000, 85000, 40000]
        }
        pivot_df = pd.DataFrame(pivot_data)
        pivot_df.to_excel(writer, sheet_name='Quarterly_Sales', index=False)
        
        # Sheet 3: Customer data with different issues
        customer_data = {
            'CUSTOMER_ID': [f'CUST-{i:04d}' for i in range(1, 101)],
            'Customer Name': [f'Customer {i}' for i in range(1, 101)],
            'Email': [f'customer{i}@email.com' if i % 10 != 0 else '' for i in range(1, 101)],
            'Phone': [f'(555) {random.randint(100, 999)}-{random.randint(1000, 9999)}' for i in range(100)],
            'Registration_Date': [(datetime(2023, 1, 1) + timedelta(days=random.randint(0, 500))).strftime('%Y-%m-%d') for i in range(100)],
            'Total_Orders': [str(random.randint(1, 50)) for i in range(100)],  # Numbers as text
            'Lifetime_Value': [f'${random.randint(100, 10000)}.00' for i in range(100)]  # Currency as text
        }
        customer_df = pd.DataFrame(customer_data)
        customer_df.to_excel(writer, sheet_name='Customer_Data', index=False)

# Generate the dataset
create_practice_dataset()

print("Dataset created: ecommerce_sales_practice.xlsx")
print("\nThis dataset includes the following practice scenarios:")
print("📋 SHEET 1 - Sales_Data:")
print("  • Unwanted header rows (practice removing top rows)")
print("  • Mixed data types (numbers stored as text)")
print("  • Currency symbols in price fields")
print("  • Percentage symbols in discount fields")  
print("  • Missing/null values")
print("  • Duplicate records")
print("  • Invalid data entries")

print("\n📋 SHEET 2 - Quarterly_Sales:")
print("  • Pivot table format (practice unpivoting)")
print("  • Date columns spread horizontally")

print("\n📋 SHEET 3 - Customer_Data:")
print("  • Numbers stored as text")
print("  • Currency formatting issues")
print("  • Missing email addresses")
print("  • Inconsistent date formats")

print("\n🎯 SKILLS YOU'LL PRACTICE:")
print("  ✅ Remove unwanted rows (top rows)")
print("  ✅ Set proper headers")
print("  ✅ Change data types")
print("  ✅ Filter rows by condition (remove nulls/blanks)")
print("  ✅ Remove duplicate rows")
print("  ✅ Remove unnecessary columns")
print("  ✅ Unpivot columns (quarterly data)")
print("  ✅ Rename column headers")
print("  ✅ Handle data type conversion errors")

print("\n💡 SUGGESTED PRACTICE ORDER:")
print("1. Start with Sales_Data sheet - clean headers and data types")
print("2. Move to Quarterly_Sales - practice unpivoting")  
print("3. Finish with Customer_Data - handle missing values and formatting")
print("4. Create relationships between sheets")
print("5. Build visualizations with your clean data!")

print("\n🔧 Run this script to generate your practice dataset!")
