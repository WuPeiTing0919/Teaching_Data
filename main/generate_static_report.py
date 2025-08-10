import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.io as pio
from datetime import datetime
import os

def load_and_analyze_data():
    """Load and analyze the orders_clean data"""
    try:
        # Load the data
        df = pd.read_excel('clean/student_case_clean.xlsx', sheet_name='orders_clean')
        
        # Convert date columns to datetime with error handling
        df['order_date'] = pd.to_datetime(df['order_date'], errors='coerce')
        
        # Remove rows with invalid dates
        df = df.dropna(subset=['order_date'])
        
        # Extract month and year for analysis
        df['order_month'] = df['order_date'].dt.to_period('M')
        df['order_year'] = df['order_date'].dt.year
        
        # Calculate summary statistics
        total_orders = len(df)
        total_revenue = df['total_with_tax'].sum()
        total_products = df['product_id'].nunique()
        total_categories = df['category'].nunique()
        
        print(f"Loaded {total_orders} orders with valid dates")
        print(f"Data range: {df['order_date'].min()} to {df['order_date'].max()}")
        
        return df, {
            'total_orders': total_orders,
            'total_revenue': total_revenue,
            'total_products': total_products,
            'total_categories': total_categories
        }
    except Exception as e:
        print(f"Error loading data: {e}")
        return None, None

def create_monthly_revenue_chart(df):
    """Create monthly revenue chart"""
    monthly_revenue = df.groupby('order_month')['total_with_tax'].sum().reset_index()
    monthly_revenue['order_month'] = monthly_revenue['order_month'].astype(str)
    
    fig = px.line(monthly_revenue, x='order_month', y='total_with_tax',
                   title='Monthly Revenue Trend',
                   labels={'order_month': 'Month', 'total_with_tax': 'Revenue ($)'})
    fig.update_layout(height=500)
    
    return fig

def create_category_revenue_chart(df):
    """Create category revenue chart"""
    category_revenue = df.groupby('category')['total_with_tax'].sum().sort_values(ascending=True)
    
    fig = px.bar(x=category_revenue.values, y=category_revenue.index, orientation='h',
                  title='Revenue by Category',
                  labels={'x': 'Revenue ($)', 'y': 'Category'})
    fig.update_layout(height=500)
    
    return fig

def create_product_revenue_chart(df):
    """Create top product revenue chart"""
    product_revenue = df.groupby('product_name')['total_with_tax'].sum().sort_values(ascending=False).head(15)
    
    fig = px.bar(x=product_revenue.values, y=product_revenue.index, orientation='h',
                  title='Top 15 Products by Revenue',
                  labels={'x': 'Revenue ($)', 'y': 'Product'})
    fig.update_layout(height=600)
    
    return fig

def create_discount_analysis_chart(df):
    """Create discount analysis chart"""
    # Create scatter plot of discount vs revenue
    fig = px.scatter(df, x='discount', y='total_with_tax', 
                      title='Discount vs Revenue Analysis',
                      labels={'discount': 'Discount (%)', 'total_with_tax': 'Revenue ($)'})
    fig.update_layout(height=500)
    
    return fig

def create_summary_cards_html(summary_stats):
    """Create HTML for summary cards"""
    cards_html = f"""
    <div class="summary-cards">
        <div class="card">
            <h3>Total Orders</h3>
            <p class="number">{summary_stats['total_orders']:,}</p>
        </div>
        <div class="card">
            <h3>Total Revenue</h3>
            <p class="number">${summary_stats['total_revenue']:,.2f}</p>
        </div>
        <div class="card">
            <h3>Total Products</h3>
            <p class="number">{summary_stats['total_products']:,}</p>
        </div>
        <div class="card">
            <h3>Total Categories</h3>
            <p class="number">{summary_stats['total_categories']:,}</p>
        </div>
    </div>
    """
    return cards_html

def generate_static_report():
    """Generate static HTML report"""
    print("Loading and analyzing data...")
    df, summary_stats = load_and_analyze_data()
    
    if df is None:
        print("Failed to load data. Please check if the Excel file exists.")
        return
    
    # Create output directory
    os.makedirs('static_report', exist_ok=True)
    
    print("Generating charts...")
    
    # Generate individual chart HTML files
    charts = {
        'monthly_revenue': create_monthly_revenue_chart(df),
        'category_revenue': create_category_revenue_chart(df),
        'product_revenue': create_product_revenue_chart(df),
        'discount_analysis': create_discount_analysis_chart(df)
    }
    
    # Save charts as HTML files
    for name, fig in charts.items():
        html_file = f'static_report/{name}.html'
        fig.write_html(html_file, include_plotlyjs='cdn')
        print(f"Generated: {html_file}")
    
    # Create main report HTML
    main_html = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Student Case Orders Analysis Report</title>
        <style>
            body {{
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                margin: 0;
                padding: 20px;
                background-color: #f5f5f5;
            }}
            .container {{
                max-width: 1200px;
                margin: 0 auto;
                background-color: white;
                padding: 30px;
                border-radius: 10px;
                box-shadow: 0 0 20px rgba(0,0,0,0.1);
            }}
            h1 {{
                color: #2c3e50;
                text-align: center;
                margin-bottom: 30px;
                border-bottom: 3px solid #3498db;
                padding-bottom: 15px;
            }}
            .summary-cards {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
                gap: 20px;
                margin-bottom: 40px;
            }}
            .card {{
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                color: white;
                padding: 25px;
                border-radius: 15px;
                text-align: center;
                box-shadow: 0 5px 15px rgba(0,0,0,0.2);
            }}
            .card h3 {{
                margin: 0 0 15px 0;
                font-size: 18px;
                opacity: 0.9;
            }}
            .card .number {{
                font-size: 28px;
                font-weight: bold;
                margin: 0;
            }}
            .chart-section {{
                margin-bottom: 40px;
                padding: 20px;
                border: 1px solid #e0e0e0;
                border-radius: 10px;
                background-color: #fafafa;
            }}
            .chart-section h2 {{
                color: #34495e;
                margin-top: 0;
                margin-bottom: 20px;
                padding-bottom: 10px;
                border-bottom: 2px solid #3498db;
            }}
            iframe {{
                width: 100%;
                height: 600px;
                border: none;
                border-radius: 8px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            }}
            .footer {{
                text-align: center;
                margin-top: 40px;
                padding-top: 20px;
                border-top: 1px solid #e0e0e0;
                color: #7f8c8d;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>📊 Student Case Orders Analysis Report</h1>
            
            {create_summary_cards_html(summary_stats)}
            
            <div class="chart-section">
                <h2>📈 Monthly Revenue Trend</h2>
                <iframe src="monthly_revenue.html"></iframe>
            </div>
            
            <div class="chart-section">
                <h2>🏷️ Revenue by Category</h2>
                <iframe src="category_revenue.html"></iframe>
            </div>
            
            <div class="chart-section">
                <h2>📦 Top Products by Revenue</h2>
                <iframe src="product_revenue.html"></iframe>
            </div>
            
            <div class="chart-section">
                <h2>💰 Discount vs Revenue Analysis</h2>
                <iframe src="discount_analysis.html"></iframe>
            </div>
            
            <div class="footer">
                <p>Report generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
                <p>Data source: student_case_clean.xlsx (orders_clean sheet)</p>
            </div>
        </div>
    </body>
    </html>
    """
    
    # Save main report
    with open('static_report/main_report.html', 'w', encoding='utf-8') as f:
        f.write(main_html)
    
    print(f"\n✅ Static report generated successfully!")
    print(f"📁 Main report: static_report/main_report.html")
    print(f"📊 Individual charts saved in static_report/ folder")
    print(f"\nYou can now open 'static_report/main_report.html' in your browser to view the complete report.")

if __name__ == "__main__":
    generate_static_report()
