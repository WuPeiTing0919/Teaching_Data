import pandas as pd
import json
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

def load_sales_data():
    """
    讀取 sales_clean.xlsx 的 Cleaned_Data 工作表
    """
    try:
        # 讀取 Excel 檔案的 Cleaned_Data 工作表
        df = pd.read_excel('../clean/sales_clean.xlsx', sheet_name='Cleaned_Data')
        print(f"成功讀取資料，共 {len(df)} 筆記錄")
        print(f"資料欄位: {df.columns.tolist()}")
        return df
    except Exception as e:
        print(f"讀取檔案時發生錯誤: {e}")
        return None

def create_html_dashboard(df):
    """
    創建 HTML 儀表板
    """
    # 計算統計數據
    total_sales = df['line_amount'].sum()
    total_orders = df['OrderID'].nunique()
    total_products = df['Product'].nunique()
    total_regions = df['Region'].nunique()
    
    # 當日銷售總額
    daily_sales = df.groupby('Order Date')['line_amount'].sum().reset_index()
    daily_sales['Order Date'] = pd.to_datetime(daily_sales['Order Date'])
    daily_sales = daily_sales.sort_values('Order Date')
    
    # 前 5 名商品
    top_products = df.groupby('Product')['line_amount'].sum().sort_values(ascending=False).head(5)
    
    # 按區域彙總
    region_sales = df.groupby('Region')['line_amount'].sum().reset_index()
    
    # 準備圖表數據
    chart_data = {
        'daily_sales': {
            'labels': daily_sales['Order Date'].dt.strftime('%Y-%m-%d').tolist(),
            'data': daily_sales['line_amount'].tolist()
        },
        'top_products': {
            'labels': top_products.index.tolist(),
            'data': top_products.values.tolist()
        },
        'region_sales': {
            'labels': region_sales['Region'].tolist(),
            'data': region_sales['line_amount'].tolist()
        }
    }
    
    # HTML 模板
    html_content = f"""
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>銷售資料分析儀表板</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
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
        }}
        .header {{
            text-align: center;
            margin-bottom: 30px;
            color: #333;
        }}
        .stats-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}
        .stat-card {{
            background: white;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            text-align: center;
        }}
        .stat-value {{
            font-size: 2em;
            font-weight: bold;
            margin: 10px 0;
        }}
        .stat-label {{
            color: #666;
            font-size: 0.9em;
        }}
        .charts-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(500px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}
        .chart-container {{
            background: white;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }}
        .chart-title {{
            text-align: center;
            margin-bottom: 20px;
            color: #333;
        }}
        .data-table {{
            background: white;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            margin-bottom: 30px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
        }}
        th, td {{
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid #ddd;
        }}
        th {{
            background-color: #f8f9fa;
            font-weight: bold;
        }}
        .text-primary {{ color: #007bff; }}
        .text-success {{ color: #28a745; }}
        .text-info {{ color: #17a2b8; }}
        .text-warning {{ color: #ffc107; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>銷售資料分析儀表板</h1>
            <p>資料更新時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        </div>
        
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-label">總銷售額</div>
                <div class="stat-value text-primary">NT$ {total_sales:,.0f}</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">總訂單數</div>
                <div class="stat-value text-success">{total_orders:,}</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">商品種類</div>
                <div class="stat-value text-info">{total_products}</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">銷售區域</div>
                <div class="stat-value text-warning">{total_regions}</div>
            </div>
        </div>
        
        <div class="charts-grid">
            <div class="chart-container">
                <div class="chart-title">當日銷售總額趨勢</div>
                <canvas id="dailySalesChart"></canvas>
            </div>
            <div class="chart-container">
                <div class="chart-title">前 5 名商品銷售總額</div>
                <canvas id="topProductsChart"></canvas>
            </div>
        </div>
        
        <div class="charts-grid">
            <div class="chart-container">
                <div class="chart-title">各區域銷售佔比</div>
                <canvas id="regionPieChart"></canvas>
            </div>
            <div class="chart-container">
                <div class="chart-title">各區域銷售總額</div>
                <canvas id="regionBarChart"></canvas>
            </div>
        </div>
        
        <div class="data-table">
            <h3>銷售資料明細 (前 50 筆)</h3>
            <table>
                <thead>
                    <tr>
                        <th>訂單ID</th>
                        <th>日期</th>
                        <th>產品</th>
                        <th>數量</th>
                        <th>單價</th>
                        <th>小計</th>
                        <th>區域</th>
                    </tr>
                </thead>
                <tbody>
"""
    
    # 添加表格數據
    for _, row in df.head(50).iterrows():
        html_content += f"""
                    <tr>
                        <td>{row['OrderID']}</td>
                        <td>{row['Order Date']}</td>
                        <td>{row['Product']}</td>
                        <td>{row['Qty']}</td>
                        <td>NT$ {row['Unit Price']:,.0f}</td>
                        <td>NT$ {row['line_amount']:,.0f}</td>
                        <td>{row['Region']}</td>
                    </tr>
"""
    
    html_content += f"""
                </tbody>
            </table>
        </div>
    </div>
    
    <script>
        // 圖表數據
        const chartData = {json.dumps(chart_data, ensure_ascii=False)};
        
        // 當日銷售總額趨勢圖
        new Chart(document.getElementById('dailySalesChart'), {{
            type: 'line',
            data: {{
                labels: chartData.daily_sales.labels,
                datasets: [{{
                    label: '當日銷售總額',
                    data: chartData.daily_sales.data,
                    borderColor: '#007bff',
                    backgroundColor: 'rgba(0, 123, 255, 0.1)',
                    tension: 0.1
                }}]
            }},
            options: {{
                responsive: true,
                plugins: {{
                    title: {{
                        display: true,
                        text: '當日銷售總額趨勢'
                    }}
                }},
                scales: {{
                    y: {{
                        beginAtZero: true,
                        ticks: {{
                            callback: function(value) {{
                                return 'NT$ ' + value.toLocaleString();
                            }}
                        }}
                    }}
                }}
            }}
        }});
        
        // 前 5 名商品銷售總額圖
        new Chart(document.getElementById('topProductsChart'), {{
            type: 'bar',
            data: {{
                labels: chartData.top_products.labels,
                datasets: [{{
                    label: '銷售總額',
                    data: chartData.top_products.data,
                    backgroundColor: '#ffc107',
                    borderColor: '#e0a800',
                    borderWidth: 1
                }}]
            }},
            options: {{
                responsive: true,
                indexAxis: 'y',
                plugins: {{
                    title: {{
                        display: true,
                        text: '前 5 名商品銷售總額'
                    }}
                }},
                scales: {{
                    x: {{
                        beginAtZero: true,
                        ticks: {{
                            callback: function(value) {{
                                return 'NT$ ' + value.toLocaleString();
                            }}
                        }}
                    }}
                }}
            }}
        }});
        
        // 各區域銷售佔比圖
        new Chart(document.getElementById('regionPieChart'), {{
            type: 'doughnut',
            data: {{
                labels: chartData.region_sales.labels,
                datasets: [{{
                    data: chartData.region_sales.data,
                    backgroundColor: [
                        '#007bff',
                        '#28a745',
                        '#ffc107',
                        '#dc3545',
                        '#6f42c1'
                    ]
                }}]
            }},
            options: {{
                responsive: true,
                plugins: {{
                    title: {{
                        display: true,
                        text: '各區域銷售佔比'
                    }},
                    legend: {{
                        position: 'bottom'
                    }}
                }}
            }}
        }});
        
        // 各區域銷售總額圖
        new Chart(document.getElementById('regionBarChart'), {{
            type: 'bar',
            data: {{
                labels: chartData.region_sales.labels,
                datasets: [{{
                    label: '銷售總額',
                    data: chartData.region_sales.data,
                    backgroundColor: '#28a745',
                    borderColor: '#1e7e34',
                    borderWidth: 1
                }}]
            }},
            options: {{
                responsive: true,
                plugins: {{
                    title: {{
                        display: true,
                        text: '各區域銷售總額'
                    }}
                }},
                scales: {{
                    y: {{
                        beginAtZero: true,
                        ticks: {{
                            callback: function(value) {{
                                return 'NT$ ' + value.toLocaleString();
                            }}
                        }}
                    }}
                }}
            }}
        }});
    </script>
</body>
</html>
"""
    
    return html_content

def main():
    """
    主函數
    """
    print("開始讀取銷售資料...")
    
    # 讀取資料
    df = load_sales_data()
    if df is None:
        print("無法讀取資料檔案，請檢查檔案路徑")
        return
    
    print("資料讀取成功！")
    print(f"資料筆數: {len(df)}")
    print(f"資料欄位: {df.columns.tolist()}")
    
    # 創建 HTML 儀表板
    print("創建 HTML 儀表板...")
    html_content = create_html_dashboard(df)
    
    # 儲存 HTML 檔案
    output_file = "sales_dashboard.html"
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"HTML 儀表板已儲存至: {output_file}")
    print("請在瀏覽器中開啟此檔案以查看分析結果")

if __name__ == "__main__":
    main()
