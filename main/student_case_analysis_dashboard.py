import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import dash
from dash import dcc, html, Input, Output, dash_table
import dash_bootstrap_components as dbc
import numpy as np
from datetime import datetime

def load_and_analyze_data():
    """
    載入並分析 student_case_clean.xlsx 中的 orders_clean 資料
    """
    try:
        # 讀取 Excel 檔案
        file_path = 'clean/student_case_clean.xlsx'
        orders_df = pd.read_excel(file_path, sheet_name='orders_clean')
        
        print(f"成功載入 orders_clean 資料，共 {len(orders_df)} 筆記錄")
        print(f"欄位: {orders_df.columns.tolist()}")
        
        # 基本資料清理和轉換
        orders_df['order_date'] = pd.to_datetime(orders_df['order_date'])
        orders_df['month'] = orders_df['order_date'].dt.month
        orders_df['month_name'] = orders_df['order_date'].dt.strftime('%B')
        orders_df['year'] = orders_df['order_date'].dt.year
        
        # 計算總金額（如果沒有 total_with_tax 欄位）
        if 'total_with_tax' not in orders_df.columns:
            if 'subtotal' in orders_df.columns:
                orders_df['total_with_tax'] = orders_df['subtotal']
            else:
                orders_df['total_with_tax'] = orders_df['qty'] * orders_df['unit_price'] * (1 - orders_df['discount'] / 100)
        
        return orders_df
        
    except Exception as e:
        print(f"載入資料時發生錯誤: {str(e)}")
        return None

def create_dashboard():
    """
    創建 Dash 儀表板
    """
    # 載入資料
    df = load_and_analyze_data()
    if df is None:
        return "無法載入資料"
    
    # 創建 Dash 應用
    app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
    
    # 儀表板佈局
    app.layout = dbc.Container([
        dbc.Row([
            dbc.Col([
                html.H1("Student Case Orders 分析儀表板", 
                        className="text-center text-primary mb-4"),
                html.Hr()
            ])
        ]),
        
        # 統計摘要卡片
        dbc.Row([
            dbc.Col([
                dbc.Card([
                    dbc.CardBody([
                        html.H4(f"{len(df)}", className="card-title text-center"),
                        html.P("總訂單數", className="card-text text-center")
                    ])
                ], className="text-center mb-3")
            ], width=3),
            dbc.Col([
                dbc.Card([
                    dbc.CardBody([
                        html.H4(f"${df['total_with_tax'].sum():,.2f}", className="card-title text-center"),
                        html.P("總營收", className="card-text text-center")
                    ])
                ], className="text-center mb-3")
            ], width=3),
            dbc.Col([
                dbc.Card([
                    dbc.CardBody([
                        html.H4(f"{df['product_id'].nunique()}", className="card-title text-center"),
                        html.P("產品種類", className="card-text text-center")
                    ])
                ], className="text-center mb-3")
            ], width=3),
            dbc.Col([
                dbc.Card([
                    dbc.CardBody([
                        html.H4(f"{df['category'].nunique()}", className="card-title text-center"),
                        html.P("產品類別", className="card-text text-center")
                    ])
                ], className="text-center mb-3")
            ], width=3)
        ]),
        
        # 圖表區域
        dbc.Row([
            # 左側圖表
            dbc.Col([
                dbc.Card([
                    dbc.CardHeader("營收趨勢（按月份）"),
                    dbc.CardBody([
                        dcc.Graph(id='monthly-revenue-chart')
                    ])
                ], className="mb-3"),
                
                dbc.Card([
                    dbc.CardHeader("產品類別營收分布"),
                    dbc.CardBody([
                        dcc.Graph(id='category-revenue-chart')
                    ])
                ], className="mb-3")
            ], width=6),
            
            # 右側圖表
            dbc.Col([
                dbc.Card([
                    dbc.CardHeader("產品營收排行"),
                    dbc.CardBody([
                        dcc.Graph(id='product-revenue-chart')
                    ])
                ], className="mb-3"),
                
                dbc.Card([
                    dbc.CardHeader("折扣分析"),
                    dbc.CardBody([
                        dcc.Graph(id='discount-analysis-chart')
                    ])
                ], className="mb-3")
            ], width=6)
        ]),
        
        # 資料表格
        dbc.Row([
            dbc.Col([
                dbc.Card([
                    dbc.CardHeader("訂單詳細資料"),
                    dbc.CardBody([
                        dash_table.DataTable(
                            id='orders-table',
                            columns=[
                                {"name": "訂單日期", "id": "order_date"},
                                {"name": "產品名稱", "id": "product_name"},
                                {"name": "類別", "id": "category"},
                                {"name": "數量", "id": "qty"},
                                {"name": "單價", "id": "unit_price"},
                                {"name": "折扣", "id": "discount"},
                                {"name": "總金額", "id": "total_with_tax"}
                            ],
                            data=df.to_dict('records'),
                            page_size=10,
                            style_table={'overflowX': 'auto'},
                            style_cell={'textAlign': 'center'},
                            style_header={'backgroundColor': 'rgb(230, 230, 230)', 'fontWeight': 'bold'}
                        )
                    ])
                ])
            ])
        ]),
        
        # 篩選器
        dbc.Row([
            dbc.Col([
                dbc.Card([
                    dbc.CardHeader("資料篩選"),
                    dbc.CardBody([
                        dbc.Row([
                            dbc.Col([
                                html.Label("產品類別:"),
                                dcc.Dropdown(
                                    id='category-filter',
                                    options=[{'label': cat, 'value': cat} for cat in sorted(df['category'].unique())],
                                    value=sorted(df['category'].unique()),
                                    multi=True,
                                    placeholder="選擇產品類別"
                                )
                            ], width=4),
                            dbc.Col([
                                html.Label("日期範圍:"),
                                dcc.DatePickerRange(
                                    id='date-filter',
                                    start_date=df['order_date'].min(),
                                    end_date=df['order_date'].max(),
                                    display_format='YYYY-MM-DD'
                                )
                            ], width=4),
                            dbc.Col([
                                html.Label("最小金額:"),
                                dcc.Input(
                                    id='min-amount-filter',
                                    type='number',
                                    placeholder='最小金額',
                                    value=0
                                )
                            ], width=4)
                        ])
                    ])
                ], className="mb-3")
            ])
        ])
        
    ], fluid=True)
    
    # 回調函數：更新圖表
    @app.callback(
        [Output('monthly-revenue-chart', 'figure'),
         Output('category-revenue-chart', 'figure'),
         Output('product-revenue-chart', 'figure'),
         Output('discount-analysis-chart', 'figure'),
         Output('orders-table', 'data')],
        [Input('category-filter', 'value'),
         Input('date-filter', 'start_date'),
         Input('date-filter', 'end_date'),
         Input('min-amount-filter', 'value')]
    )
    def update_charts(selected_categories, start_date, end_date, min_amount):
        # 篩選資料
        filtered_df = df.copy()
        
        if selected_categories:
            filtered_df = filtered_df[filtered_df['category'].isin(selected_categories)]
        
        if start_date and end_date:
            filtered_df = filtered_df[
                (filtered_df['order_date'] >= start_date) & 
                (filtered_df['order_date'] <= end_date)
            ]
        
        if min_amount is not None:
            filtered_df = filtered_df[filtered_df['total_with_tax'] >= min_amount]
        
        # 1. 月度營收趨勢圖
        monthly_revenue = filtered_df.groupby(['year', 'month', 'month_name'])['total_with_tax'].sum().reset_index()
        monthly_revenue = monthly_revenue.sort_values(['year', 'month'])
        
        fig1 = px.line(
            monthly_revenue, 
            x='month_name', 
            y='total_with_tax',
            color='year',
            title='月度營收趨勢',
            labels={'month_name': '月份', 'total_with_tax': '營收 ($)', 'year': '年份'}
        )
        fig1.update_layout(xaxis={'categoryorder': 'array', 'categoryarray': ['January', 'February', 'March']})
        
        # 2. 產品類別營收分布
        category_revenue = filtered_df.groupby('category')['total_with_tax'].sum().reset_index()
        fig2 = px.pie(
            category_revenue, 
            values='total_with_tax', 
            names='category',
            title='產品類別營收分布'
        )
        
        # 3. 產品營收排行
        product_revenue = filtered_df.groupby('product_name')['total_with_tax'].sum().reset_index()
        product_revenue = product_revenue.sort_values('total_with_tax', ascending=True)
        fig3 = px.bar(
            product_revenue, 
            x='total_with_tax', 
            y='product_name',
            orientation='h',
            title='產品營收排行',
            labels={'total_with_tax': '營收 ($)', 'product_name': '產品名稱'}
        )
        
        # 4. 折扣分析
        discount_stats = filtered_df.groupby('discount')['total_with_tax'].sum().reset_index()
        fig4 = px.scatter(
            discount_stats, 
            x='discount', 
            y='total_with_tax',
            size='total_with_tax',
            title='折扣與營收關係',
            labels={'discount': '折扣 (%)', 'total_with_tax': '營收 ($)'}
        )
        
        # 更新表格資料
        table_data = filtered_df.to_dict('records')
        
        return fig1, fig2, fig3, fig4, table_data
    
    return app

def main():
    """
    主函數
    """
    print("開始創建 Student Case Orders 分析儀表板...")
    
    # 創建儀表板
    app = create_dashboard()
    
    if isinstance(app, str):
        print(f"錯誤: {app}")
        return
    
    # 啟動應用
    print("儀表板已創建完成！")
    print("請在瀏覽器中開啟 http://127.0.0.1:8050 來查看儀表板")
    
    app.run_server(debug=True, host='127.0.0.1', port=8050)

if __name__ == "__main__":
    main()
