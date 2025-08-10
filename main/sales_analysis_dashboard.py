import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import dash
from dash import dcc, html, dash_table
from dash.dependencies import Input, Output
import dash_bootstrap_components as dbc
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

def load_sales_data():
    """
    讀取 sales_clean.xlsx 的 Cleaned_Data 工作表
    """
    try:
        # 讀取 Excel 檔案的 Cleaned_Data 工作表
        df = pd.read_excel('clean/sales_clean.xlsx', sheet_name='Cleaned_Data')
        print(f"成功讀取資料，共 {len(df)} 筆記錄")
        print(f"資料欄位: {df.columns.tolist()}")
        return df
    except Exception as e:
        print(f"讀取檔案時發生錯誤: {e}")
        return None

def create_daily_sales_chart(df):
    """
    創建當日銷售總額圖表
    """
    # 按日期彙總銷售總額
    daily_sales = df.groupby('Order Date')['line_amount'].sum().reset_index()
    daily_sales['Order Date'] = pd.to_datetime(daily_sales['Order Date'])
    daily_sales = daily_sales.sort_values('Order Date')
    
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=daily_sales['Order Date'],
        y=daily_sales['line_amount'],
        mode='lines+markers',
        name='當日銷售總額',
        line=dict(color='#1f77b4', width=3),
        marker=dict(size=8)
    ))
    
    fig.update_layout(
        title='當日銷售總額趨勢',
        xaxis_title='日期',
        yaxis_title='銷售總額 (NT$)',
        template='plotly_white',
        height=400
    )
    
    return fig

def create_top_products_chart(df):
    """
    創建前 5 名商品圖表
    """
    # 按產品彙總銷售總額，取前 5 名
    top_products = df.groupby('Product')['line_amount'].sum().sort_values(ascending=False).head(5)
    
    fig = go.Figure(data=[
        go.Bar(
            x=top_products.values,
            y=top_products.index,
            orientation='h',
            marker_color='#ff7f0e',
            text=[f'NT$ {val:,.0f}' for val in top_products.values],
            textposition='auto'
        )
    ])
    
    fig.update_layout(
        title='前 5 名商品銷售總額',
        xaxis_title='銷售總額 (NT$)',
        yaxis_title='商品名稱',
        template='plotly_white',
        height=400
    )
    
    return fig

def create_region_summary_chart(df):
    """
    創建按區域彙總圖表
    """
    # 按區域彙總銷售總額
    region_sales = df.groupby('Region')['line_amount'].sum().reset_index()
    
    fig = go.Figure(data=[
        go.Pie(
            labels=region_sales['Region'],
            values=region_sales['line_amount'],
            hole=0.4,
            textinfo='label+percent',
            textposition='outside',
            marker=dict(colors=['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd'])
        )
    ])
    
    fig.update_layout(
        title='各區域銷售佔比',
        template='plotly_white',
        height=400
    )
    
    return fig

def create_region_bar_chart(df):
    """
    創建區域銷售柱狀圖
    """
    # 按區域彙總銷售總額
    region_sales = df.groupby('Region')['line_amount'].sum().reset_index()
    
    fig = go.Figure(data=[
        go.Bar(
            x=region_sales['Region'],
            y=region_sales['line_amount'],
            marker_color='#2ca02c',
            text=[f'NT$ {val:,.0f}' for val in region_sales['line_amount']],
            textposition='auto'
        )
    ])
    
    fig.update_layout(
        title='各區域銷售總額',
        xaxis_title='區域',
        yaxis_title='銷售總額 (NT$)',
        template='plotly_white',
        height=400
    )
    
    return fig

def create_summary_stats(df):
    """
    創建摘要統計資訊
    """
    # 計算各種統計數據
    total_sales = df['line_amount'].sum()
    total_orders = df['OrderID'].nunique()
    total_products = df['Product'].nunique()
    total_regions = df['Region'].nunique()
    avg_order_value = total_sales / total_orders if total_orders > 0 else 0
    
    # 最新銷售日期
    latest_date = pd.to_datetime(df['Order Date'].max())
    
    return {
        'total_sales': total_sales,
        'total_orders': total_orders,
        'total_products': total_products,
        'total_regions': total_regions,
        'avg_order_value': avg_order_value,
        'latest_date': latest_date
    }

def create_dashboard():
    """
    創建 Dash 儀表板
    """
    # 讀取資料
    df = load_sales_data()
    if df is None:
        return "無法讀取資料檔案"
    
    # 創建 Dash 應用
    app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
    
    # 獲取摘要統計
    stats = create_summary_stats(df)
    
    # 創建圖表
    daily_sales_fig = create_daily_sales_chart(df)
    top_products_fig = create_top_products_chart(df)
    region_summary_fig = create_region_summary_chart(df)
    region_bar_fig = create_region_bar_chart(df)
    
    # 儀表板佈局
    app.layout = dbc.Container([
        dbc.Row([
            dbc.Col([
                html.H1("銷售資料分析儀表板", className="text-center mb-4"),
                html.Hr()
            ])
        ]),
        
        # 摘要統計卡片
        dbc.Row([
            dbc.Col([
                dbc.Card([
                    dbc.CardBody([
                        html.H4("總銷售額", className="card-title"),
                        html.H2(f"NT$ {stats['total_sales']:,.0f}", className="text-primary")
                    ])
                ], className="text-center")
            ], width=3),
            dbc.Col([
                dbc.Card([
                    dbc.CardBody([
                        html.H4("總訂單數", className="card-title"),
                        html.H2(f"{stats['total_orders']:,}", className="text-success")
                    ])
                ], className="text-center")
            ], width=3),
            dbc.Col([
                dbc.Card([
                    dbc.CardBody([
                        html.H4("商品種類", className="card-title"),
                        html.H2(f"{stats['total_products']}", className="text-info")
                    ])
                ], className="text-center")
            ], width=3),
            dbc.Col([
                dbc.Card([
                    dbc.CardBody([
                        html.H4("銷售區域", className="card-title"),
                        html.H2(f"{stats['total_regions']}", className="text-warning")
                    ])
                ], className="text-center")
            ], width=3)
        ], className="mb-4"),
        
        # 圖表區域
        dbc.Row([
            dbc.Col([
                dcc.Graph(figure=daily_sales_fig)
            ], width=6),
            dbc.Col([
                dcc.Graph(figure=top_products_fig)
            ], width=6)
        ], className="mb-4"),
        
        dbc.Row([
            dbc.Col([
                dcc.Graph(figure=region_summary_fig)
            ], width=6),
            dbc.Col([
                dcc.Graph(figure=region_bar_fig)
            ], width=6)
        ], className="mb-4"),
        
        # 資料表格
        dbc.Row([
            dbc.Col([
                html.H3("銷售資料明細", className="mb-3"),
                html.Div([
                    dash_table.DataTable(
                        id='sales-table',
                        columns=[
                            {"name": "訂單ID", "id": "OrderID"},
                            {"name": "日期", "id": "Order Date"},
                            {"name": "產品", "id": "Product"},
                            {"name": "數量", "id": "Qty"},
                            {"name": "單價", "id": "Unit Price"},
                            {"name": "小計", "id": "line_amount"},
                            {"name": "區域", "id": "Region"}
                        ],
                        data=df.head(100).to_dict('records'),  # 只顯示前100筆
                        page_size=20,
                        style_table={'overflowX': 'auto'},
                        style_cell={'textAlign': 'center'},
                        style_header={'backgroundColor': 'rgb(230, 230, 230)', 'fontWeight': 'bold'}
                    )
                ])
            ])
        ])
    ], fluid=True)
    
    return app

if __name__ == "__main__":
    # 創建儀表板
    app = create_dashboard()
    
    if isinstance(app, str):
        print(app)
    else:
        print("啟動銷售分析儀表板...")
        print("請在瀏覽器中開啟: http://127.0.0.1:8050")
        app.run_server(debug=True, host='127.0.0.1', port=8050)
