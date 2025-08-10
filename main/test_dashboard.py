import pandas as pd
import plotly.express as px
import dash
from dash import dcc, html, Input, Output, dash_table
import dash_bootstrap_components as dbc

def create_simple_dashboard():
    """
    創建簡化版儀表板
    """
    # 載入資料
    try:
        df = pd.read_excel('clean/student_case_clean.xlsx', sheet_name='orders_clean')
        print(f"成功載入資料，共 {len(df)} 筆記錄")
    except Exception as e:
        print(f"載入資料失敗: {e}")
        return None
    
    # 創建 Dash 應用
    app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
    
    # 簡化佈局
    app.layout = dbc.Container([
        html.H1("Student Case Orders 分析儀表板", className="text-center text-primary mb-4"),
        html.Hr(),
        
        # 統計摘要
        dbc.Row([
            dbc.Col([
                dbc.Card([
                    dbc.CardBody([
                        html.H4(f"{len(df)}", className="card-title text-center"),
                        html.P("總訂單數", className="card-text text-center")
                    ])
                ])
            ], width=3),
            dbc.Col([
                dbc.Card([
                    dbc.CardBody([
                        html.H4(f"${df['total_with_tax'].sum():,.2f}", className="card-title text-center"),
                        html.P("總營收", className="card-text text-center")
                    ])
                ])
            ], width=3),
            dbc.Col([
                dbc.Card([
                    dbc.CardBody([
                        html.H4(f"{df['product_id'].nunique()}", className="card-title text-center"),
                        html.P("產品種類", className="card-text text-center")
                    ])
                ])
            ], width=3),
            dbc.Col([
                dbc.Card([
                    dbc.CardBody([
                        html.H4(f"{df['category'].nunique()}", className="card-title text-center"),
                        html.P("產品類別", className="card-text text-center")
                    ])
                ])
            ], width=3)
        ], className="mb-4"),
        
        # 圖表
        dbc.Row([
            dbc.Col([
                dbc.Card([
                    dbc.CardHeader("產品類別營收分布"),
                    dbc.CardBody([
                        dcc.Graph(
                            figure=px.pie(
                                df.groupby('category')['total_with_tax'].sum().reset_index(),
                                values='total_with_tax',
                                names='category',
                                title='產品類別營收分布'
                            )
                        )
                    ])
                ])
            ], width=6),
            dbc.Col([
                dbc.Card([
                    dbc.CardHeader("產品營收排行"),
                    dbc.CardBody([
                        dcc.Graph(
                            figure=px.bar(
                                df.groupby('product_name')['total_with_tax'].sum().reset_index().sort_values('total_with_tax', ascending=True),
                                x='total_with_tax',
                                y='product_name',
                                orientation='h',
                                title='產品營收排行'
                            )
                        )
                    ])
                ])
            ], width=6)
        ], className="mb-4"),
        
        # 資料表格
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
        
    ], fluid=True)
    
    return app

def main():
    """
    主函數
    """
    print("開始創建簡化版 Student Case Orders 分析儀表板...")
    
    # 創建儀表板
    app = create_simple_dashboard()
    
    if app is None:
        print("儀表板創建失敗")
        return
    
    # 啟動應用
    print("儀表板已創建完成！")
    print("請在瀏覽器中開啟 http://127.0.0.1:8050 來查看儀表板")
    
    app.run_server(debug=True, host='127.0.0.1', port=8050)

if __name__ == "__main__":
    main()
