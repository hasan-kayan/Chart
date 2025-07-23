# excel_chart_ui/charts/chart_factory.py

import plotly.express as px

def create_chart(df, chart_type, x_col, y_cols):
    if chart_type == "line":
        return px.line(df, x=x_col, y=y_cols, title="Line Chart")
    elif chart_type == "bar":
        return px.bar(df, x=x_col, y=y_cols[0], title="Bar Chart")
    elif chart_type == "column":
        return px.bar(df, x=x_col, y=y_cols[0], orientation='v', title="Column Chart")
    elif chart_type == "scatter":
        return px.scatter(df, x=x_col, y=y_cols[0], title="Scatter Plot")
    elif chart_type == "area":
        return px.area(df, x=x_col, y=y_cols[0], title="Area Chart")
    elif chart_type == "pie":
        return px.pie(df, names=x_col, values=y_cols[0], title="Pie Chart")
    else:
        raise ValueError(f"Unsupported chart type: {chart_type}")
