# excel_chart_ui/charts/exporter.py

def export_chart(fig, path):
    if path.endswith(".html"):
        fig.write_html(path)
    else:
        fig.write_image(path, engine="kaleido")
