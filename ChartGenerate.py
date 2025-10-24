import pandas as pd
import json
import numpy as np

# Read data from Excel
df_chart = pd.read_excel('NewData.xlsx', sheet_name='ChartData')
strikes = df_chart['Strike'].dropna().tolist()
gex_oi = df_chart['GEX-OI'].dropna().tolist()
abs_oi = df_chart['ABS-OI'].dropna().tolist()
gex_vol = df_chart['GEX-VOL'].dropna().tolist()
abs_vol = df_chart['ABS-VOL'].dropna().tolist()

spot = df_chart.iloc[0, 7]

df_sheet1 = pd.read_excel('NewData.xlsx', sheet_name='Sheet1')
call_vol = pd.to_numeric(df_sheet1['Call Volume'], errors='coerce').fillna(0).tolist()
put_vol = pd.to_numeric(df_sheet1['Put Volume'], errors='coerce').fillna(0).tolist()

# Table data
df_table = pd.read_excel('NewData.xlsx', sheet_name='ChartData', usecols='G:I', nrows=12)
df_table.columns = ['Label', 'QQQ', 'NQ']
df_table['NQ'] = np.round(df_table['NQ'] * 4) / 4  # Round to nearest quarter
df_table = df_table.dropna(subset=['Label'])  # Drop any rows with NaN in Label

# Generate HTML table with custom row colors
table_rows = []
for idx, row in df_table.iterrows():
    label = row['Label']
    qqq = row['QQQ']
    nq = row['NQ']
    
    # Determine color class
    color_class = ''
    if label in ['PG-OI', 'FG-OI', 'PG-TT', 'FG-TT']:
        color_class = 'green-text'
    elif label in ['FR-OI', 'NG-OI', 'FR-TT', 'NG-TT']:
        color_class = 'red-text'
    elif label in ['ABS-OI', 'ABS-VOL']:
        color_class = 'purple-text'
    
    table_rows.append(f"""
    <tr class="{color_class}">
        <td>{label}</td>
        <td>{qqq}</td>
        <td>{nq}</td>
    </tr>
    """)

table_html = f"""
<table class="table-style">
    <thead>
        <tr>
            <th>Label</th>
            <th>QQQ</th>
            <th>NQ</th>
        </tr>
    </thead>
    <tbody>
        {''.join(table_rows)}
    </tbody>
</table>
"""

# Ensure all lists have the same length (trim if necessary)
min_len = min(len(strikes), len(gex_oi), len(abs_oi), len(gex_vol), len(abs_vol), len(call_vol), len(put_vol))
strikes = strikes[:min_len]
gex_oi = gex_oi[:min_len]
abs_oi = abs_oi[:min_len]
gex_vol = gex_vol[:min_len]
abs_vol = abs_vol[:min_len]
call_vol = call_vol[:min_len]
put_vol = put_vol[:min_len]

# Prepare positive and negative for GEX bars
positive_gex_oi = [max(0, x) for x in gex_oi]
negative_gex_oi = [min(0, x) for x in gex_oi]
positive_gex_vol = [max(0, x) for x in gex_vol]
negative_gex_vol = [min(0, x) for x in gex_vol]

# Prepare data as list of {x, y}
def prepare_data(strikes, values):
    return [{"x": s, "y": v} for s, v in zip(strikes, values)]

pos_gex_oi_data = prepare_data(strikes, positive_gex_oi)
neg_gex_oi_data = prepare_data(strikes, negative_gex_oi)
abs_oi_data = prepare_data(strikes, abs_oi)
pos_gex_vol_data = prepare_data(strikes, positive_gex_vol)
neg_gex_vol_data = prepare_data(strikes, negative_gex_vol)
abs_vol_data = prepare_data(strikes, abs_vol)
call_vol_data = prepare_data(strikes, call_vol)
put_vol_data = prepare_data(strikes, put_vol)

# JSON dumps for JS
pos_gex_oi_json = json.dumps(pos_gex_oi_data)
neg_gex_oi_json = json.dumps(neg_gex_oi_data)
abs_oi_json = json.dumps(abs_oi_data)
pos_gex_vol_json = json.dumps(pos_gex_vol_data)
neg_gex_vol_json = json.dumps(neg_gex_vol_data)
abs_vol_json = json.dumps(abs_vol_data)
call_vol_json = json.dumps(call_vol_data)
put_vol_json = json.dumps(put_vol_data)
spot_json = json.dumps(spot)

# Generate HTML
html_content = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Options Data Chart</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-annotation@3/dist/chartjs-plugin-annotation.min.js"></script>
    <style>
        body {{ background-color: black; color: white; font-family: Arial, sans-serif; padding: 20px; margin: 0; }}
        .container {{ display: flex; gap: 20px; height: 600px; }}
        #chartContainer {{ flex: 1; background-color: #111; }}
        .table-container {{ flex: 0 0 300px; background-color: #111; padding: 10px; overflow-y: auto; }}
        table.table-style {{ 
            border-collapse: collapse; 
            width: 100%; 
            color: white; 
            background-color: #111; 
        }}
        table.table-style th, table.table-style td {{ 
            border: 1px solid #333; 
            padding: 8px; 
            text-align: left; 
        }}
        table.table-style th {{ 
            background-color: #222; 
            font-weight: bold; 
        }}
        table.table-style tr:nth-child(even) {{ background-color: #222; }}
        table.table-style tr:hover {{ background-color: #333; }}
        .green-text {{ color: #00ff00; }}
        .red-text {{ color: #ff0000; }}
        .purple-text {{ color: #9370DB; }}
    </style>
</head>
<body>
    <div class="container">
        <div id="chartContainer">
            <canvas id="myChart"></canvas>
        </div>
        <div class="table-container">
            {table_html}
        </div>
    </div>

    <script>
        const ctx = document.getElementById('myChart').getContext('2d');
        const spot = {spot_json};
        const data = {{
            datasets: [
                {{
                    label: 'GEX-OI',
                    data: {pos_gex_oi_json},
                    type: 'bar',
                    backgroundColor: '#007bff',
                    stack: 'gex-oi',
                    borderColor: '#007bff',
                    borderWidth: 1,
                    yAxisID: 'y'
                }},
                {{
                    label: '',
                    data: {neg_gex_oi_json},
                    type: 'bar',
                    backgroundColor: '#dc3545',
                    stack: 'gex-oi',
                    borderColor: '#dc3545',
                    borderWidth: 1,
                    yAxisID: 'y'
                }},
                {{
                    label: 'ABS-OI',
                    data: {abs_oi_json},
                    type: 'line',
                    borderColor: '#6f42c1',
                    backgroundColor: '#6f42c1',
                    fill: false,
                    tension: 0.1,
                    yAxisID: 'y1',
                    borderWidth: 2
                }},
                {{
                    label: 'GEX-VOL',
                    data: {pos_gex_vol_json},
                    type: 'bar',
                    backgroundColor: '#007bff',
                    stack: 'gex-vol',
                    borderColor: '#007bff',
                    borderWidth: 1,
                    yAxisID: 'y'
                }},
                {{
                    label: '',
                    data: {neg_gex_vol_json},
                    type: 'bar',
                    backgroundColor: '#dc3545',
                    stack: 'gex-vol',
                    borderColor: '#dc3545',
                    borderWidth: 1,
                    yAxisID: 'y'
                }},
                {{
                    label: 'ABS-VOL',
                    data: {abs_vol_json},
                    type: 'line',
                    borderColor: '#8b5cf6',
                    backgroundColor: '#8b5cf6',
                    fill: false,
                    tension: 0.1,
                    yAxisID: 'y1',
                    borderWidth: 2
                }},
                {{
                    label: 'Call-VOL',
                    data: {call_vol_json},
                    type: 'line',
                    borderColor: '#f59e0b',
                    backgroundColor: '#f59e0b',
                    fill: false,
                    tension: 0.1,
                    yAxisID: 'y1',
                    borderWidth: 2
                }},
                {{
                    label: 'Put-VOL',
                    data: {put_vol_json},
                    type: 'line',
                    borderColor: '#10b981',
                    backgroundColor: '#10b981',
                    fill: false,
                    tension: 0.1,
                    yAxisID: 'y1',
                    borderWidth: 2
                }}
            ]
        }};

        const config = {{
            type: 'bar',
            data: data,
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{
                    legend: {{
                        labels: {{
                            filter: function(legendItem) {{
                                return legendItem.text !== '';
                            }},
                            color: 'white',
                            usePointStyle: true
                        }},
                        onClick: function(e, legendItem, legend) {{
                            const ci = this.chart;
                            const index = legendItem.datasetIndex;
                            let indices = [];
                            if (index === 0) {{ // GEX-OI
                                indices = [0,1];
                            }} else if (index === 3) {{ // GEX-VOL
                                indices = [3,4];
                            }} else {{
                                indices = [index];
                            }}
                            indices.forEach(idx => {{
                                const meta = ci.getDatasetMeta(idx);
                                meta.hidden = !meta.hidden;
                            }});
                            ci.update();
                        }}
                    }},
                    annotation: {{
                        annotations: {{
                            spotLine: {{
                                type: 'line',
                                scaleID: 'x',
                                value: spot,
                                borderColor: 'white',
                                borderWidth: 2,
                                borderDash: [5, 5],
                                label: {{
                                    display: true,
                                    content: 'Spot ' + spot,
                                    position: 'top'
                                }}
                            }}
                        }}
                    }}
                }},
                scales: {{
                    x: {{
                        type: 'linear',
                        stacked: true,
                        ticks: {{
                            color: 'white',
                            stepSize: 1,
                            maxRotation: 45
                        }},
                        grid: {{
                            display: false
                        }}
                    }},
                    y: {{
                        type: 'linear',
                        display: true,
                        position: 'left',
                        stacked: true,
                        ticks: {{
                            color: 'white',
                            callback: function(value) {{
                                return value.toLocaleString();
                            }}
                        }},
                        grid: {{
                            display: false
                        }}
                    }},
                    y1: {{
                        type: 'linear',
                        display: true,
                        position: 'right',
                        stacked: false,
                        ticks: {{
                            color: 'white',
                            callback: function(value) {{
                                return value.toLocaleString();
                            }},
                            maxTicksLimit: 10
                        }},
                        grid: {{
                            display: false
                        }}
                    }}
                }},
                interaction: {{
                    intersect: false,
                    mode: 'index'
                }}
            }}
        }};

        const myChart = new Chart(ctx, config);
    </script>
</body>
</html>
"""

# Write to file
with open('options_chart.html', 'w') as f:
    f.write(html_content)

print("HTML chart generated: options_chart.html")