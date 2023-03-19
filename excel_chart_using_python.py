import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference, Series, shapes
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.axis import DateAxis
from openpyxl.chart.label import DataLabel
# Create example data
df = pd.DataFrame({
    'Time': ['2021-01-01', '2021-01-02', '2021-01-03', '2021-01-04', '2021-01-05'],
    'Value': [10, 15, 13, 17, 20],
    'Anomalous': [False, True, True, False, False],
    'Product': ['A', 'A', 'A', 'B', 'B']
})

# Create workbook and sheets
wb = Workbook()
ws_filter = wb.active
ws_filter.title = 'Filter'
ws_filter['A1'] = 'Product'
ws_filter['A2'] = df['Product'].unique()[0]  # Set default product

for product in df['Product'].unique():
    # Filter data by product
    df_product = df[df['Product'] == product].copy()

    # Create chart
    chart = LineChart()
    chart.title = product
    chart.x_axis.title = 'Time'
    chart.y_axis.title = 'Value'
    chart.y_axis.crossAx = 500
    chart.x_axis = DateAxis(crossAx=100)

    # Add data series to chart
    dates = Reference(ws_filter, min_col=2, min_row=2, max_row=len(df_product)+1)
    values = Reference(ws_filter, min_col=3, min_row=2, max_row=len(df_product)+1)
    series = Series(values, xvalues=dates, title='Value')
    chart.series.append(series)

    # Set line color to red for anomalous values
    for idx, val in enumerate(df_product['Value']):
        data_label = DataLabel(val)
        if df_product['Anomalous'][idx] == True:
            chart.series[0].graphicalProperties.line.solidFill = 'FF0000'  # Set point color to red
            chart.series[0].dataLabel = data_label
        else:
            chart.series[0].graphicalProperties.solidFill = '000000'  # Set point color to black
            chart.series[0].dataLabel = None


    # Add chart to worksheet
    ws_product = wb.create_sheet(title=product)
    ws_product.add_chart(chart, 'A1')

# Add product filter dropdown
products = df['Product'].unique().tolist()
products.sort()
ws_filter['A1'].data_validation = \
    DataValidation(type="list", formula1='"{}"'.format('","'.join(products)), allow_blank=True)

# Save workbook
wb.save('output.xlsx')
