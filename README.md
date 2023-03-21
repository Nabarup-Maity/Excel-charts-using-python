# Time Series Chart in Excel using Python Openpyxl

This code creates a time series chart using the Openpyxl package in Python. The chart is based on a pandas DataFrame that has three columns, "Time" and "Value" and "Anomaly". It creates a bar chart and line chart to show the trend of the data and additionally it highlights the anomalous values in the chart based on Anomaly column.

__Input Data Frame__

|Time	|  Value	|   Anomaly  |
| --------- |:---------:| :---------: |
|2021-01-01	|10| False|
|2021-01-02	|15|	True|
|2021-01-03	|12|	True|
|2021-01-04	|17|	False|
|2021-01-05	|22|	False|
|2021-01-06	|14|	True|
|2021-01-07	|11|	False|


__Output:__

![image](https://user-images.githubusercontent.com/45371293/226655327-b0b843e8-6f4a-409b-bd52-7a069c6fc89a.png)

__Requirements__<br>
The following packages need to be installed:

* openpyxl<br>
* pandas

__Usage__<br>
1. Create a pandas DataFrame with columns "Time" and "Value". Optionally, a third column "Anomaly" can be included to highlight anomalous values in the chart.
2. Modify the DataFrame as required.
3. Run the script time_series_chart_in_excel.py


__Reference__<br>
https://openpyxl.readthedocs.io/en/latest/charts/introduction.html
