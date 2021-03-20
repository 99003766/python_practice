import openpyxl

# import BarChart class from openpyxl.chart sub_module
from openpyxl.chart import BarChart, Reference

#import openpyxl
#from openpyxl import Workbook


wb = openpyxl.load_workbook('Final.xlsx')

excel_file = wb
# Get workbook active sheet
# from the active attribute.
sheet = wb['MasterSheet11']

# write o to 9 in 1st column of the active sheet
# create data for plotting
values = Reference(sheet, min_col=2, min_row=9,
                   max_col=sheet.max_column, max_row=17)

# Create object of BarChart class
chart = BarChart()

# adding data to the Bar chart object
chart.add_data(values)

# set the title of the chart
chart.title = " BAR-CHART "

# set the title of the x-axis
chart.x_axis.title = " X_AXIS "

# set the title of the y-axis
chart.y_axis.title = " Y_AXIS "

# add chart to the sheet
# the top-left corner of a chart
# is anchored to cell E2 .
sheet.add_chart(chart, "E2")

# save the file
wb.save("Final.xlsx")
