"""

CODE IS A WORK IN PROGRESS

FEEL FREE TO MODIFY IT BASED ON YOUR WORKFLOW

LET'S GET FRACTAL !

"""

import PySimpleGUI as sg
import plotly.express as px
import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio
pio.renderers.default = 'svg'
pio.renderers.default = 'browser'
import xlsxwriter as xw
import xlrd
import openpyxl
import os

# Create a file selection dialogue box
layout = [
    [sg.Text("Select file:"), sg.Input(key="-FILE-"), sg.FileBrowse(file_types=(("Excel Files", "*.xlsx"),("Excel Files", "*.xls"), ("CSV Files", "*.csv")))],
    [sg.Text("Sheet Number / CSV File:"), sg.Input(key="-SHEET-NUM-", default_text = "1")],
    [sg.Text("Select Plot Type:"), sg.Combo(["Line Plot", "Scatter Plot", "Bar Plot", "Histogram", "Box Plot", "Pie Chart", "Heatmap", "3D Plot", "Violin Plot", "Area Plot"], key="-PLOT-TYPE-", default_value="Line Plot")],
    [sg.Text("Plot Title:"), sg.Input(key="-PLOT-TITLE-")],
    [sg.Button("Select Columns"), sg.Button("Plot"), sg.Button("Exit")]
]

#filename = sg.Input(key="-PLOT-TITLE-")
window = sg.Window("Plot Data", layout)

x_columns = []
y_columns = []

# Event loop
while True:
    event, values = window.read()

    if event == "Exit" or event == sg.WINDOW_CLOSED:
        break

    if event == "-FILE-":
        file_path = values["-FILE-"]

        if file_path:
            try:
                extension = file_path.split(".")[-1]

                if extension == "xlsx" or "xls":
                    # Read the available sheet names from the Excel file
                    xl = pd.ExcelFile(file_path)
                    sheet_names = xl.sheet_names

                    # Update the sheet selection to the first sheet by default
                    if sheet_names:
                        window["-SHEET-NUM-"].update(value="1")

                elif extension == "csv":
                    # Disable sheet selection for CSV files
                    window["-SHEET-NUM-"].update(disabled=True)

            except Exception as e:
                sg.popup_error(f"An error occurred while reading the file:\n{str(e)}")
                window["-FILE-"].update("")
                window["-SHEET-NUM-"].update("")


    if event == "Select Columns":
        file_path = values["-FILE-"]
        selected_sheet_num = values["-SHEET-NUM-"]

        if file_path:
            try:
                extension = file_path.split(".")[-1]

                if extension in ["xlsx", "xls"] and selected_sheet_num:
                    # Load the data from the selected sheet number in the Excel file
                    sheet_num = int(selected_sheet_num) - 1
                    df = pd.read_excel(file_path, sheet_name=sheet_num)
                elif extension == "csv":
                    # Load the data from the CSV file
                    df = pd.read_csv(file_path)
                else:
                    sg.popup_error("Unsupported file format.")
                
            # Get the column names for selection
                column_names = df.columns.tolist()

                # Create a column selection dialog
                layout = [
                    [sg.Text("Select X columns:")],
                    [sg.Listbox(column_names, size=(30, 10), key="-X-COLUMNS-", select_mode=sg.LISTBOX_SELECT_MODE_MULTIPLE)],
                    [sg.Text("Select Y columns:")],
                    [sg.Listbox(column_names, size=(30, 10), key="-Y-COLUMNS-", select_mode=sg.LISTBOX_SELECT_MODE_MULTIPLE)],

                    #[sg.Text("Select Z columns:")],
                    #[sg.Listbox(column_names, size=(30, 10), key="-Z-COLUMNS-", select_mode=sg.LISTBOX_SELECT_MODE_MULTIPLE)],

                    [sg.Button("OK"), sg.Button("Cancel")]
                ]

                column_selection_window = sg.Window("Select Columns", layout)

            # Event loop for column selection
                while True:
                    event, values = column_selection_window.read()
    
                    if event == "Cancel" or event == sg.WINDOW_CLOSED:
                        break
    
                    if event == "OK":
                        x_columns = values["-X-COLUMNS-"]
                        y_columns = values["-Y-COLUMNS-"]
                        column_selection_window.close()
                        break

                column_selection_window.close()

            except Exception as e:
                sg.popup_error(f"An error occurred while reading the file:\n{str(e)}")

    if event == "Plot":
        file_path = values["-FILE-"]
        selected_sheet_num = values["-SHEET-NUM-"]
        plot_type = values["-PLOT-TYPE-"]
        plot_title = values["-PLOT-TITLE-"]
    
        if file_path and selected_sheet_num and x_columns and y_columns:
            try:
                # Load the data from the selected sheet number in the Excel file
                if extension == "xlsx" or "xls":
                    sheet_num = int(selected_sheet_num) - 1
                    df = pd.read_excel(file_path, sheet_name=sheet_num)
    
                elif extension == "csv":
                    df = pd.read_csv(file_path)
    
                fig = go.Figure()
                # Update layout and display the plot
                fig.update_layout(title=plot_title, xaxis_title="X", yaxis_title="Y", legend_title="Columns")
                #fig.show()
                
                #
    
                # Plot each combination of X and Y columns
                for x_col in x_columns:
                    for y_col in y_columns:
                        if plot_type == "Line Plot":
                            fig.add_trace(go.Scatter(x=df[x_col], y=df[y_col], mode='lines', name=f'{y_col} vs {x_col}'))
                        elif plot_type == "Scatter Plot":
                            fig.add_trace(go.Scatter(x=df[x_col], y=df[y_col], mode='markers', name=f'{y_col} vs {x_col}'))
                        elif plot_type == "Bar Plot":
                            fig.add_trace(go.Bar(x=df[x_col], y=df[y_col], name=f'{y_col} vs {x_col}'))
                        elif plot_type == "Histogram":
                            fig.add_trace(go.Histogram(x=df[x_col], y=df[y_col], name=f'{y_col} vs {x_col}'))
                        
                        elif plot_type == "Box Plot":
                            fig.add_trace(go.Box(x=df[x_col], y=df[y_col], name=f'{y_col} vs {x_col}')) #boxmode='group') #boxmean='sd') # group together boxes of the different traces for each value of x

                        elif plot_type == "Pie Chart":
                            fig.add_trace(go.Pie(labels=df[x_col], values=df[y_col], name=f'{y_col} vs {x_col}'))
                        elif plot_type == "Heatmap":
                            fig.add_trace(go.Heatmap(x=df[x_col], y=df[y_col], z=df[y_col], name=f'{y_col} vs {x_col}'))
                        elif plot_type == "3D Plot":
                            fig.add_trace(go.Scatter3d(x=df[x_col], y=df[y_col], z=df[y_col], mode='markers', name=f'{y_col} vs {x_col}'))
                        elif plot_type == "Violin Plot":
                            fig.add_trace(go.Violin(x=df[x_col], y=df[y_col], name=f'{y_col} vs {x_col}'))
                        elif plot_type == "Area Plot":
                            fig.add_trace(go.Scatter(x=df[x_col], y=df[y_col], mode='lines', fill='tozeroy', name=f'{y_col} vs {x_col}'))
    
    
                    # Update layout and display the plot
                    fig.update_layout(title=plot_title, xaxis_title="X", yaxis_title="Y", legend_title="Columns")
                    #fig.show()
                    pio.show(fig)
                    
                    #Save the plot as an HTML file in the same directory as the script file
                    script_directory = os.path.dirname(os.path.abspath(__file__))
                    #plot_filename = os.path.join(script_directory, "plot.html")
                   
                    # Save plot as HTML file
                    pio.write_html(fig, plot_title +'.html')
                   
                       #plot_filename = (plot_title + ".html")
                       #fig.write_html(plot_filename)
                       
                    print(f"Plot saved as an HTML file here : {script_directory}")
    
                    


            except Exception as e:
                sg.popup_error(f"An error occurred while plotting the data:\n{str(e)}")       

window.close()
