import os
import pandas as pd
import plotly.express as px

# Get the current directory of the Python script
current_directory = os.path.dirname(os.path.abspath(__file__))

# Path to the Excel file in the same folder as the Python script
excel_file_path = os.path.join(current_directory, 'ImputExcel.xlsx')

# Convert the 'Starttime' and 'endtime' columns to pandas datetime objects
df = pd.read_excel(excel_file_path)
df['Starttime'] = pd.to_datetime(df['Starttime'])
df['endtime'] = pd.to_datetime(df['endtime'])

# Sort the DataFrame by the 'Starttime' date
df.sort_values(by='Starttime', inplace=True)

# Create the interactive Gantt chart using Plotly Express
fig = px.timeline(df, x_start='Starttime', x_end='endtime', y='tag_id', text='tag_id')

# Update layout for better appearance
fig.update_layout(
    title='Interactive Gantt Chart',
    xaxis_title='Date and Time',
    yaxis_title='tag_id',
    height=500,  # Adjust height as needed
)

# Save the interactive HTML content to 'index.html'
html_file_path = os.path.join(current_directory, 'index.html')
fig.write_html(html_file_path)

# Display the file path to the user for convenience
print(f'Interactive Gantt chart HTML file saved: {html_file_path}')
