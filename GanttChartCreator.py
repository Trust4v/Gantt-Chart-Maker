import os
import pandas as pd
import plotly.express as px
import random

# Get the current directory of the Python script
current_directory = os.path.dirname(os.path.abspath(__file__))

# Get a list of all files in the "input" folder
input_folder = os.path.join(current_directory, 'input')
input_files = os.listdir(input_folder)

# Check if there are any Excel files in the "input" folder
excel_files = [file for file in input_files if file.endswith('.xlsx')]
if not excel_files:
    raise FileNotFoundError("No Excel file found in the 'input' folder.")

# Assume there is only one Excel file in the folder, so we use the first one
excel_file_name = excel_files[0]
excel_file_path = os.path.join(input_folder, excel_file_name)

# Get the filename without the extension for the chart title
chart_title = os.path.splitext(excel_file_name)[0]

# Convert the 'Starttime' and 'endtime' columns to pandas datetime objects
df = pd.read_excel(excel_file_path)
df['starttime'] = pd.to_datetime(df['starttime'])
df['endtime'] = pd.to_datetime(df['endtime'])

# Sort the DataFrame by the 'Starttime' date
df.sort_values(by='starttime', inplace=True)

# Get unique tasks in 'tag_id' column
unique_tasks = df['tag_id'].unique()

# Define the specific tag_id for which you want to set a particular color
specified_tag_id = 'Cowbrush'

# Create a dictionary to store the color mapping for each unique tag_id
task_colors = {}

# Set a specific color for the specified tag_id and randomize colors for others
for task in unique_tasks:
    if task == specified_tag_id:
        task_colors[task] = 'rgb(255,0,0)'  
    else:
        task_colors[task] = 'rgb(0,0,255)'

# Create DataFrames for specified and other tag_ids
specified_df = df[df['tag_id'] == specified_tag_id]
other_df = df[df['tag_id'] != specified_tag_id]

# Concatenate the DataFrames to move the specified tag_id to the top
df = pd.concat([specified_df, other_df], sort=False)

# Ensure the specified_tag_id is the first element in unique_tasks
unique_tasks = [specified_tag_id] + [task for task in unique_tasks if task != specified_tag_id]

# Calculate the height of the chart based on the number of unique tags
min_height = 200  # Minimum height to avoid excessive scaling for a large number of tags
height = min(min_height + len(unique_tasks) * 50, 1000)  # Adjust the scaling factor (50) as needed

# Create the interactive Gantt chart using Plotly Express
fig = px.timeline(df, x_start='starttime', x_end='endtime', y='tag_id',
                  color='tag_id', color_discrete_map=task_colors,
                  category_orders={'tag_id': unique_tasks})  # Specify the order of y-axis categories

# Update layout to remove the color legend and set the dynamic height
fig.update_layout(
    title=f'Interactive Gantt Chart - {chart_title}',  # Add the chart title with the filename
    xaxis_title='Date and Time',
    yaxis_title='Tag_ID',
    height=height,  # Set the dynamic height
    showlegend=False  # Remove the color legend from the chart
)

# Save the interactive HTML content to 'index.html'
html_file_path = os.path.join(current_directory, 'index.html')
fig.write_html(html_file_path)

# Display the file path to the user for convenience
print(f'Interactive Gantt chart HTML file saved: {html_file_path}')
