import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import mpld3
import webbrowser
import threading

# Get the current directory of the Python script
current_directory = os.path.dirname(os.path.abspath(__file__))

# Path to the Excel file in the same folder as the Python script
excel_file_path = os.path.join(current_directory, 'ImputExcel.xlsx')

# Load the Excel data into a pandas DataFrame
df = pd.read_excel(excel_file_path)

# Convert the 'Start' and 'End' columns to datetime objects
df['Start'] = pd.to_datetime(df['Starttime'])
df['End'] = pd.to_datetime(df['endtime'])

# Sort the DataFrame by the 'Start' date
df.sort_values(by='Start', inplace=True)

# Set up the figure and axis
fig, ax = plt.subplots(figsize=(12, 8))  # Adjust the width and height as needed

# Create a dictionary to map group names (tag_id) to y-axis positions
groups = df['tag_id'].unique()
y_positions = range(len(groups))
group_dict = dict(zip(groups, y_positions))

# Increase the height of the task bars
bar_height = 0.8  # Adjust the height as needed

# Plot the Gantt bars for each task
for index, row in df.iterrows():
    task = row['tag_id']  # Replace 'tag_id' with the correct column name
    start = mdates.date2num(row['Start'])
    end = mdates.date2num(row['End'])
    group = row['tag_id']
    y = group_dict[group]
    ax.barh(y=y, left=start, width=end - start, height=bar_height, align='center', label='')

# Format the date and time ticks on the x-axis
date_format = mdates.DateFormatter('%Y-%m-%d %H:%M:%S')
ax.xaxis.set_major_formatter(date_format)
plt.xlabel('Date and Time')

# Set the y-axis labels to display group names
plt.yticks(y_positions, groups)
plt.ylabel('Group')

# Rotate the x-axis labels to prevent overlapping
plt.xticks(rotation=45)

# Add a legend
plt.legend(loc='center left', bbox_to_anchor=(1, 0.5))

# Convert the plot to an interactive HTML
interactive_html = mpld3.fig_to_html(fig)

# Create a function to serve the interactive HTML content in a web page
def serve_html():
    html_content = f'''
    <!DOCTYPE html>
    <html>
    <head>
        <title>Interactive Gantt Chart</title>
        <script src="https://mpld3.github.io/js/mpld3.v0.5.1.dev1.min.js"></script>
    </head>
    <body>
        <h1>Interactive Gantt Chart</h1>
        <div id="gantt_chart_div">
            {interactive_html}
        </div>
    </body>
    </html>
    '''
    with open('gantt_chart.html', 'w') as f:
        f.write(html_content)
    
    # Open the HTML file in the default web browser
    webbrowser.open('gantt_chart.html')

# Run the web server in a separate thread
server_thread = threading.Thread(target=serve_html)
server_thread.start()

# Show the plot (optional, you can comment this line if you only want to display in the web browser)
plt.show()
