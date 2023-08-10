import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as sr
import datetime
from datetime import timedelta

# Setting the Streamlit page configuration with title, icon and layout
sr.set_page_config(page_title = "Dashboard",
                   page_icon = ":chart_with_upwards_trend:",
                   layout = "wide"
)

# Function to read power data from the Excel file and return it as a Dataframe with specific settings
@sr.cache_data(experimental_allow_widgets = True)
def get_data_from_excel():
    df = pd.read_excel(
        io = 'C:/Users/soumarya.samanta/.spyder-py3/2022_05_13_HourlyPowerData.xlsx',
        engine = 'openpyxl',
        sheet_name = 'Sheet3',
        skiprows = 0,
        usecols = 'B:L',
        nrows = 87649
    )
    return df

# Function to format energy values into appropriate units and scientific notation if necessary
def format_energy(value):
    units = ['MWh', 'GWh', 'TWh', 'ZWh']
    idx = 0
    while abs(value) >= 1000 and idx < len(units) - 1:
        value /= 1000.0
        idx += 1
    energy = f"{value}"
    if 'e' in energy:
        energy = f"{value:.3e} {units[idx]}"
        energy = energy.replace('e', ' x 10^')
    else:
        energy = f"{value:.3f} {units[idx]}"
    return energy

# Function to create the Daily Representation tab of the Streamlit application
def daily():
    df = get_data_from_excel() # Getting data from Excel
    
    # Creating a sidebar with options to select a specific day and hour range
    sr.sidebar.markdown("<p style = 'font-weight: bold; font-size: 20px;'>Select Day:</p>", unsafe_allow_html = True)
    date = sr.sidebar.date_input(
        " ",
        label_visibility = "collapsed",
        min_value = datetime.date(2010, 1, 1),
        max_value = datetime.date(2019, 12, 31),
        value = datetime.date(2010, 1, 1)
    )
    
    day_of_year = date.timetuple().tm_yday # Calculating the day of the year
    
    # Calculating the slct value based on the year, which determines the data range to be selected
    if ((int(date.year)%10) <= 2):
        slct = ((int(date.year)%10)*365)+day_of_year
    elif (((int(date.year)%10) > 2) and ((int(date.year)%10) <= 6)):
        slct = (((int(date.year)%10)*365)+1)+day_of_year
    else:
        slct = (((int(date.year)%10)*365)+2)+day_of_year
    
    # Filtering the DataFrame based on the selected day
    df_selection = df.query(
        "Day == @slct"
    )
    
    # Creating a sidebar with options to select an hour range
    sr.sidebar.markdown("<p style = 'font-weight: bold; font-size: 20px;'>Select Hour Range:</p>", unsafe_allow_html = True)
    min_hour, max_hour = sr.sidebar.slider(
        " ",
        label_visibility = "collapsed",
        min_value = int(df_selection["Hours"].min()),
        max_value = int(df_selection["Hours"].max()),
        value = (int(df_selection["Hours"].min()), int(df_selection["Hours"].max())),
        step = 1
    )
    
    df_filtered = df_selection[(df_selection["Hours"] >= min_hour) & (df_selection["Hours"] <= max_hour)] # Fitering the DataFrame based on the selected hour range
    
    # Separating the positive and negative values of BESS
    positive_values = df_filtered[df_filtered['BESS'] > 0]['BESS']
    negative_values = df_filtered[df_filtered['BESS'] < 0]['BESS']
    positive_sum = positive_values.sum()
    negative_sum = negative_values.sum()
    
    #Create sidebars with options to select generation, load and total data types
    sr.sidebar.markdown("<p style = 'font-weight: bold; font-size: 20px;'>Select Generation Type:</p>", unsafe_allow_html = True)
    y1_variables = sr.sidebar.multiselect(
        " ",
        label_visibility = "collapsed",
        options = ['PV', 'Wind', 'Grid', 'BESS']
    )
    
    sr.sidebar.markdown("<p style = 'font-weight: bold; font-size: 20px;'>Select Load Type:</p>", unsafe_allow_html = True)
    y2_variables = sr.sidebar.multiselect(
        " ",
        label_visibility = "collapsed",
        options = ['Plant', 'Electrolyzer']
    )
    
    sr.sidebar.markdown("<p style = 'font-weight: bold; font-size: 20px;'>Select Total Data Type:</p>", unsafe_allow_html = True)
    y3_variables = sr.sidebar.multiselect(
        " ",
        label_visibility = "collapsed",
        options = ['Generation', 'Load']
    )
    
    # Calculating total energy generation, total energy load and total energy for each selected data type
    total_energy_generation = {}
    total_energy_load = {}
    total_energy = {}
    for y_variable in y1_variables:
        energy_generation = df_filtered[(df_filtered["Hours"] >= min_hour) & (df_filtered["Hours"] <= max_hour)][y_variable].sum()
        total_energy_generation[y_variable] = energy_generation
    
    for y_variable in y2_variables:
        energy_load = df_filtered[(df_filtered["Hours"] >= min_hour) & (df_filtered["Hours"] <= max_hour)][y_variable].sum()
        total_energy_load[y_variable] = energy_load
    
    for y_variable in y3_variables:
        energy = df_filtered[(df_filtered["Hours"] >= min_hour) & (df_filtered["Hours"] <= max_hour)][y_variable].sum()
        total_energy[y_variable] = energy
    
    # Calculating total energy sum for each selected data type
    total_energy_generation_sum = 0
    total_energy_load_sum = 0
    for y_variable in y1_variables:
        if y_variable in total_energy_generation:
            total_energy_generation_sum += total_energy_generation[y_variable]
    
    for y_variable in y2_variables:
        if y_variable in total_energy_load:
            total_energy_load_sum += total_energy_load[y_variable]
    
    sr.title(":chart_with_upwards_trend: Daily Data") # Defining a title for the data visualization
    
    sr.markdown("----") # Adding a horizontal line
    
    # Creating the graph for generation data
    fig_generation = px.line(
        df_filtered,
        x = "Hours",
        y = y1_variables,
        markers = True,
        color_discrete_sequence = ["#0083B8","#5F94AE","#FFB752","#F9584B"],
        template = "plotly_white"
    )
    
    # Updating the layout of the graph
    fig_generation.update_layout(
        title = dict(text = "<b>Generation Graph</b>", font_size = 24),
        xaxis = dict(tickmode = "linear"),
        plot_bgcolor = "rgba(0,0,0,0)",
        yaxis = (dict(title = "Value (MW)", showgrid = False))
    )
    
    sr.plotly_chart(fig_generation, use_container_width = True) # Displaying the graph
    
    # Displaying selected generation energy and individual energy values for each selected generation type
    sr.markdown(f"<h5>Selected Generation Energy: {format_energy(total_energy_generation_sum)}</h5>", unsafe_allow_html = True)
    for y1_variables, energy_generation in total_energy_generation.items():
        if y1_variables == 'BESS':
            sr.markdown(f"**{y1_variables} Energy**: {format_energy(energy_generation)} ({format_energy(abs(negative_sum))} Charging, {format_energy(positive_sum)} Discharging)")
        else:
            sr.markdown(f"**{y1_variables} Energy**: {format_energy(energy_generation)}")
            
    sr.markdown("####") # Adding a separator
    
    # Creating the graph for load data
    fig_load = px.line(
        df_filtered,
        x = "Hours",
        y = y2_variables,
        markers = True,
        color_discrete_sequence = ["#8B7335","#FFD12B"],
        template = "plotly_white"
    )
    
    # Updating the layout of the graph
    fig_load.update_layout(
        title = dict(text = "<b>Load Graph</b>", font_size = 24),
        xaxis = dict(tickmode = "linear"),
        plot_bgcolor = "rgba(0,0,0,0)",
        yaxis = (dict(title = "Value (MW)", showgrid = False))
    )
    
    sr.plotly_chart(fig_load, use_container_width = True) # Displaying the graph
    
    # Dislaying selected load energy and individual energy values for each selected load type
    sr.markdown(f"<h5>Selected Load Energy: {format_energy(total_energy_load_sum)}</h5>", unsafe_allow_html = True)
    for y2_variables, energy_load in total_energy_load.items():
        sr.markdown(f"**{y2_variables} Energy**: {format_energy(energy_load)}")
        
    sr.markdown("####") # Adding a separator
    
    # Creating the graph for total generation and total load data
    fig_parallel = px.line(
        df_filtered,
        x = "Hours",
        y = y3_variables,
        markers = True,
        color_discrete_sequence = ["#F4D03F","#E74C3C"],
        template = "plotly_white"
    )
    
    # Updating the layout of the graph
    fig_parallel.update_layout(
        title = dict(text = "<b>Total Generation & Total Load Graph</b>", font_size = 24),
        xaxis = dict(tickmode = "linear"),
        plot_bgcolor = "rgba(0,0,0,0)",
        yaxis = (dict(title = "Value (MW)", showgrid = False))
    )
    
    sr.plotly_chart(fig_parallel,use_container_width = True) # Displaying the graph
    
    # Displaying total energy for each selected total data type
    for y3_variables, energy in total_energy.items():
        sr.markdown(f"**{y3_variables} Total Energy**: {format_energy(energy)}")
    
    # Hiding Streamlit's main menu, header and footer for a cleaner UI
    hide_sr_style = """
                <style>
                #MainMenu {visibility: hidden;}
                footer {visibility: hidden;}
                header {visibility: hidden;}
                </style>
    """
    sr.markdown(hide_sr_style, unsafe_allow_html = True)

# Function to create the Custom Date Range Representation tab of the Streamlit application
def custom_date():
    df = get_data_from_excel() # Getting data from Excel
    
    # Creating sidebars with options to select a custom date range
    sr.sidebar.markdown("<p style = 'font-weight: bold; font-size: 20px;'>Select Start Date:</p>", unsafe_allow_html = True)
    start_date = sr.sidebar.date_input(
        label = " ",
        label_visibility = "collapsed",
        min_value = datetime.date(2010, 1, 1),
        max_value = datetime.date(2019, 12, 30),
        value = datetime.date(2010, 1, 1)
    )
    
    sr.sidebar.markdown("<p style = 'font-weight: bold; font-size: 20px;'>Select End Date:</p>", unsafe_allow_html = True)
    end_date = sr.sidebar.date_input(
        label = " ",
        label_visibility = "collapsed",
        min_value = datetime.date(2010, 1, 2),
        max_value = datetime.date(2019, 12, 31),
        value = datetime.date(2010, 1, 2)
    )
    
    # Calculating the day of the year for the start and end dates
    start_day_of_year = start_date.timetuple().tm_yday
    end_day_of_year = end_date.timetuple().tm_yday
    
    # Calculating the slct value for the start and end dates based on the year
    if ((int(start_date.year)%10) <= 2):
        start_slct = ((int(start_date.year)%10)*365)+start_day_of_year
    elif (((int(start_date.year)%10) > 2) and ((int(start_date.year)%10) <= 6)):
        start_slct = (((int(start_date.year)%10)*365)+1)+start_day_of_year
    else:
        start_slct = (((int(start_date.year)%10)*365)+2)+start_day_of_year
    
    if ((int(end_date.year)%10) <= 2):
        end_slct = ((int(end_date.year)%10)*365)+end_day_of_year
    elif (((int(end_date.year)%10) > 2) and ((int(end_date.year)%10) <= 6)):
        end_slct = (((int(end_date.year)%10)*365)+1)+end_day_of_year
    else:
        end_slct = (((int(end_date.year)%10)*365)+2)+end_day_of_year
    
    # Filtering the DataFrame based on the selected date range
    df_selection = df.query(
        "Day >= @start_slct and Day <= @end_slct"
    )
    
    # Separating positive and negative values of BESS
    positive_values = df_selection[df_selection['BESS'] > 0]['BESS']
    negative_values = df_selection[df_selection['BESS'] < 0]['BESS']
    positive_sum = positive_values.sum()
    negative_sum = negative_values.sum()
    
    # Creating sidebars with options to select generation, load and total data types
    sr.sidebar.markdown("<p style = 'font-weight: bold; font-size: 20px;'>Select Generation Type:</p>", unsafe_allow_html = True)
    y1_variables = sr.sidebar.multiselect(
        " ",
        label_visibility = "collapsed",
        options = ['PV', 'Wind', 'Grid', 'BESS']
    )
    
    sr.sidebar.markdown("<p style = 'font-weight: bold; font-size: 20px;'>Select Load Type:</p>", unsafe_allow_html = True)
    y2_variables = sr.sidebar.multiselect(
        " ",
        label_visibility = "collapsed",
        options = ['Plant', 'Electrolyzer']
    )
    
    sr.sidebar.markdown("<p style = 'font-weight: bold; font-size: 20px;'>Select Total Data Type:</p>", unsafe_allow_html = True)
    y3_variables = sr.sidebar.multiselect(
        " ",
        label_visibility = "collapsed",
        options = ['Generation', 'Load']
    )
    
    # Calculating total energy generation, total energy load and total energy for each selected data type
    total_energy_generation = {}
    total_energy_load = {}
    total_energy = {}
    for y_variable in y1_variables:
        energy_generation = df_selection.groupby("Day")[y_variable].sum().sum()
        total_energy_generation[y_variable] = energy_generation
    
    for y_variable in y2_variables:
        energy_load = df_selection.groupby("Day")[y_variable].sum().sum()
        total_energy_load[y_variable] = energy_load
    
    for y_variable in y3_variables:
        energy = df_selection.groupby("Day")[y_variable].sum().sum()
        total_energy[y_variable] = energy
    
    # Calculating total energy sum for each selected data type
    total_energy_generation_sum = 0
    total_energy_load_sum = 0
    for y_variable in y1_variables:
        if y_variable in total_energy_generation:
            total_energy_generation_sum += total_energy_generation[y_variable]
    
    for y_variable in y2_variables:
        if y_variable in total_energy_load:
            total_energy_load_sum += total_energy_load[y_variable]
    
    sr.title(":chart_with_upwards_trend: Custom Date Range Data") # Defining a title for the data visualization
    
    sr.markdown("----") # Adding a horizontal line
    
    fig_generation = go.Figure() # Creating the graph for generation data
    
    colors = ["#A93226","#6C3483","#0E6655","#F1C40F","#D35400","#616A6B","#2C3E50","#2ECC71"] # Defining a list of colors to be used for different lines in the graphs
    
    # Grouping the data by day and calculating maximum and minimum values
    max_values = df_selection.groupby("Day").max().reset_index()
    min_values = df_selection.groupby("Day").min().reset_index()
    
    # Converting day numbers to actual dates for x-axis ticks
    start_date_datetime = datetime.date(int(start_date.year), 1, 1) + timedelta(days = start_day_of_year - 1)
    end_date_datetime = datetime.date(int(end_date.year), 1, 1) + timedelta(days = end_day_of_year - 1)
    
    x_ticks = pd.date_range(start_date_datetime, end_date_datetime).strftime('%d/%m/%y').tolist() # Creating custom x-axis ticks with formatted dates in 'DD/MM/YY' format
    
    for i, y_variable in enumerate(y1_variables): # Iterating through the different generation sources
    
        # Calculating the maximum and minimum values for the current generation source
        max_values = df_selection.groupby("Day")[y_variable].max().reset_index()
        min_values = df_selection.groupby("Day")[y_variable].min().reset_index()
        
        # Adding a trace for the maximum values to the graph
        fig_generation.add_trace(go.Line(
            x = max_values["Day"],
            y = max_values[y_variable],
            mode = 'lines+markers',
            line = dict(color = colors[i % len(colors)]),
            name = f"{y_variable} (Max)")
        )
        
        # Adding a trace for the minimum values to the graph with a dashed line
        fig_generation.add_trace(go.Line(
            x = min_values["Day"],
            y = min_values[y_variable],
            mode = 'lines+markers',
            line = dict(color = colors[(i + 2) % len(colors)], dash = "dash"),
            name = f"{y_variable} (Min)")
        )
    
    # Updating the layout of the graph
    fig_generation.update_layout(
        title = dict(text = "<b>Generation Graph</b>", font_size = 24),
        xaxis = dict(title = "Days", tickmode = "array", ticktext = x_ticks, tickvals = max_values["Day"]),
        yaxis = dict(title = "Value (MW)", showgrid = False),
        template = "plotly_white",
        showlegend = True
    )
    
    sr.plotly_chart(fig_generation, use_container_width = True) # Displaying the graph
    
    # Displaying selected generation energy and individual energy values for each selected generation type
    sr.markdown(f"<h5>Selected Generation Energy: {format_energy(total_energy_generation_sum)}</h5>", unsafe_allow_html = True)
    for y1_variables, energy_generation in total_energy_generation.items():
        if y1_variables == 'BESS':
            sr.markdown(f"**{y1_variables} Energy**: {format_energy(energy_generation)} ({format_energy(abs(negative_sum))} Charging, {format_energy(positive_sum)} Discharging)")
        else:
            sr.markdown(f"**{y1_variables} Energy**: {format_energy(energy_generation)}")
            
    sr.markdown("####") # Adding a separator
    
    fig_load = go.Figure() # Creating the graph for load data
    
    colors = ["#A93226","#6C3483","#0E6655","#F1C40F","#D35400","#616A6B","#2C3E50","#2ECC71"] # Defining a list of colors to be used for different lines in the graphs
    
    # Grouping the data by day and calculating maximum and minimum values
    max_values = df_selection.groupby("Day").max().reset_index()
    min_values = df_selection.groupby("Day").min().reset_index()
    
    for i, y_variable in enumerate(y2_variables): # Iterating through the different load sources
    
        # Calculating the maximum and minimum values for the current load source
        max_values = df_selection.groupby("Day")[y_variable].max().reset_index()
        min_values = df_selection.groupby("Day")[y_variable].min().reset_index()
        
        # Adding a trace for the maximum values to the graph
        fig_load.add_trace(go.Line(
            x = max_values["Day"],
            y = max_values[y_variable],
            mode = 'lines+markers',
            line = dict(color = colors[i % len(colors)]),
            name = f"{y_variable} (Max)")
        )
        
        # Adding a trace for the minimum values to the graph with a dashed line
        fig_load.add_trace(go.Line(
            x = min_values["Day"],
            y = min_values[y_variable],
            mode = 'lines+markers',
            line = dict(color = colors[(i + 2) % len(colors)], dash = "dash"),
            name = f"{y_variable} (Min)")
        )
    
    # Updating the layout of the graph
    fig_load.update_layout(
        title = dict(text = "<b>Load Graph</b>", font_size = 24),
        xaxis = dict(title = "Days", tickmode = "array", ticktext = x_ticks, tickvals = max_values["Day"]),
        yaxis = dict(title = "Value (MW)", showgrid = False),
        template = "plotly_white",
        showlegend = True
    )
    
    sr.plotly_chart(fig_load, use_container_width = True) # Displaying the graph
    
    # Dislaying selected load energy and individual energy values for each selected load type
    sr.markdown(f"<h5>Selected Load Energy: {format_energy(total_energy_load_sum)}</h5>", unsafe_allow_html = True)
    for y2_variables, energy_load in total_energy_load.items():
        sr.markdown(f"**{y2_variables} Energy**: {format_energy(energy_load)}")
        
    sr.markdown("####") # Adding a separator
    
    fig_parallel = go.Figure() # Creating the graph for total generation and load data
    
    colors = ["#A93226","#6C3483","#0E6655","#F1C40F","#D35400","#616A6B","#2C3E50","#2ECC71"] # Defining a list of colors to be used for different lines in the graphs
    
    # Grouping the data by day and calculating maximum and minimum values
    max_values = df_selection.groupby("Day").max().reset_index()
    min_values = df_selection.groupby("Day").min().reset_index()
    
    for i, y_variable in enumerate(y3_variables): # Iterating through the different total energy sources
    
        # Calculating the maximum and minimum values for the current total energy source
        max_values = df_selection.groupby("Day")[y_variable].max().reset_index()
        min_values = df_selection.groupby("Day")[y_variable].min().reset_index()
        
        # Adding a trace for the maximum values to the graph
        fig_parallel.add_trace(go.Line(
            x = max_values["Day"],
            y = max_values[y_variable],
            mode = 'lines+markers',
            line = dict(color = colors[(i + 4) % len(colors)]),
            name = f"{y_variable} (Max)")
        )
        
        # Adding a trace for the minimum values to the graph with a dashed line
        fig_parallel.add_trace(go.Line(
            x = min_values["Day"],
            y = min_values[y_variable],
            mode = 'lines+markers',
            line = dict(color = colors[(i + 6) % len(colors)], dash = "dash"),
            name = f"{y_variable} (Min)")
        )
    
    # Updating the layout of the graph
    fig_parallel.update_layout(
        title = dict(text = "<b>Total Generation & Total Load Graph</b>", font_size = 24),
        xaxis = dict(title = "Days", tickmode = "array", ticktext = x_ticks, tickvals = max_values["Day"]),
        yaxis = dict(title = "Value (MW)", showgrid = False),
        template = "plotly_white",
        showlegend = True
    )
    
    sr.plotly_chart(fig_parallel, use_container_width = True) # Displaying the graph
    
    # Displaying total energy for each selected total data type
    for y3_variables, energy in total_energy.items():
        sr.markdown(f"**{y3_variables} Total Energy**: {format_energy(energy)}")
    
    # Hiding Streamlit's main menu, header and footer for a cleaner UI
    hide_sr_style = """
                <style>
                #MainMenu {visibility: hidden;}
                footer {visibility: hidden;}
                header {visibility: hidden;}
                </style>
    """
    sr.markdown(hide_sr_style, unsafe_allow_html=True)

# Dictionary to map tab names to their corresponding functions
tabs = {
        "Daily Representation": daily,
        "Custom Date Range Representation": custom_date,
}

default_tab = "Daily Representation" # Default tab to be shown when the application is loaded

# Creating a sidebar for user interaction, allowing users to choose the tab
sr.sidebar.markdown("<h3 style = 'font-size:20px; font-weight:bold;'>Select Tab:</h3>", unsafe_allow_html = True)
selected_tab = sr.sidebar.radio(
    " ",
    label_visibility = "collapsed",
    options = list(tabs.keys()),
    index = list(tabs.keys()).index(default_tab)
)

tabs[selected_tab]() # Calling the selected tab function to dislay its content