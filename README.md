# Plotly & Powerpoint

A library used to create powerpoint slides including plotly charts. Use this to automate powerpoint creation including certiain charts/visualizations.

## Getting Started

Below is a quick tutorial taking you through how to setup and use this library. I will use public data to showcase how it works.

### Step 1 - Install Dependencies
I have yet to figure out how to install required packages with the install of this library. Therefore, please ensure you have installed the following packages. For plotly, please refer to the [getting started](https://plotly.com/python/getting-started/) page to learn how to install everything properly. 
- pandas
- plotly.express
- plotly.graph_objects
- plotly.subplots
- numerize
- scipy.stats

Additionally, if you want to be able to visualize plotly charts, you may have to install additonal requirements. Refer to the plotly getting started page and scroll down for your proper IDE (Jupyter Notebook, Lab, etc.).

### Step 2 - Install Package
    pip install plotlyPowerpoint

### Step 3 - Prepare Your Powerpoint
The main function of this library is meant to generate slides which include a plotly chart. In order to do this, you must prepare a powerpoint template with the proper layout. You can design any layout you desire, but make sure that each element you desire to fill with python is created as a placeholder. You can use a powerpoint template I've included in my source code to make things easy (assets/powerpoint_templates/template.pptx).

In order to do this, open up your desired powerpoint and go into the slide master. Insert a new slide or use a current one already created. I recommend to drag this to be first in the order, this will make your life easier later down the road. Now insert the proper elements and arragne them as you please. For the chart, insert an image placeholder and make the apsect ratio similar to a laptop screen. Most images created are in landscape view but not too wide.

Here is an example of what my template looks like.
![](/assets/images/powerpoint_slide_template.jpg)

### Step 4 - Load Library and Prepare Data
If you want to follow along with my tutorial, feel free and copy/paste my code. However, ensure you have `pydataset` installed before you do.
    
    from pptx import Presentation
    import pandas as pd
    from pydataset import data
    import plotlyPowerpoint as pp

    ############
    ## Prepare Data
    ############

    #load datasets
    df = data('InsectSprays')
    df2 = data("JohnsonJohnson")

    #Data transformation
    df['m2'] = df['count'] * 1.1
    df2['year'] = df2['time'].astype(int)
    df2 = df2.groupby(['year']).agg({'JohnsonJohnson': 'mean'}).reset_index()

### Step 5 - Find Your Powerpoint Elements
In order for this library to know where to place each element, you must give it each element index. See the following code to learn how to do this:

    #load presentation
    prs = Presentation("template.pptx")

    #set the slide we are after - the first slide in the master layout is at index 0
    slide = prs.slides.add_slide(prs.slide_layouts[0])

    #print index and name
    for shape in slide.placeholders:
        print('%d %s' % (shape.placeholder_format.idx, shape.name))

Now set each index for the elements you want to insert

    #set global index for each item on template slide
    setItemIndex('slide', 0)
    setItemIndex("title", 0)
    setItemIndex("chart", 10)
    setItemIndex("description", 11)

As of now, this library only supports `slide`, `title`, `chart`, and `description`.


### Step 6 - Set Template
Now that we have our powerpoint template, we need to tell the library where to find this file. Do this with the following:

    pp.setTemplate("path_to_template")

### Step 7 - Define Charts
Here is where the work is done. Creating slides with charts is done by defining an array of dictionary objects. Each object represents one slide. As a start, I will define two basic charts:

    charts = [
        { #Line Chart - stock prices
            "data": df2,
            "type": "line",
            "name": "Stock Prices by Company",
            "filename": 'charts/stock-prices-by-company',
            "metrics": [
                {"name": "JohnsonJohnson", "prettyName": "Stock Price", "method": "mean"}
            ],
            "axis": "year",
            "x-axis-title": 'Year',
            "y-axis-title": "Average Stock Price",
            "description": "Grouping by additional variables is easy",
            "filters": [
                {"variable": "year", "operation": ">=", "value": "1970", "type":"int"}
            ]
        },
        { #Bar chart of insect sprays
            "data": df,
            "type": "bar",
            "name": "Avg Spray Effictiveness by Type",
            "filename": 'charts/spray-by-type',
            "metrics": [
                {"name": "count", "prettyName": "Effectiveness", "method": "mean"},
                {"name": "m2", "prettyName": "Effectiveness 2", "method": "mean"}
            ],
            "axis": "spray",
            "x-axis-title": "Effectiveness",
            "size": "wide",
            "description": "this slide has data on it!",
            'options': {
                'orientation': 'horizontal',
                'color-grouping': 'metric'
            }
        }
    ]

I will go into more detail on each one of these variables, but for now just note I am including a folder in my filename. I suggest storing all of your images in a separate folder, but if you do, make sure you create it in the root of your project folder.

### Step 8 - Run Function
    #run function
    pp.createSlides(charts)

This will output a file called `output.pptx`. I suggest you do not make this your final file, as each time you run this function you will overwrite the powerpoint. Create a copy of this file and start to create your analysis/report there. From here, you can use `output.pptx` as a slide library, where you can include or delete any chart you create. 

## Documentation