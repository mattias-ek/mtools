# mtools
A VBA plugin for Excel to make plots and import files

## Usage
You can install the MTools either as a plugin which will be avaliable for all your excel files or you can use *MTools.xlsm* as a template where the ribbon only appears for that file. In the latter case ensure that you have macros enabled.

### Installing MTools as a plugin
First dowmload the *MTools.xlam* file and save it somewhere sensible on your harddrive. Then in Excel go to *File -> Options -> Add-ins -> Manage: Excel Add-ins Go...* then click *Browse* and select the *MTools.xlam* file. To disable the plugin untick the plugin in this dialog. See the [offical documentation](https://support.microsoft.com/en-us/office/add-or-remove-add-ins-in-excel-0af570c4-5cf3-4fa9-9b88-403625a0b460) for additional assistance.

## Macros
  - [Plot](#Plot)
  - [XMean](#XMean)
  - [Format markers default](#Format markers default)


### <img src="https://github.com/mattias-ek/mtools/blob/main/Icons/Scatter.png" width="25" height="25"> Plot
Creates a scatter plot with errorbars. 

If there is an active plot when the macro is executed then the data will be added to that plot.

#### Input dialogs
**Select x values** - If two columns are selected then the values in the second column values are used for the error bars. If the first cell contains text then that will be used as the axis label and ignored as a data point (See Example 2). Non-continous selections are allowed if each selection has the same number of columns.

**Select y values (Optional)** - Same as for Select x values. If canceled then the values will be plotted with at an incremental y value like a Caltech/Carnegie plot (See Example 3).

**Select series name (Optional)** - If a single cell is selected all data points will be plotted under that name (See Example 1). If multiple cells are selected then individual data points will be plotted under the corresponding series name (See Example 2). If canceled then all the data points are plotted as a single series under a default name.

#### Notes
If a new plot is created it will be placed at the top left coordinates of the active cell. If multiple cells are active then the size of the plot will span from the top left of the first selected cell to the bottom right of the last selected cell (See Example 3).

If a series name already exists in the plot it will be replaced by the new data. If no y values are given for an existing series then the postion of other series in the plot will not be updated and there is a chance series will overlap if the number of data points changes. In this case it is best to create a new plot.

#### Example 1
https://user-images.githubusercontent.com/21944548/166474097-691fa83a-5817-4022-96e0-2b60d790b6ab.mov

Creating a XY Scatter plot with two series in two steps.

#### Example 2
https://user-images.githubusercontent.com/21944548/166473729-fa9de08a-988b-4322-8163-87fc97de06f6.mov

Creating a XY scatter plot with two series in one step. Since the first cell in the X and Y selections are text they are used as the axis titles.

#### Example 3

https://user-images.githubusercontent.com/21944548/166475251-741407b2-4e89-4609-a30f-c5aae32d5115.mov

Creating a Caltech/Carnegie plot with incremental y values. The size of the plot is the same as the selected cells.

### <img src="https://github.com/mattias-ek/mtools/blob/main/Icons/XMean.png" width="25" height="25"> XMean
Adds a median line and uncertianty box around an existing series in a Calech/Carnegie plot.

#### Input dialogs
**Select x values** - If two columns are selected then the values in the second column values are used to create a uncertianty box around the series. Non-continous selections are allowed if each selection has the same number of columns.

**Select series name (Optional)** - The series name the x values correspond too. If canceled the x values will be applied to the last series in the plot.

#### Notes
Only use this function on Caltech/Carnegie plots. There is a risk it will crash if there are N/A values in the Y values.

If multiple rows of x values are given and no series name is given then the last row of the x values will be applied to the last series in the graph. Other x values will be ignored.

#### Example
https://user-images.githubusercontent.com/21944548/166475959-27c507e2-c9e1-4ff5-825a-4ae359ac8f29.mov

Adds a vertical line at the mean and a dashed box at 2 SD of the two series in the plot.

### <img src="https://github.com/mattias-ek/mtools/blob/main/Icons/FormatMarkers.png" width="25" height="25"> Format markers default
Formats the markers of series in the active plot with predefined values.

The predefined marker colours are: Blue1, Red, Pink, Green, Yellow, Blue2, Orange (Repeating)
The predefined marker symbols are: Square, Circle, Triangle, Diamond (Repeating)
The predefined marker size is **8**.

#### Notes
The marker + colour combinations will repeat after 28 series.

The colours are taken from [Points of view: Color blindness](https://www.nature.com/articles/nmeth.1618) and are suitable for those with color vision deficiencies.

#### Example
https://user-images.githubusercontent.com/21944548/166479547-003a0cbc-b5c7-4214-bec0-9b45b3fd348d.mov

Formats markers with the default values.

### <img src="https://github.com/mattias-ek/mtools/blob/main/Icons/FormatMarkers.png" width="25" height="25"> Format markers custom
Allows you to format the marker symbol, colour and size of series in a plot.

#### Input dialogs
**Series name (Optional)** - The series name that the marker symbol, colour and/or size should be applied to. If cancelled the marker symbols and colours will be applied sequantially to series in the plot, repeating if necessary.

**Marker symbol (Optional)** - Select the markers to be used. The name of the avaliable marker symbols are listed below. If cancelled the symbols will not be changed.

Avaliable markers are: Square, Circle, Triangle, Diamond, Plus, Star, Dash, Dot, None (No marker)

**Marker colours (Optional)** - Select the colours to be used. Colours can be specified using either the name or number listed below. For alternative colours a RGB code can be given as *R, G, B*. If cancelled the colours will not be changed.

| Colour | Name | Number | RGB | 
|--------|------|--------|-----|
|![](https://via.placeholder.com/15/000000/000000?text=+) | Black | 0 | 0, 0, 0 |
|![](https://via.placeholder.com/15/0072b2/000000?text=+) | Blue, Blue1 | 1 | 0, 114, 178 |
|![](https://via.placeholder.com/15/d55e00/000000?text=+) | Red | 2 | 213, 94, 0 |
|![](https://via.placeholder.com/15/cc79a7/000000?text=+) | Pink | 3 | 204, 121, 167 |
|![](https://via.placeholder.com/15/009e73/000000?text=+) | Green | 4 | 0, 158, 115 |
|![](https://via.placeholder.com/15/f0e442/000000?text=+) | Yellow | 5 | 240, 228, 66 |
|![](https://via.placeholder.com/15/56b4df/000000?text=+) | Blue2 | 6 | 86, 180, 223 |
|![](https://via.placeholder.com/15/e69f00/000000?text=+) | Orange | 7 | 230, 159, 0 |

**Marker size (Optional)** - Size of the markers. If cancelled the size of the markers will not be changed.


#### Notes
The colours are taken from [Points of view: Color blindness](https://www.nature.com/articles/nmeth.1618) and are suitable for those with color vision deficiencies.

The symbol and colour names are not case sensitive.

#### Example 1
https://user-images.githubusercontent.com/21944548/166479574-74d0e475-a404-45bc-85bd-2ce34fd02c99.mov

The marker, colour and size of the given series are changed.

#### Example 2
https://user-images.githubusercontent.com/21944548/166479600-90bfa6f7-399c-48f4-90f2-703e6a31c5f7.mov

Since no series name is given the marker symbols and colours are repeated.

### <img src="https://github.com/mattias-ek/mtools/blob/main/Icons/FormatAxis.png" width="25" height="25"> Format axis

Besides changing the font size, x & y axis labels the macro will change the colour of the axes to black.

#### Input dialogs
**Font size (Optional)** - The font size of the axis labels, axis tick labels and the legend labels. If cancelled the font size is unchanged.

**x axis label (Optional)** - The label used for the x axis. If cancelled the label is unchanged.

**y axis label (Optional)** - The label used for the x axis. If cancelled the label is unchanged.

#### Example
https://user-images.githubusercontent.com/21944548/166480441-eb2ce5ac-36b8-48d1-aeaa-d544d5dbb8eb.mov

Changing the font size and adding x and y axis labels.

### <img src="https://github.com/mattias-ek/mtools/blob/main/Icons/ImportExpData.png" width="25" height="25"> Import EXP data
Imports the data portion of a Neptune/Triton EXP file into the active cell.

### <img src="https://github.com/mattias-ek/mtools/blob/main/Icons/ImportExpFile.png" width="25" height="25"> Import EXP file
Imports the entire content of a Neptune/Triton EXP file into the active cell.

### <img src="https://github.com/mattias-ek/mtools/blob/main/Icons/ImportCSVFile.png" width="25" height="25"> Import CSV file
Imports the entire content of a CSV file into the active cell.

#### Input dialogs
**What delimiter does the file use?** - The delimiter the file uses to seperate values. For *tab* delimited files write TAB (case insensitive).

### <img src="https://github.com/mattias-ek/mtools/blob/main/Icons/ImportExpData.png" width="25" height="25"> Import multiple EXP data
Imports the data portion of one or more Neptune/Triton EXP files into seperate sheets with name of the file.

### <img src="https://github.com/mattias-ek/mtools/blob/main/Icons/ImportExpFile.png" width="25" height="25"> Import multiple EXP files
Imports the entire contents of one or more Neptune/Triton EXP files into seperate sheets with name of the file.

### <img src="https://github.com/mattias-ek/mtools/blob/main/Icons/ImportCSVFile.png" width="25" height="25"> Import multiple CSV files
Imports the entire contents of one or more CSV files into seperate sheets with name of the file.

#### Input dialogs
**What delimiter does the file use?** - The delimiter the file(s) uses to seperate values. For *tab* delimited files write TAB (case insensitive).

#### Help
Opens the github page of MTools which contains the readme file with instructions on the macros in the plugin.

# License

Copyright (c) 2022 Mattias Ek.

Licensed under the [MIT](https://github.com/mattias-ek/mtools/blob/main/LICENSE) license.
