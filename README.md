# mtools
A VBA plugin for Excel to make plots and import files

## Installation

If you dont want to install MTools as a plugin you can use the 'MTools.xlsm' workbook which has the tools 


## Macros
  - [Plot](#Plot)
  - [XMean](#XMean)


### Plot
Created/Updates a scatter plot with errorbars. If a plot is selected then the data will be added to that plot otherwise a new plot will be created.

#### Input
**Select x values** - If two columns are selected then the values in the second column values are used for the error bars. If the first cell contains text then that will be used as the axis label and ignored as a data point (See Example 2). Non-continous selections are allowed if each selection has the same number of columns.

**Select y values (Optional)** - Same as for Select x values. If canceled then the values will be plotted with at an incremental y value like a Caltech/Carnegie plot (See Example 3).

**Select series name (Optional)** - If a single cell is selected all data points will be plotted under that name (See Example 1). If multiple cells are selected then individual data points will be plotted under the corresponding series name (See Example 2). If canceled then all the data points are plotted as a single series under a default name.

#### Notes
If a new plot is created it will be placed at the top left coordinated of the selected cell. If a range is selected the size of the plot will match the selected range (See Example 3).

If a series name already exists in the plot it will be replaced by the new data. If no y values are given for an existing series then the postion of other series in the plot will not be updated and there is a chance series will overlap if the number of data points changes. In this case it is best to create a new plot instead.

#### Examples

Image
Example 1 - 

Image
Example 2 -

Image
Example 3 -

### XMean

#### Format markers default

#### Format markers custom
TODO Rename icon in ribbon

#### Format axis

#### Import EXP data

#### Import EXP file

#### Import CSV file

#### Import multiple EXP data

#### Import multiple EXP files

#### Import multiple CSV files

#### Help
Opens this page

# License

Copyright (c) 2022 Mattias Ek.

Licensed under the [MIT](https://github.com/mattias-ek/mtools/blob/main/LICENSE) license.
