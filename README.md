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

#### Example 1

https://user-images.githubusercontent.com/21944548/166474097-691fa83a-5817-4022-96e0-2b60d790b6ab.mov

Creating a XY Scatter plot with two series in two steps.

#### Example 2

https://user-images.githubusercontent.com/21944548/166473729-fa9de08a-988b-4322-8163-87fc97de06f6.mov

Creating a XY scatter plot with two series in one step.

#### Example 3

https://user-images.githubusercontent.com/21944548/166475251-741407b2-4e89-4609-a30f-c5aae32d5115.mov


Description

### XMean

#### Input


#### Notes

#### Example



https://user-images.githubusercontent.com/21944548/166475959-27c507e2-c9e1-4ff5-825a-4ae359ac8f29.mov


#### Format markers default



https://user-images.githubusercontent.com/21944548/166479547-003a0cbc-b5c7-4214-bec0-9b45b3fd348d.mov


#### Format markers custom
TODO Rename icon in ribbon


https://user-images.githubusercontent.com/21944548/166479574-74d0e475-a404-45bc-85bd-2ce34fd02c99.mov



https://user-images.githubusercontent.com/21944548/166479600-90bfa6f7-399c-48f4-90f2-703e6a31c5f7.mov


#### Format axis


https://user-images.githubusercontent.com/21944548/166480441-eb2ce5ac-36b8-48d1-aeaa-d544d5dbb8eb.mov


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
