# How to freeze header row in exported excel file in WPF DataGrid (SfDataGrid) ?

How to freeze header row in exported excel file in WPF DataGrid (SfDataGrid) ?

# About the sample

In SfDataGrid, you can freeze the header row in the exported Excel sheet using FreezePanes method in IRange interface.

```c#
//Select freeze pane range
//To freeze a row or column, the selected range should be next to the row or column.
IRange range = worksheet[2, 1];
//Create freeze pane in first row
range.FreezePanes();
```
Note: Once you run the sample exported excel file saved inside the bin folder.

## Requirements to run the demo
 Visual Studio 2015 and above versions
