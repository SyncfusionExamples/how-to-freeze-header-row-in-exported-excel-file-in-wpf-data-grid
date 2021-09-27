# How to freeze header row in exported excel file in WPF DataGrid (SfDataGrid) ?

This sample show cases how to freeze header row in exported excel file in [WPF DataGrid](https://www.syncfusion.com/wpf-ui-controls/datagrid) (SfDataGrid)?

# About the sample

You can freeze the header row in the exported Excel sheet using FreezePanes method in `IRange` interface in [WPF DataGrid](https://www.syncfusion.com/wpf-ui-controls/datagrid) (SfDataGrid).

```c#
//Select freeze pane range
//To freeze a row or column, the selected range should be next to the row or column.
IRange range = worksheet[2, 1];
//Create freeze pane in first row
range.FreezePanes();
```
Note: Once you run the sample exported excel file saved inside the bin folder.

KB article - [How to freeze header row in exported excel file in WPF DataGrid (SfDataGrid)?](https://www.syncfusion.com/kb/12007/how-to-freeze-header-row-in-exported-excel-file-in-wpf-datagrid-sfdatagrid)

## Requirements to run the demo
 Visual Studio 2015 and above versions
