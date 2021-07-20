# How to exclude the Hidden Columns while exporting in WinForms DataGrid (SfDataGrid)?

## About the sample
This example illustrates how to exclude the Hidden Columns while exporting in [WinForms DataGrid](https://www.syncfusion.com/winforms-ui-controls/datagrid) (SfDataGrid)?

[WinForms DataGrid](https://www.syncfusion.com/winforms-ui-controls/datagrid) (SfDataGrid) does not provide the direct support to exclude the hidden columns while exporting. You can exclude the hidden columns while exporting by using [ExcludeColumns](https://help.syncfusion.com/cr/windowsforms/Syncfusion.WinForms.DataGridConverter.ExcelExportingOptions.html#Syncfusion_WinForms_DataGridConverter_ExcelExportingOptions_ExcludeColumns) field in [ExcelExportingOptions](https://help.syncfusion.com/cr/windowsforms/Syncfusion.WinForms.DataGridConverter.ExcelExportingOptions.html).

```C#

private void BtnExportExcel_Click(object sender, EventArgs e)
{
    var file_name = "Sample.xlsx";
    var options = new ExcelExportingOptions
    {
        ExcelVersion = ExcelVersion.Excel2016,
    };

    //get the columns in SfDataGrid
    foreach (var column in sfDataGrid.Columns)
    {
        //check the columns is Visible or not
        if (!column.Visible)
            //While exporting Hidden column stop by Add the MappingName of hidden column in ExcludeColumns in ExcelExportingOptions
            options.ExcludeColumns.Add(column.MappingName);
    }

    var excelEngine = sfDataGrid.ExportToExcel(sfDataGrid.View, options);
    var workBook = excelEngine.Excel.Workbooks[0];
    workBook.SaveAs(file_name);
    _ = Process.Start(file_name);
}

```

![Shows the exclude the hidden column while exporting in SfDataGrid](ExcludeHiddenColumn.gif)

Take a moment to peruse the [WinForms DataGrid - Export To Excel](https://help.syncfusion.com/windowsforms/datagrid/exporttoexcel) documentation, where you can find about export to excel with code examples.

## Requirements to run the demo
Visual Studio 2015 and above versions
