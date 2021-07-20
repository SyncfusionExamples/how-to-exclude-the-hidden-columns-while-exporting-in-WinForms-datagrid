﻿using Syncfusion.Data;
using Syncfusion.WinForms.DataGrid;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Forms;
using Syncfusion.WinForms.DataGridConverter;
using Syncfusion.XlsIO;
using System.Diagnostics;

namespace SfDataGridDemo
{
    /// <summary>
    /// Summary description for Form1.
    /// </summary>
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            sfDataGrid.AutoGenerateColumns = false;
            sfDataGrid.DataSource = new ViewModel().Orders;
            sfDataGrid.LiveDataUpdateMode = LiveDataUpdateMode.AllowDataShaping;
            sfDataGrid.AllowEditing = true;

            GridNumericColumn gridTextColumn1 = new GridNumericColumn() { MappingName = "OrderID", HeaderText = "Order ID" };
            GridTextColumn gridTextColumn2 = new GridTextColumn() { MappingName = "CustomerID", HeaderText = "Customer ID" };
            GridTextColumn gridTextColumn3 = new GridTextColumn() { MappingName = "CustomerName", HeaderText = "Customer Name" ,Visible = false };
            GridTextColumn gridTextColumn4 = new GridTextColumn() { MappingName = "Country", HeaderText = "Country" };
            GridTextColumn gridTextColumn5 = new GridTextColumn() { MappingName = "ShipCity", HeaderText = "Ship City" };
            GridCheckBoxColumn checkBoxColumn = new GridCheckBoxColumn() { MappingName = "IsShipped", HeaderText = "Is Shipped" };

            sfDataGrid.Columns.Add(gridTextColumn1);
            sfDataGrid.Columns.Add(gridTextColumn2);
            sfDataGrid.Columns.Add(gridTextColumn3);
            sfDataGrid.Columns.Add(gridTextColumn4);
            sfDataGrid.Columns.Add(gridTextColumn5);
            sfDataGrid.Columns.Add(checkBoxColumn);

            btnExportExcel.Click += BtnExportExcel_Click;
        }

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
    }

    public class OrderInfo : INotifyPropertyChanged
    {
        decimal? orderID;
        string customerId;
        string country;
        string customerName;
        string shippingCity;
        bool isShipped;

        public OrderInfo()
        {

        }

        public decimal? OrderID
        {
            get { return orderID; }
            set { orderID = value; this.OnPropertyChanged("OrderID"); }
        }

        public string CustomerID
        {
            get { return customerId; }
            set { customerId = value; this.OnPropertyChanged("CustomerID"); }
        }

        public string CustomerName
        {
            get { return customerName; }
            set { customerName = value; this.OnPropertyChanged("CustomerName"); }
        }

        public string Country
        {
            get { return country; }
            set { country = value; this.OnPropertyChanged("Country"); }
        }

        public string ShipCity
        {
            get { return shippingCity; }
            set { shippingCity = value; this.OnPropertyChanged("ShipCity"); }
        }

        public bool IsShipped
        {
            get { return isShipped; }
            set { isShipped = value; this.OnPropertyChanged("IsShipped"); }
        }


        public OrderInfo(decimal? orderId, string customerName, string country, string customerId, string shipCity, bool isShipped)
        {
            this.OrderID = orderId;
            this.CustomerName = customerName;
            this.Country = country;
            this.CustomerID = customerId;
            this.ShipCity = shipCity;
            this.IsShipped = isShipped;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
                this.PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class ViewModel
    {
        private ObservableCollection<OrderInfo> orders;
        public ObservableCollection<OrderInfo> Orders
        {
            get { return orders; }
            set { orders = value; }
        }

        public ViewModel()
        {
            orders = new ObservableCollection<OrderInfo>();
            orders.Add(new OrderInfo(1001, "Thomas Hardy", "Germany", "ALFKI", "Berlin", true));
            orders.Add(new OrderInfo(1002, "Laurence Lebihan", "Mexico", "ANATR", "Mexico", false));
            orders.Add(new OrderInfo(1003, "Antonio Moreno", "Mexico", "ANTON", "Mexico", true));
            orders.Add(new OrderInfo(1004, "Thomas Hardy", "UK", "AROUT", "London", true));
            orders.Add(new OrderInfo(1005, "Christina Berglund", "Sweden", "BERGS", "Lula", false));
        }
    }
}
