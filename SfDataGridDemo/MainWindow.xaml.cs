using Syncfusion.UI.Xaml.Grid.Converter;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace SfDataGridDemo
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();                   
        }

        private void OnExcelExporting(object sender, RoutedEventArgs e)
        {
            var options = new ExcelExportingOptions();
            options.ExcelVersion = ExcelVersion.Excel2016;
            options.ExportAllPages = true;          
            var excelEngine = dataGrid.ExportToExcel(dataGrid.View, options);
            var workBook = excelEngine.Excel.Workbooks[0];           
            
            // To remove the empty sheets
            for (int i = workBook.Worksheets.Count - 1; i >= 0; i--)
            {
                var worksheet = workBook.Worksheets[i];
                if (worksheet.UsedRange.LastRow == 0 || worksheet.UsedRange.LastColumn == 0)
                {
                    workBook.Worksheets.Remove(i);
                }
            }

            workBook.SaveAs("Export.xlsx");
            Process.Start(new ProcessStartInfo { FileName = "Export.xlsx" });
        }
    }
}
