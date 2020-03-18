using System;
using System.Collections.Generic;
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
using Microsoft.Win32;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace UpdateDbWpf
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }



        private void LoadExcelFile(string filePath)
        {
            Excel.Application exApp = new Excel.Application();
            Excel.Workbook wb = exApp.Workbooks.Open(filePath);
            if (wb.Worksheets.Count != 1)
                textBlock.Text = "Wrong file";
            else
            {
                textBlock.Text = wb.Worksheets.Count.ToString();
            }
        }

        private void StuffButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "XLSX-files|*.xlsx";
            openFileDialog.ShowDialog();
            //openFileDialog.
            LoadExcelFile(openFileDialog.FileName);
        }

        private void LoadExcelSheetToTable(string fileName, string sheet)
        {
            DataTable table = new DataTable();
            using (System.Data.OleDb.OleDbConnection conn =
                new System.Data.OleDb.OleDbConnection(
                    "Provider=Microsoft.ACE.OLEDB.12.0; " +
                    "Data Source ='" + fileName + "';" +
                    "Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\""))

                using (System.Data.OleDb.OleDbDataAdapter import =
                    new System.Data.OleDb.OleDbDataAdapter(
                        "select * from [" + sheet + "$]", conn))
                    import.Fill(table);
        }
    }
}
