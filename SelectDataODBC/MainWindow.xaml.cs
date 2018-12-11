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
using System.Data.Odbc;

namespace SelectDataODBC
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string foldername;
        public MainWindow()
        {
            InitializeComponent();
            login.Text = "jmajcher";
            columnsList.Text = "COLUMN1, COLUMN2";
            table.Text = "TEST2";


        }

        private void start_Click(object sender, RoutedEventArgs e)
        {
            string query = "SELECT " + columnsList.Text + " FROM " + table.Text;
            string[] columns = columnsList.Text.Replace(" ","").Split(',');
            int noOfColumns = columns.Length;

            OracleODBC odbc = new OracleODBC();
            string[] data = odbc.GetData(query, login.Text, password.Password.ToString(), noOfColumns);

            GenerateExcel excel = new GenerateExcel();
            excel.GenerateExcelFile(columns, noOfColumns, data, foldername);
        }

        private void exportLink_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            // Create OpenFileDialog 
            var dlg = new System.Windows.Forms.FolderBrowserDialog();
            var result = dlg.ShowDialog().ToString();

            // Open file 
            if (result == "OK")
            {
                foldername = dlg.SelectedPath.ToString();
                exportLink.Text = foldername;
            }
        }

    }
}