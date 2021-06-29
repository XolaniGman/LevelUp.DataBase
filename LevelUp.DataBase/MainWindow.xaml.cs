using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Azure.ActiveDirectory.GraphClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
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

namespace LevelUp.DataBase
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        DataSet ds = new DataSet();
        string connetionString = "";

        string ConnectionXML = "";

        string Connection = "";
        string DataBase = "";
        string databaseLogin = "";
        ExcelEngine excelEngine = new ExcelEngine();
        private object txtExcelToDatabase;

        public object ExcelVersion { get; private set; }

        public MainWindow()
        {
            InitializeComponent();
            FillDataGrid();
        }

        private void FillDataGrid()
        {


            System.Data.SqlClient.SqlConnection connection = new System.Data.SqlClient.SqlConnection(@"Server =.\LevelUpMedical; Database = LevelUpDrugs; Trusted_Connection = True;");

            connection.Open();

            //listBox2.Items.Clear();

            //string Tariff = myComboBox1.Text.Trim();

            string sql = "";

            sql = " SELECT ipkDrug ,sNappiCode ,sDescription from Drug    ";


            connetionString = @"Server =.\LevelUpMedical; Database = LevelUpDrugs; Trusted_Connection = True;";



            string CmdString = string.Empty;
            using (SqlConnection cnn = new SqlConnection(@"Server =.\LevelUpMedical; Database = LevelUpDrugs; Trusted_Connection = True;"))
            {
                cnn.Open();
                using (SqlDataAdapter sda = new SqlDataAdapter(" SELECT ipkDrug, sType, sNappiCode, sDescription, fPackPrice,fPackSize, fSchedule,fListPrice,fCostPrice,sStrength,sValid OldPrice from Drug   ", cnn))
                {
                    DataTable dt = new DataTable("Drug");
                    sda.Fill(dt);
                    DataGrid1.ItemsSource = dt.DefaultView;
                }
            }


        }


        private void btnSaveFile_Click(object sender, EventArgs e)
        {
            string fileName; 
            Spire.DataExport.XLS.CellExport cellExport = new Spire.DataExport.XLS.CellExport();
            Spire.DataExport.XLS.WorkSheet worksheet1 = new Spire.DataExport.XLS.WorkSheet();
            worksheet1.DataSource = Spire.DataExport.Common.ExportSource.DataTable;
            worksheet1.DataTable = this.DataGrid1.DataContext as DataTable;
            worksheet1.StartDataCol = ((System.Byte)(0));
            cellExport.Sheets.Add(worksheet1);
            cellExport.ActionAfterExport = Spire.DataExport.Common.ActionType.OpenView;
            string txtFileName = "";
            fileName = txtFileName.ToString() + "Boo1.xls";
            cellExport.SaveToFile(fileName);
        }


        //private void btnImport_Click(object sender, EventArgs e)
        //{
        //    string fileName;
        //    fileName = txtExcelToDatabase.ToString();
        //    Workbook workbook = new Workbook();
        //    workbook.LoadFromFile(fileName);
        //    Worksheet sheet = workbook.Worksheets[0];
        //    this.DataGrid1.DataSource = sheet.ExportDataTable();

        //}
        //privatevoid btnsaveTodatabase_Click(object sender, EventArgs e)
        //{
        //    SqlConnection Con = newSqlConnection("Data Source=Data-Source;Initial Catalog=Database-Name;Integrated Security=true");
        //    SqlCommand com;
        //    string str;
        //    Con.Open();
        //    for (int index = 0; index < dataGridView1.Rows.Count - 1; index++)
        //    {
        //        str = @ "Insert Into Employee(Emp_Id,Emp_Name,Manager_Id,Project_Id) Values(" + dataGridView1.Rows[index].Cells[0].Value.ToString() + ", '" + dataGridView1.Rows[index].Cells[1].Value.ToString() + "'," + dataGridView1.Rows[index].Cells[2].Value.ToString() + "," + dataGridView1.Rows[index].Cells[3].Value.ToString() + ")";
        //        com = newSqlCommand(str, Con);
        //        com.ExecuteNonQuery();
        //    }
        //    Con.Close();
        //}
        public class Drug
        {
            public int ipkDrug { get; set; }
            public string sType { get; set; }
            public string sNappiCode { get; set; }
            public string sDescription { get; set; }
            public double fPackPrice { get; set; }
            public string sStrength { get; set; }
            public double fPackSize { get; set; }
            public double fSchedule { get; set; }
            public double fListPrice { get; set; }
            public double fCostPrice { get; set; }
            public string sValid { get; set; }
            public double OldPrice { get; set; }

        }


    }
}
