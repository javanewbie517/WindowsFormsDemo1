using ExcelDataReader;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsDemo1.App;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsDemo1
{
    public partial class Form1 : Form
    {
        //设置数据库路径
        private string accessFilePath = AccessDAO.Property.accessFilePath;
        private string excelFilePath ;


        public Form1()
        {
            InitializeComponent();
        }

        private void btnConnect_Click(object sender, EventArgs e)
        {
           
            string sqlCommand = "select * from SysUser";
            DataSet dataSet = AccessDAO.getDataSetFromAccessTable(sqlCommand, accessFilePath);
            dataGridViewMain.DataSource = dataSet.Tables[0];
        }


        private void Form1_Load(object sender, EventArgs e)
        {

        }


        //打开文件
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "(*.xls;*.xlsx)|*.xls;*.xlsx";
            dialog.Title = "选择文件";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = dialog.SafeFileName;//获取文件名
                excelFilePath = dialog.FileName;//获取整个路径
                textBox2.Text = excelFilePath;
            }
           
        }
      

        public DataSet ReadExcelToDataSet(string fileNmaePath)
        {
            FileStream stream = null;
            IExcelDataReader excelReader = null;
            DataSet dataSet = null;
            try
            {
                //stream = File.Open(fileNmaePath, FileMode.Open, FileAccess.Read);
                stream = new FileStream(fileNmaePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            }
            catch
            {
                return null;
            }
            string extension = Path.GetExtension(fileNmaePath);

            if (extension.ToUpper() == ".XLS")
            {
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            else if (extension.ToUpper() == ".XLSX")
            {
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }
            else
            {
                MessageBox.Show("格式错误");
                return null;

            }
            //dataSet = excelReader.AsDataSet();//第一行当作数据读取
            dataSet = excelReader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });//第一行当作列名读取
            excelReader.Close();
            return dataSet;
        }
       
        //导入excel数据到access数据库
        private void button2_Click(object sender, EventArgs e)
        {
            //if (textBox1.Text.Length == 0)
            //{
            //    MessageBox.Show("请选择导入数据的Execl文件", "提示");
            //}

            DataSet dataSet = ReadExcelToDataSet(excelFilePath);
            dataGridViewMain.DataSource = dataSet.Tables[0];
            DataTable dt = dataSet.Tables[0];

            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    //写入数据库数据
                    string MySql = "insert into SysUser([username],[password]) values('" + dt.Rows[i]["username"].ToString() + "','" + dt.Rows[i]["password"].ToString() + "')";
                    //MessageBox.Show(MySql);
                    AccessDAO.updateAccessTable(MySql, accessFilePath);
            
                }
            
                MessageBox.Show("数据导入成功！");
            }
        }

        //将access数据导出为excel
        private void button3_Click(object sender, EventArgs e)
        {
            Excel.Application myexcelApplication = new Excel.Application();
            if (myexcelApplication != null)
            {
                Excel.Workbook myexcelWorkbook = myexcelApplication.Workbooks.Add();
                Excel.Worksheet myexcelWorksheet = (Excel.Worksheet)myexcelWorkbook.Sheets.Add();

                

                string sqlCommand = "select * from SysUser";
                DataSet dataSet = AccessDAO.getDataSetFromAccessTable(sqlCommand, accessFilePath);
               DataTable dt= dataSet.Tables[0];
                dataGridViewMain.DataSource =dt;

                //myexcelWorksheet.Cells[1, 1] = "ID";
                //myexcelWorksheet.Cells[1, 2] = "username";
                //myexcelWorksheet.Cells[1, 3] = "password";
               
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    myexcelWorksheet.Cells[ 1, i+1] = dt.Columns[i].ColumnName;
                  
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                 myexcelWorksheet.Cells[i+2, 1] = dt.Rows[i][0];
                myexcelWorksheet.Cells[i+2, 2] = dt.Rows[i][1];
                    myexcelWorksheet.Cells[i+2, 3] = dt.Rows[i][2];
                }
             
                myexcelApplication.ActiveWorkbook.SaveAs(@"F:\abc.xls", Excel.XlFileFormat.xlWorkbookNormal);
                myexcelWorkbook.Close();
                myexcelApplication.Quit();
                MessageBox.Show("导出成功");
            }
        }
    }
}
