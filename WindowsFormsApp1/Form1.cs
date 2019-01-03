using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        protected Excel.Application xlApp = new Excel.Application();
        private List<Training> objectList = new List<Training>();
        public Form1()
        {
            InitializeComponent();
        }

        private String getTextDataFromRange(Excel.Range xlRange, int row, int col)
        {
            String data = "";
            if (xlRange.Cells[row, col] != null && xlRange.Cells[row, col].Value2 != null)
            {
                data = xlRange.Cells[row, col].Value2.ToString();
            }
            return data;
        }

        private Boolean getBooleanData(Excel.Range xlRange, int row, int col)
        {
            Boolean data = true;
            if (xlRange.Cells[row, col] != null && xlRange.Cells[row, col].Value2 != null)
            {
                //if value = 2 -> data = false
                if ((double)xlRange.Cells[row, col].Value2 == 2)
                    data = false;
            }
            else
                data = false;
            return data;
        }

        private Training readDataFromFile( FileInfo file, String selectedPath)
        {
            Training info = new Training();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(selectedPath + "/" + @file.Name);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];

            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            if (xlRange.Cells[7, 3] != null && xlRange.Cells[7, 3].Value2 != null)
            {
                info.KenshuuDate = DateTime.FromOADate((double)xlRange.Cells[7, 3].Value2);
            }

            info.CourseNumber = getTextDataFromRange(xlRange, 8, 3);
            info.CourseName = getTextDataFromRange(xlRange, 9, 3);
            info.CourseNumber = getTextDataFromRange(xlRange, 10, 3);
            info.StaffCode = getTextDataFromRange(xlRange, 11, 3);
            info.StaffName = getTextDataFromRange(xlRange, 12, 3);

            info.Lecturer = getBooleanData(xlRange, 17, 5);
            info.Text = getBooleanData(xlRange, 18, 5);
            info.Content = getBooleanData(xlRange, 19, 5);
            info.Continuation = getBooleanData(xlRange, 20, 5);
            info.Time = getBooleanData(xlRange, 25, 5);
            info.Day = getBooleanData(xlRange, 26, 5);
            info.Condition = getBooleanData(xlRange, 27, 5);

            info.Lecturer_reason = getTextDataFromRange(xlRange, 17, 6);
            info.Text_reason = getTextDataFromRange(xlRange, 18, 6);
            info.Content_reason = getTextDataFromRange(xlRange, 19, 6);
            info.Continuation_reason = getTextDataFromRange(xlRange, 20, 6);
            info.Time_reason = getTextDataFromRange(xlRange, 25, 6);
            info.Day_reason = getTextDataFromRange(xlRange, 26, 6);
            info.Condition_reason = getTextDataFromRange(xlRange, 27, 6);

            xlWorkbook.Close();

            return info;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                DirectoryInfo d = new System.IO.DirectoryInfo(fbd.SelectedPath);
                FileInfo[] Files = d.GetFiles("*.xlsx");
                Training info = new Training();
                foreach (FileInfo file in Files)
                {
                    objectList.Add(readDataFromFile(file, fbd.SelectedPath));
                }
                MessageBox.Show("Read file successfully");
                objectList.Clear();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
