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
        static int x = 200;
        static int y = 200;
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
            object misValue = System.Reflection.Missing.Value;

            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            if (xlRange.Cells[7, 3] != null && xlRange.Cells[7, 3].Value2 != null)
            {
                info.KenshuuDate = DateTime.FromOADate((double)xlRange.Cells[7, 3].Value2);
            }

            info.CourseNumber = getTextDataFromRange(xlRange, 8, 3);
            info.CourseName = getTextDataFromRange(xlRange, 9, 3);
            info.Department = getTextDataFromRange(xlRange, 10, 3);
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

            xlWorkbook.Close(false, misValue, misValue);

            return info;
        }

        public int setStatus(Boolean tmp)
        {
            if (tmp == true) return 1;
            else return 2;
        }
        
        private void writeExcelFile(String fileName, List<Training> dataList) 
        {
            xlApp.StandardFont = "游ゴシック";
            xlApp.StandardFontSize = 10;
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Add(Type.Missing);

            Excel._Worksheet xlWorkSheet = (Excel._Worksheet) xlWorkBook.ActiveSheet;  
            
            //Set workSheet Info
            xlWorkSheet.Name = "合計";
               
            xlWorkSheet.Cells[1, 1] = "No.";
            xlWorkSheet.Cells[1, 2] = "研修日";
            xlWorkSheet.Cells[1, 3] = "コースNo.";
            xlWorkSheet.Cells[1, 4] = "研修名";
            xlWorkSheet.Cells[1, 5] = "署名";
            xlWorkSheet.Cells[1, 6] = "社員番号";
            xlWorkSheet.Cells[1, 7] = "氏名";
            
            xlWorkSheet.Cells[1, 8] = "講師";
            xlWorkSheet.Cells[1, 9] = "テキスト";
            xlWorkSheet.Cells[1, 10] = "内容";
            xlWorkSheet.Cells[1, 11] = "継続有無";
            xlWorkSheet.Cells[1, 12] = "理由：講師";
            xlWorkSheet.Cells[1, 13] = "理由：テキスト";
            xlWorkSheet.Cells[1, 14] = "理由：内容";
            xlWorkSheet.Cells[1, 15] = "理由：継続有無";

            xlWorkSheet.Cells[1, 16] = "時期";
            xlWorkSheet.Cells[1, 17] = "日数";
            xlWorkSheet.Cells[1, 18] = "対象者条件";
            xlWorkSheet.Cells[1, 19] = "理由：時期";
            xlWorkSheet.Cells[1, 20] = "理由：日数";
            xlWorkSheet.Cells[1, 21] = "理由：対象者条件";

            //Write data into excel File
            int i = 2;
            foreach(Training dataInfo in dataList) 
            {
                xlWorkSheet.Cells[i, 1] = i-1;
                xlWorkSheet.Cells[i, 2] = dataInfo.KenshuuDate;
                xlWorkSheet.Cells[i, 3] = dataInfo.CourseNumber;
                xlWorkSheet.Cells[i, 4] = dataInfo.CourseName;
                xlWorkSheet.Cells[i, 5] = dataInfo.Department;
                xlWorkSheet.Cells[i, 6] = dataInfo.StaffCode;
                xlWorkSheet.Cells[i, 7] = dataInfo.StaffName;
                
                xlWorkSheet.Cells[i, 8] = setStatus(dataInfo.Lecturer);
                xlWorkSheet.Cells[i, 9] = setStatus(dataInfo.Text);
                xlWorkSheet.Cells[i, 10] = setStatus(dataInfo.Content);
                xlWorkSheet.Cells[i, 11] = setStatus(dataInfo.Continuation);
                xlWorkSheet.Cells[i, 12] = dataInfo.Lecturer_reason;
                xlWorkSheet.Cells[i, 13] = dataInfo.Text_reason;
                xlWorkSheet.Cells[i, 14] = dataInfo.Content_reason;
                xlWorkSheet.Cells[i, 15] = dataInfo.Continuation_reason;
                              
                xlWorkSheet.Cells[i, 16] = setStatus(dataInfo.Time);
                xlWorkSheet.Cells[i, 17] = setStatus(dataInfo.Day);
                xlWorkSheet.Cells[i, 18] = setStatus(dataInfo.Condition);
                xlWorkSheet.Cells[i, 19] = dataInfo.Time_reason;
                xlWorkSheet.Cells[i, 20] = dataInfo.Day_reason;
                xlWorkSheet.Cells[i, 21] = dataInfo.Condition_reason;

                ++i;
            }



            xlWorkBook.SaveAs(@"C:\Users\n3835\Desktop\キー\" + fileName);
            xlWorkBook.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            String fileName = DateTime.Now.ToString("dd_MM_yyyy_HHmmss") + ".xlsx";
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                DirectoryInfo d = new System.IO.DirectoryInfo(fbd.SelectedPath);
                FileInfo[] Files = d.GetFiles("*.xlsx");
                Training info = new Training();
                foreach (FileInfo file in Files)
                {
                    objectList.Add(readDataFromFile(file, fbd.SelectedPath));
                }
                writeExcelFile(fileName, objectList);
                MessageBox.Show("ファイル集合が完了しました");
                objectList.Clear();
                xlApp.Quit();
            }
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }
    }
}
