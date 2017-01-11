using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System.IO;

namespace robot
{
    class Program
    {
        static object missing = Type.Missing;
        public static List<string> fileContents = new List<string>();
        public static string fileContentString;
        #region //ReadStringFromWord
        public static void ReadStringFromWord(string filePath)
        {
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            Document doc = new Document();
            object fileName = filePath;
            object missing = Type.Missing;
            if (File.Exists(filePath))
            {
                doc = word.Documents.Open(ref fileName, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                for (int i = 0; i < doc.Paragraphs.Count; i++)
                {
                    string temp = doc.Paragraphs[i + 1].Range.Text.Trim();
                    if (temp != string.Empty & temp != "\r\n" & temp != "\n" & temp != "\r")
                    {
                        fileContents.Add(temp);
                    }
                }
                Console.WriteLine(doc.Paragraphs.Count);
                doc.Close();
                word.Quit();
            }
        }
        #endregion
        #region //AnalyzeString
        public static void AnalyzeString(string fileContentString)
        {
            foreach (string temp in fileContents)
            {
                fileContents.Remove(temp);
                temp.Replace(char.ConvertFromUtf32(1), string.Empty).Replace(char.ConvertFromUtf32(7), string.Empty)
                    .Replace(char.ConvertFromUtf32(21), string.Empty);
                //fileContents = temp.Split(new string[3] { "\r", "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries)
                //.Select(s => s.Trim()).Where(s => !string.IsNullOrWhiteSpace(s) & s != "/").ToList();
                fileContents.Add(temp);
            }

        }
        #endregion
        #region //WriteToExcel
        public static void WriteToExcel(string excelPath)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            //excelApp.Application.Workbooks.Add(true);
            excelApp.Visible = true;
            Workbook wbook = excelApp.Workbooks.Open(excelPath, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
            Worksheet worksheet = (Worksheet)wbook.Worksheets["UT template"];
            string temp = (string)((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[1, 2]).Text;
            worksheet.Cells[1, 2] = "列内容";
            wbook.Save();
            //worksheet.SaveAs(excelPath, missing, missing, missing, missing, missing, missing, missing, missing);
            wbook.Close();
            excelApp.Quit();
            /*7设置某个单元格里的格式
                Excel.Range rtemp=worksheet.get_Range("A1","A1");
                rtemp.Font.Name="宋体";
                rtemp.Font.FontStyle="加粗"；
                rtemp.Font.Size=5;*/
            /*4:如果是新建一个excel文件:
            Application app = new Application();
            Workbook wbook = app.Workbook.Add(Type.missing);
            Worksheet worksheet = (Worksheet)wbook.Worksheets[1];*/

        }

        #endregion
        static void Main(string[] args)
        {
            string filePath = @"C:\\En-51job_周瑞福(7502927).docx";
            //ReadStringFromWord(filePath);
            string excelPath = @"C:\HR RPA\UT template.xlsx";
            WriteToExcel(excelPath);
            Console.ReadKey();
        }
    }
}