using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Data;

namespace robot
{
    class Program
    {
        static object missing = Type.Missing;
        public static List<string> fileContents = new List<string>();
        public static string fileContentString;
        public static string skills = "";
        #region //InitializeDatatable
        public static void InitializeDatatable(System.Data.DataTable candidates)
        {
            candidates.Columns.Add("ResumeRecommendDate", Type.GetType("System.String"));
            candidates.Columns.Add("InterviewDate", Type.GetType("System.String"));
            candidates.Columns.Add("ChineseName", Type.GetType("System.String"));
            candidates.Columns.Add("EnglishName", Type.GetType("System.String"));
            candidates.Columns.Add("JobStatus", Type.GetType("System.String"));
            candidates.Columns.Add("Channel", Type.GetType("System.String"));
            candidates.Columns.Add("ResourceName", Type.GetType("System.String"));
            candidates.Columns.Add("Skill", Type.GetType("System.String"));
            candidates.Columns.Add("RelatedYears", Type.GetType("System.String"));
            candidates.Columns.Add("YearGraduated", Type.GetType("System.String"));
            candidates.Columns.Add("CurrentCompany", Type.GetType("System.String"));
            candidates.Columns.Add("University", Type.GetType("System.String"));
            candidates.Columns.Add("ExpectedSalary", Type.GetType("System.String"));
            candidates.Columns.Add("PhoneNumber", Type.GetType("System.String"));
            candidates.Columns.Add("Email", Type.GetType("System.String"));
        }
        #endregion
        #region //ReadStringFromWord
        public static void ReadStringFromWord(string filePath)
        {
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            Document doc = new Document();
            object fileName = filePath;
            object missing = Type.Missing;
            if (File.Exists(filePath))
            {
                try
                {
                    word.Visible = true;
                    doc = word.Documents.Open(ref fileName, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                    fileContentString = doc.Content.Text;
                    /*for (int i = 0; i < doc.Paragraphs.Count; i++)
                    {
                        string temp = doc.Paragraphs[i + 1].Range.Text.Trim();
                        if (temp != string.Empty & temp != "\r\a" & temp != "\a" & temp != "\r")
                        {
                            fileContents.Add(temp);
                        }
                    }
                    foreach (var item in fileContents)
                    {
                        Console.WriteLine(item); Console.WriteLine("___");
                    }*/
                    //test:表示向txt写入文本
                    StreamWriter sw = new StreamWriter(@"C:\HR RPA\1.txt");
                    sw.Write(fileContentString);
                    sw.Close();
                    //Console.WriteLine(doc.Paragraphs.Count);
                    doc.Close();
                    word.Quit();
                }
                catch (Exception ex)
                {
                    doc.Close();
                    word.Quit();
                }
            }
        }
        #endregion
        #region //AnalyzeFileContents
        public static void AnalyzeFileContents(string skills)
        {
            System.Data.DataTable candidates = new System.Data.DataTable("candidates");
            InitializeDatatable(candidates);
            fileContentString = fileContentString.Replace(char.ConvertFromUtf32(1), string.Empty).Replace(char.ConvertFromUtf32(7), string.Empty).Replace(char.ConvertFromUtf32(21), string.Empty);
            /* foreach (string temp in fileContents)
             {
                 fileContents.Remove(temp);
                 temp.Replace(char.ConvertFromUtf32(1), string.Empty).Replace(char.ConvertFromUtf32(7), string.Empty)
                     .Replace(char.ConvertFromUtf32(21), string.Empty);
                 //fileContents = temp.Split(new string[3] { "\r", "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries)
                 //.Select(s => s.Trim()).Where(s => !string.IsNullOrWhiteSpace(s) & s != "/").ToList();
                 fileContents.Add(temp);
             }*/
            fileContents = fileContentString.Split(new string[3] { "\r", "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).Where(s => !string.IsNullOrWhiteSpace(s) & s != "/").ToList();
            DataRow candidate = candidates.Rows.Add();
            candidate["ResumeRecommendDate"] = DateTime.Now;
            candidate["ChineseName"] = fileContents[0];
            candidate["ResourceName"] = candidate["ChineseName"];
            candidate["PhoneNumber"] = fileContents[4].Split(new string[1] { "︳" }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToList().ElementAtOrDefault(0);
            candidate["Email"] = fileContents[4].Split(new string[1] { "︳" }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToList().ElementAtOrDefault(1);
            foreach(string fileContent in fileContents)
            {
                if (fileContent.Contains("毕业")){
                    int yearGraduate = int.Parse(fileContent.Split(new string[1] { "年毕业" }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToList().ElementAtOrDefault(0));
                    candidate["YearGraduated"] = int.Parse(DateTime.Now.Year.ToString()) - yearGraduate + "年";
                                   }
            }
            if (!string.IsNullOrEmpty(skills))
            {
                if (skills.Contains("|"))
                {
                    candidate["Skill"] = string.Join("，", skills.Split(new string[1] { "|" }, StringSplitOptions.RemoveEmptyEntries).Where(s => fileContentString.IndexOf(s, StringComparison.InvariantCultureIgnoreCase) >= 0));
                }
                else
                {
                    candidate["Skill"] = string.Join(", ", skills.Split(new string[1] { "" }, StringSplitOptions.RemoveEmptyEntries).Where(s => fileContentString.IndexOf(s, StringComparison.InvariantCultureIgnoreCase) >= 0));
                }
            }
            else
            {
                candidate["Skill"] = "";
            }
            candidate["Channel"] = "LaGou";

        }
        #endregion
        #region //WriteToExcel
        public static void WriteToExcel(string excelPath)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;
            if (File.Exists(excelPath))
            {
                Workbook wbook = excelApp.Workbooks.Open(excelPath, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
                Worksheet worksheet = (Worksheet)wbook.Worksheets["UT template"];
                string temp = (string)((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[1, 2]).Text;
                worksheet.Cells[1, 2] = "列内容";
                wbook.Save();
                //worksheet.SaveAs(excelPath, missing, missing, missing, missing, missing, missing, missing, missing);
                wbook.Close();
                excelApp.Quit();
            }
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
        #region   //DataTabletoExcel
        public static void DataTabletoExcel(System.Data.DataTable candidates, string excelPath)
        {
            if (candidates == null)
            {
                return;
            }
            int rowNum = candidates.Rows.Count;
            int columnNum = candidates.Columns.Count;
            int rowIndex = 1;
            int columnIndex = 0;
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.DefaultFilePath = @"C:\HR RPA\UT template.xlsx";
            excelApp.DisplayAlerts = true;
            Workbook wbook = excelApp.Workbooks.Open(excelPath, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
            Worksheet worksheet = (Worksheet)wbook.Worksheets["UT template"];
            //excelApp.SheetsInNewWorkbook = 1;
            //Workbook wbook = excelApp.Workbooks.Add(true);

            //将DataTable的列名导入Excel表第一行
            foreach (DataColumn dataColumn in candidates.Columns)
            {
                columnIndex++;
                excelApp.Cells[rowIndex, columnIndex] = dataColumn.ColumnName;
            }
            //将DataTable中的数据导入Excel中
            for (int i = 0; i < rowNum; i++)
            {
                rowIndex++;
                columnIndex = 0;
                for (int j = 0; j < columnNum; j++)
                {
                    columnIndex++;
                    excelApp.Cells[rowIndex, columnIndex] = candidates.Rows[i][j].ToString();
                }
            }
            wbook.Save();
        }
        #endregion
        static void Main(string[] args)
        {
            Console.WriteLine(string.Format("Pelease input wanted skills: "));
            
            string filePath = @"C:\HR RPA\Recruiting Team\candidatesResume\Zhao Zi Jun-赵子君的简历-Lagou.doc";
            
            string excelPath = @"C:\HR RPA\UT template.xlsx";
            //WriteToExcel(excelPath);
            string rootPath = @"C:\HR RPA\Recruiting Team\candidatesResume";
            foreach (string file in Directory.GetFiles(rootPath, "*.doc*"))
            {
                if (file.ToLower().Contains("lagou"))
                {
                    //Console.WriteLine(file);
                    //ReadStringFromWord(file);
                    //AnalyzeFileContents(skills);
                    ///DatatableToExcel(excelPath);
                }
            }
            ReadStringFromWord(filePath);
            AnalyzeFileContents(skills);

            //skills = Console.ReadLine();
        }
    }
}