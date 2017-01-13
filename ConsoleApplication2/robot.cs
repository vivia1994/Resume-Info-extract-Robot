using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Data;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace robot
{
    class Program
    {
        static object missing = Type.Missing;
        public static List<string> fileContents = new List<string>();
        public static string fileContentString;
        public static string skills = "";
        public static System.Data.DataTable candidates = new System.Data.DataTable("candidates");
        public static int fileCount = 0;
        public static string excelPath = @"C:\result\Candidate Database1.xlsx";
        private Stopwatch wath = new Stopwatch();
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
            candidates.Columns.Add(" CurrentSalary", Type.GetType("System.String"));
            candidates.Columns.Add("ExpectedSalary", Type.GetType("System.String"));
            candidates.Columns.Add("PhoneNumber", Type.GetType("System.String"));
            candidates.Columns.Add("Email", Type.GetType("System.String"));
        }
        #endregion
        #region   //WriteToText
        public static void WriteToText(string s)
        {
            //test:表示向txt写入文本
            StreamWriter sw = new StreamWriter(@"C:\HR RPA\1.txt");
            sw.Write(s);
            sw.Close();
        }
        #endregion
        #region //ReadStringFromWord
        public static void ReadStringFromWord(string filePath)
        {
            //fileContentString = string.Empty;
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
            if (candidates==null)
            {
                InitializeDatatable(candidates);
            }
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
            fileCount++;
            candidate["ResumeRecommendDate"] = DateTime.Now;
            candidate["ChineseName"] = fileContents[0];
            candidate["ResourceName"] = candidate["ChineseName"];
            candidate["PhoneNumber"] = fileContents[4].Split(new string[1] { "︳" }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToList().ElementAtOrDefault(0);
            candidate["Email"] = fileContents[4].Split(new string[1] { "︳" }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToList().ElementAtOrDefault(1);
            foreach (string fileContent in fileContents)
            {
                if (fileContent.Contains("毕业"))
                {
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
            string stepId = "(Start)";
            foreach (string row in fileContents)
            {
                switch (row.Replace("　", string.Empty).Replace("：", string.Empty))
                {
                    case "工作经历":
                        stepId = "CurrentCompany";
                        break;
                    case "经历":
                        stepId = "University";
                        break; ;
                    case "目前年收入":
                        stepId = "CurrentSalary";
                        break;
                    case "期望工作":
                        stepId = "ExpectedSalary";
                        break;
                    case "教育经历":
                        stepId = "YearGraduated";
                        break;
                    case "个人信息":
                        stepId = "YearsExperience";
                        break;
                    default:
                        #region //default
                        if (row.StartsWith("工作经历"))
                        {
                            stepId = "CurrentCompany";
                        }
                        if (row.Contains("年工作经验"))
                        {
                            IEnumerable<string> relatedYears = row.Split(new string[1] { "︳" }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToList();
                            candidate["RelatedYears"] = relatedYears.ElementAtOrDefault(2).Split(new string[1] { "年工作经验" }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToList().OfType<string>().FirstOrDefault() + "年";
                        }
                        //Switch<string> - Assign field values
                        switch (stepId)
                        {
                            case "Gender|Age|Address|YearsExperience":
                                List<string> valueArray = row.Split(new string[1] { "    " }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).OfType<string>().ToList();
                                candidate["RelatedYears"] = valueArray.Where(s => s.Contains("工作经历")).Select(s => s.Replace("工作经历", string.Empty).Trim()).OfType<string>().FirstOrDefault();
                                break;
                            case "CurrentSalary":
                                candidate["CurrentSalary"] = row.Substring(0, new List<int>(new int[] {
    row.IndexOf("("),row.Length}).Where(i => i >= 0).FirstOrDefault()).Replace("目前年收入：", string.Empty).Trim()
        ;
                                break;
                            case "YearGraduated":
                                candidate["University"] = row;
                                break;
                            case "CurrentCompany":
                                candidate["CurrentCompany"] = fileContents.ElementAtOrDefault(2).Split(new string[1] { " · " }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).OfType<string>().ToList().ElementAtOrDefault(1);
                                break;
                            case "JobStatus":
                                candidate["JobStatus"] = row;
                                break;
                            case "ExpectedSalary":
                                candidate["ExpectedSalary"] = row.Split(new string[1] { "，" }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).OfType<string>().ToList().ElementAtOrDefault(3);
                                break;
                            default:
                                if(!string.IsNullOrWhiteSpace(stepId) & candidates.Columns.OfType<DataColumn>().Any(column => column.ColumnName == stepId))
                                    candidate[stepId] = row;
                                break;
                        }
                        #endregion
                        //TODO:assign the next step
                        switch (stepId)
                        {
                            default:
                                stepId = string.Empty;
                                break; 
                        }
                        break;
                }


            }
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
        #region   //WriteDataTabletoExcel
        //TODO:追加
        public static void WriteDataTabletoExcel(System.Data.DataTable candidates, string excelPath)
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
            Workbook wbook = excelApp.Workbooks.Open(excelPath);
            Worksheet worksheet = (Worksheet)wbook.Worksheets[1];
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
        #region   //AppendDataTabletoExcel
        //TODO:追加
        public static void AppendDataTabletoExcel(System.Data.DataTable candidates, string excelPath)
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
            Worksheet worksheet = (Worksheet)wbook.Worksheets[1];
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
        #region  //ReadExcelToDatatable
        public static System.Data.DataTable ReadExcelToDatatable(string excelPath)
        {
            //excel
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = null;
            Worksheet worksheet;
System.Data.DataTable datatable = new System.Data.DataTable("datatable");
            try
            {
                if(excelApp == null)
                    return null;
                workbook = excelApp.Workbooks.Open(excelPath);
                worksheet = (Worksheet)workbook.Worksheets[1];
                int rowExcelCount = worksheet.UsedRange.Rows.Count;
                int columnExcelCount = worksheet.UsedRange.Columns.Count;
                int rowExcelIndex = 1;
                int columnExcelIndex = 1;
                
                Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[rowExcelIndex, columnExcelIndex];
                //datatable name
                while (!string.IsNullOrEmpty(range.Text.ToString().Trim()))
                {
                    datatable.Columns.Add(range.Text.ToString(), Type.GetType("System.String"));
                    range = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[1, ++columnExcelIndex];
                }
                //datatable contents
                for (int i = 2; i <= rowExcelCount; i++)
                {
                    DataRow datarow = datatable.NewRow();
                    for (int j = 1; j <= columnExcelCount; j++)
                    {
                        range = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[i, j];
                        datarow[j - 1] = string.IsNullOrEmpty(range.Text.ToString().Trim()) ? "" : range.Text.ToString();
                    }
                    datatable.Rows.Add(datarow);
                }
                return datatable;
            }
            catch
            {
                return null;
            }
            finally
            {
                workbook.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                workbook = null;
                excelApp.Workbooks.Close();
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                excelApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
           
        }
        #endregion
        /* #region //ThreadReadExcel
        /// <summary>
        /// 使用COM，多线程读取Excel（1 主线程、4 副线程）
        /// </summary>
        /// <param name="excelFilePath">路径</param>
        /// <returns>DataTabel</returns>
        public System.Data.DataTable ThreadReadExcel(string excelFilePath)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Sheets sheets = null;
            Workbook workbook = null;
            object oMissiong = System.Reflection.Missing.Value;
            System.Data.DataTable datatable = new System.Data.DataTable();
            wath.Start();
            try
            {
                if (app == null)
                {
                    return null;
                }
                workbook = app.Workbooks.Open(excelFilePath, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong,
                    oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong);
                //将数据读入到DataTable中——Start   
                sheets = workbook.Worksheets;
                Worksheet worksheet = (Worksheet)sheets.get_Item(1);//读取第一张表
                if (worksheet == null)
                    return null;
                string cellContent;
                int rowExcelCount = worksheet.UsedRange.Rows.Count;
                int colCount = worksheet.UsedRange.Columns.Count;
                Microsoft.Office.Interop.Excel.Range range;
                //负责列头Start
                DataColumn dc;
                int ColumnID = 1;
                range = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[1, 1];
                while (colCount >= ColumnID)
                {
                    dc = new DataColumn();
                    dc.DataType = System.Type.GetType("System.String");
                    string strNewColumnName = range.Text.ToString().Trim();
                    if (strNewColumnName.Length == 0) strNewColumnName = "_1";
                    //判断列名是否重复
                    for (int i = 1; i < ColumnID; i++)
                    {
                        if (datatable.Columns[i - 1].ColumnName == strNewColumnName)
                            strNewColumnName = strNewColumnName + "_1";
                    }
                    dc.ColumnName = strNewColumnName;
                    datatable.Columns.Add(dc);
                    range = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[1, ++ColumnID];
                }
                //End
                //数据大于500条，使用多进程进行读取数据
                if (rowExcelCount - 1 > 500)
                {
                    //开始多线程读取数据
                    //新建线程
                    int b2 = (rowExcelCount - 1) / 10;
                    Microsoft.Office.Interop.Excel.DataTable dt1 = new Microsoft.Office.Interop.Excel.DataTable("dt1");
                    dt1 = datatable.Clone();
                    SheetOptions sheet1thread = new SheetOptions(worksheet, colCount, 2, b2 + 1, dt1);
                    Thread othread1 = new Thread(new ThreadStart(sheet1thread.SheetToDataTable));
                    othread1.Start();
                    //阻塞 1 毫秒，保证第一个读取 dt1
                    Thread.Sleep(1);
                    DataTable dt2 = new DataTable("dt2");
                    dt2 = datatable.Clone();
                    SheetOptions sheet2thread = new SheetOptions(worksheet, colCount, b2 + 2, b2 * 2 + 1, dt2);
                    Thread othread2 = new Thread(new ThreadStart(sheet2thread.SheetToDataTable));
                    othread2.Start();
                    DataTable dt3 = new DataTable("dt3");
                    dt3 = datatable.Clone();
                    SheetOptions sheet3thread = new SheetOptions(worksheet, colCount, b2 * 2 + 2, b2 * 3 + 1, dt3);
                    Thread othread3 = new Thread(new ThreadStart(sheet3thread.SheetToDataTable));
                    othread3.Start();
                    DataTable dt4 = new DataTable("dt4");
                    dt4 = datatable.Clone();
                    SheetOptions sheet4thread = new SheetOptions(worksheet, colCount, b2 * 3 + 2, b2 * 4 + 1, dt4);
                    Thread othread4 = new Thread(new ThreadStart(sheet4thread.SheetToDataTable));
                    othread4.Start();
                    //主线程读取剩余数据
                    for (int i = b2 * 4 + 2; i <= rowExcelCount; i++)
                    {
                        DataRow datarow = datatable.NewRow();
                        for (int j = 1; j <= colCount; j++)
                        {
                            range = (Excel.Range)worksheet.Cells[i, j];
                            cellContent = (range.Value2 == null) ? "" : range.Text.ToString();
                            datarow[j - 1] = cellContent;
                        }
                        datatable.Rows.Add(datarow);
                    }
                    othread1.Join();
                    othread2.Join();
                    othread3.Join();
                    othread4.Join();
                    //将多个线程读取出来的数据追加至 dt1 后面
                    foreach (DataRow datarow in datatable.Rows)
                        dt1.Rows.Add(datarow.ItemArray);
                    datatable.Clear();
                    datatable.Dispose();
                    foreach (DataRow datarow in dt2.Rows)
                        dt1.Rows.Add(datarow.ItemArray);
                    dt2.Clear();
                    dt2.Dispose();
                    foreach (DataRow datarow in dt3.Rows)
                        dt1.Rows.Add(datarow.ItemArray);
                    dt3.Clear();
                    dt3.Dispose();
                    foreach (DataRow datarow in dt4.Rows)
                        dt1.Rows.Add(datarow.ItemArray);
                    dt4.Clear();
                    dt4.Dispose();
                    return dt1;
                }
                else
                {
                    for (int i = 2; i <= rowExcelCount; i++)
                    {
                        DataRow datarow = datatable.NewRow();
                        for (int j = 1; j <= colCount; j++)
                        {
                            range = (Excel.Range)worksheet.Cells[i, j];
                            cellContent = (range.Value2 == null) ? "" : range.Text.ToString();
                            datarow[j - 1] = cellContent;
                        }
                        datatable.Rows.Add(datarow);
                    }
                }
                wath.Stop();
                TimeSpan ts = wath.Elapsed;
                //将数据读入到DataTable中——End
                return datatable;
            }
            catch
            {
                return null;
            }
            finally
            {
                workbook.Close(false, oMissiong, oMissiong);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheets);
                workbook = null;
                app.Workbooks.Close();
                app.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                app = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        #endregion*/
        static void Main(string[] args)
        {
            Console.WriteLine(string.Format("Pelease input wanted skills: "));
            skills = Console.ReadLine();
            string filePath = @"C:\HR RPA\Recruiting Team\candidatesResume\Zhao Zi Jun-赵子君的简历-Lagou.doc";
            string rootPath = @"C:\HR RPA\Recruiting Team\candidatesResume";

            //foreach (string file in Directory.GetFiles(rootPath, "*.doc*"))
            //{
            //    if (file.ToLower().Contains("lagou") )
            //    {
            //        ReadStringFromWord(filePath);
            //        if(Regex.Matches(fileContentString, "[\u4e00-\u9fa5]").Count > 20)
            //        {
            //            Console.WriteLine(file);
            //            AnalyzeFileContents(skills);
            //        }
            //    }
            //}
            //WriteDataTabletoExcel(candidates, excelPath);



            //string excelFilePath = @"C:\HR RPA\Recruiting Team\result\Candidate Database.xlsx";
            //ReadExcelToDatatable1(excelFilePath);
            //WriteDataTabletoExcel(datatable, excelPath);
            //Console.WriteLine(fileCount + " files has been updated!");
            ReadExcelToDatatable(excelPath);
            WriteDataTabletoExcel(ReadExcelToDatatable(excelPath), @"C:\result\Candidate Database.xlsx");
        }
    }
}