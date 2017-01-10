using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;

namespace robot
{
    class Program
    {
        public static List<string> fileContents = new List<string>();
        public static string fileContentString;

        /// <summary>
        /// read info from word
        /// </summary>
        /// <param name="args"></param>
        public static void ReadStringFromWord(string filePath)
        {
            Microsoft.Office.Interop.Word.ApplicationClass word = new Microsoft.Office.Interop.Word.ApplicationClass();
            Document doc = new Document();
            object fileName = filePath;
            object missing = Type.Missing;
            doc = word.Documents.Open(ref fileName,
             ref missing,ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
             ref missing, ref missing, ref missing,ref missing,   ref missing, ref missing,  ref missing,ref missing);
            //fileContentString = doc.Content.Text;
            //Console.WriteLine(fileContentString);
            //fileContentString = fileContentString.Replace(char.ConvertFromUtf32(1), string.Empty).Replace(char.ConvertFromUtf32(7), string.Empty).Replace(char.ConvertFromUtf32(21), string.Empty);
            for (int i = 0; i < doc.Paragraphs.Count; i++)
            {
                string temp = doc.Paragraphs[i + 1].Range.Text.Trim();
                Console.WriteLine(temp);
                Console.WriteLine("++++++++");
                if (temp != string.Empty & temp != "\r\n" & temp != "\n" & temp != "\r")
                    fileContents.Add(temp);
                //}
                //foreach (string item in fileContents)
                //{
                //    Console.WriteLine(item);
                //
            }
            Console.WriteLine(doc.Paragraphs.Count);
            doc.Close();
            word.Quit();
        }
        /// <summary>
        /// analyze fileContentString
        /// </summary>
        /// <param name="fileContentString"></param>
        public static void AnalyzeString(string fileContentString)
        {
            fileContents = fileContentString.Split(new string[3] { "\r", "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Trim()).Where(s => !string.IsNullOrWhiteSpace(s) & s != "/").ToList();
            //foreach (string item in fileContents)
            //{
            //    Console.WriteLine(item);
            //}
        }
        static void Main(string[] args)
        {
            string filePath = @"C:\HR RPA\Recruiting Team\candidatesResume\En-51job_周瑞福(7502927).docx";
            ReadStringFromWord(filePath);
            //foreach (string item in fileContents)
            //{
            //    if(item == string.Empty | item == "\r\n"| item == "\n" | item == "\r")
            //    {
            //        item.Replace(item,string.Empty);
            //    }
            ////}
            //foreach (string item in fileContents)
            //{
            //    Console.WriteLine(item);
            //}
            Console.ReadKey();
        }
    }
}