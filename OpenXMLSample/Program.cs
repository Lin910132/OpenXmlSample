using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.Data;
using System.IO;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLSample
{
    public class TestModel
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public int Age { get; set; }
        public bool Gender { get; set; }
        public DateTime CreateDate { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {

            File.Copy(@"C:\Users\dli\Desktop\Working Files\OpenXML\copy.xlsx",
                      @"C:\Users\dli\Desktop\Working Files\OpenXML\TestCopy.xlsx");

            Report report = new Report();
            report.Open(@"C:\Users\dli\Desktop\Working Files\OpenXML\TestCopy.xlsx");
            //report.Create(@"C:\Users\dli\Desktop\Working Files\OpenXML\TestCopy.xlsx");

            List<TestModel> data = new List<TestModel>();

            string tableName = "patient123";
            string template = "";
            char[] digits = new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };
            if (Char.IsDigit(tableName[tableName.Length - 1]))
            {
                template = tableName.Trim(digits);
            }

            for (int i = 0; i < 250; i++)
            {
                TestModel row = new TestModel
                {
                    FirstName = "Test",
                    LastName = "Name",
                    Age = 50,
                    Gender = true,
                    CreateDate = DateTime.Now
                };

                data.Add(row);
            }

            DataTable tmpTable = new DataTable();
            List<string> cols = new List<string>();
            cols.Add("Errors");
            cols.Add("Count");
            cols.Add("22");
            cols.Add("33");
            foreach (var col in cols)
            {
                tmpTable.Columns.Add(col);
            }

            int index = 4;
            for (int i = 0; i < 10; i++)
            {
                var newRow = tmpTable.Rows.Add();
                foreach (DataColumn col in tmpTable.Columns)
                {
                    newRow[col.ColumnName] = "test";
                }
                //report.InsertRowAt(newRow, tmpTable.Columns, index, qrdaSummary);
                index++;
            }

            index = index + 1;
            foreach (DataRow row in tmpTable.Rows)
            {
                //report.InsertRowAt(row, tmpTable.Columns, index, qrdaSummary);
                index++;
            }
            
            Dictionary<string, string> tokens = new Dictionary<string, string>();
            //tokens.Add("XXXX", "Replaced Item");
            tokens.Add("Program Type", "Custom Type");
            tokens.Add("@@Clientid@@", "NextGen");
            tokens.Add("@@CCN@@", "122113");
            tokens.Add("@@DateRecieved@@", DateTime.Now.ToShortDateString());
            tokens.Add("[NumberOfFiles]", "230");
            tokens.Add("[SuccessfullyLoaded]", "210");
            tokens.Add("@@ReportingYear@@", "2016");

            //var sheet = report.GetOrCreateWorkSheetByName("Test");
            //report.StartWritingWithoutTemplate(sheet);
            //report.AddDataWithoutTemplate(tmpTable);
            //report.EndWritingWithoutTemplate();

            var sheet2 = report.GetOrCreateWorkSheetByName("Test2");
            //report.StartWritingWithoutTemplate(sheet2);
            //report.EndWritingWithoutTemplate();

            var patient = report.GetOrCreateWorkSheetByName("Patient");
            var copy = report.CopySheet("Patient", "Patient1");
            var copy2 = report.CopySheet("Patient", "Patient2");
            var copy3 = report.CopySheet("Patient", "Patient3");
            var copy4 = report.CopySheet("Patient", "Patient4");
            
            //var copy = report.GetOrCreateWorkSheetByName("Patient1");
            //report.StartWritingWithoutTemplate(copy);
            //report.EndWritingWithoutTemplate();
            //var copy2 = report.GetOrCreateWorkSheetByName("Patient2");
            //var copy3 = report.GetOrCreateWorkSheetByName("Patient3");
            //var copy4 = report.GetOrCreateWorkSheetByName("Patient4");
            
            //var test = report.GetOrCreateWorkSheetByName("hello");
            report.AddDataWithTempate(tmpTable, patient);
            report.StartWritingWithoutTemplate(sheet2);
            report.AddDataWithoutTemplate(tmpTable);
            report.EndWritingWithoutTemplate();
            report.AddDataWithTempate(tmpTable, copy);
            report.AddDataWithTempate(tmpTable, copy2);
            report.AddDataWithTempate(tmpTable, copy3);
            report.AddDataWithTempate(tmpTable, copy4);
            //var x = report.GetOrCreateWorkSheetByName("summary");
            //report.ReplaceTemplateValuesByTokenDom(tokens, x);
            //report.StartWritingWithoutTemplate(newSheet);
            //report.AddDataWithoutTemplate(tmpTable);
            //report.AddDataWithoutTemplate(tmpTable);
            //report.EndWritingWithoutTemplate();

            report.Close();

        }
    }
}
