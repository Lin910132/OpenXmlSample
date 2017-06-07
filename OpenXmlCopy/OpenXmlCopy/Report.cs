using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;

namespace OpenXmlCopy
{
    public class Report
    {
        private static void CleanView(WorksheetPart worksheetPart)
        {

            //There can only be one sheet that has focus

            SheetViews views = worksheetPart.Worksheet.GetFirstChild<SheetViews>();

            if (views != null)
            {
                views.Remove();
                worksheetPart.Worksheet.Save();
            }
        }

        private static void FixupTableParts(WorksheetPart worksheetPart, int tableId)
        {

            //Every table needs a unique id and name

            foreach (TableDefinitionPart tableDefPart in worksheetPart.TableDefinitionParts)
            {
                tableId++;
                tableDefPart.Table.Id = (uint)tableId;
                tableDefPart.Table.DisplayName = "CopiedTable" + tableId;
                tableDefPart.Table.Name = "CopiedTable" + tableId;
                tableDefPart.Table.Save();
            }

        }

        private static SharedStringTablePart GetSharedStringPart(WorkbookPart workbookPart)
        {
            SharedStringTablePart shareStringPart;

            if (workbookPart.GetPartsCountOfType<SharedStringTablePart>() > 0)
                shareStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
            else
                return null;

            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            return shareStringPart;
        }

        private static void CopySharedStringPart(WorkbookPart sourceWorkbookPart, WorkbookPart clonedWorkbookPart)
        {
            SharedStringTablePart sourceSharedStringPart = GetSharedStringPart(sourceWorkbookPart);

            if (sourceSharedStringPart == null)
                return;

            SharedStringTablePart clonedSharedStringPart = clonedWorkbookPart.AddNewPart<SharedStringTablePart>();
            var items = sourceSharedStringPart.SharedStringTable.ChildElements;
            clonedSharedStringPart.SharedStringTable = new SharedStringTable();

            foreach (var item in items)
            {
                clonedSharedStringPart.SharedStringTable.Append(item.CloneNode(true));
            }

            clonedSharedStringPart.SharedStringTable.Save();
        }

        private static void CopyStyles(WorkbookPart sourceWorkbookPart, WorkbookPart clonedWorkbookPart)
        {
            if (sourceWorkbookPart.WorkbookStylesPart == null || sourceWorkbookPart.WorkbookStylesPart.Stylesheet == null)
                return;

            var sourceStyle = sourceWorkbookPart.WorkbookStylesPart.Stylesheet.ChildElements;

            if (clonedWorkbookPart.WorkbookStylesPart == null)
                clonedWorkbookPart.AddNewPart<WorkbookStylesPart>();

            if (clonedWorkbookPart.WorkbookStylesPart.Stylesheet == null)
                clonedWorkbookPart.WorkbookStylesPart.Stylesheet = new Stylesheet();
            var clonedStyle = clonedWorkbookPart.WorkbookStylesPart.Stylesheet;

            foreach (var item in sourceStyle)
            {
                clonedStyle.Append(item.CloneNode(true));
            }

            clonedStyle.Save();
        }

        public static void Copy(string from, string to)
        {
            SpreadsheetDocument source = SpreadsheetDocument.Open(from, true);
            var sourceWorkbookPart = source.WorkbookPart;
            var sourceSheets = source.WorkbookPart.Workbook.Sheets.Descendants<Sheet>().ToList();

            var copy = SpreadsheetDocument.Create(to, source.DocumentType);
            var clonedWorkbookPart = copy.AddWorkbookPart();
            clonedWorkbookPart.Workbook = new Workbook();
            clonedWorkbookPart.Workbook.AppendChild(new Sheets());

            foreach (Sheet sourceSheet in sourceSheets)
            {
                WorksheetPart sourceSheetPart = (WorksheetPart)sourceWorkbookPart.GetPartById(sourceSheet.Id);
                var clonedSheet = clonedWorkbookPart.AddPart<WorksheetPart>(sourceSheetPart);
                
                int numTableDefParts = clonedSheet.GetPartsCountOfType<TableDefinitionPart>();

                int tableId = numTableDefParts;

                if (numTableDefParts != 0)
                    FixupTableParts(clonedSheet, numTableDefParts);

                CleanView(clonedSheet);

                Sheets sheets = clonedWorkbookPart.Workbook.GetFirstChild<Sheets>();
                Sheet copiedSheet = new Sheet();

                copiedSheet.Name = sourceSheet.Name;
                copiedSheet.Id = clonedWorkbookPart.GetIdOfPart(clonedSheet);
                copiedSheet.SheetId = (uint)sheets.ChildElements.Count + 1;

                sheets.Append(copiedSheet);

                //Save Changes

                sourceWorkbookPart.Workbook.Save();
            }

            
            CopySharedStringPart(sourceWorkbookPart, clonedWorkbookPart);
            CopyStyles(sourceWorkbookPart, clonedWorkbookPart);

            source.Close();
            copy.Close();
        }
    }
}
