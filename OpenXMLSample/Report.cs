using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Threading.Tasks;
using System.Data;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLSample
{
    public class Report
    {
        private SpreadsheetDocument _document;
        private WorkbookPart _workbookPart;
        private OpenXmlWriter _writer;
        private OpenXmlReader _reader;
        private string originalPartId;
        private string replacementPartId;

        public void Create(string fileName)
        {
            _document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);
            _workbookPart = _document.AddWorkbookPart();
            _workbookPart.Workbook = new Workbook();
            _workbookPart.Workbook.ChildElements.AsEnumerable();
            _workbookPart.Workbook.AppendChild(new Sheets());
        }

        public void Open(string fileName)
        {
            try
            {
                _document = SpreadsheetDocument.Open(fileName, true);
                _workbookPart = _document.WorkbookPart;

            }
            catch
            {
                Create(fileName);
            }
        }

        public void Close()
        {
            //WorkbookView view = _workbookPart.Workbook.Descendants<WorkbookView>().First();
            //var sheets = _workbookPart.Workbook.Sheets.Descendants<Sheet>().AsEnumerable();
            //var sheet = sheets.FirstOrDefault(x => x.Name == "Summary");

            //view.ActiveTab = sheet.SheetId;
            //_workbookPart.Workbook

            _document.Close();
            _document = null;
        }

        public void StartWritingWithoutTemplate(WorksheetPart worksheetPart)
        {
            _writer = OpenXmlPartWriter.Create(worksheetPart);

            _writer.WriteStartElement(new Worksheet());
            _writer.WriteStartElement(new SheetData());
        }

        public void EndWritingWithoutTemplate()
        {
            _writer.WriteEndElement();
            _writer.WriteEndElement();

            _writer.Close();
        }

        private void StartWritingWithTemplate(WorksheetPart originalPart)
        {
            WorksheetPart replacementPart = _workbookPart.AddNewPart<WorksheetPart>();

            _writer = OpenXmlWriter.Create(replacementPart);

            if (originalPart.Worksheet == null)
            {
                originalPart.Worksheet = new Worksheet(new SheetData());
            }

            _reader = OpenXmlReader.Create(originalPart);

            originalPartId = _workbookPart.GetIdOfPart(originalPart);
            replacementPartId = _workbookPart.GetIdOfPart(replacementPart);
        }

        private void HandleMergedCells()
        {
            Worksheet original = ((WorksheetPart)_workbookPart.GetPartById(originalPartId)).Worksheet;
            Worksheet replacement = ((WorksheetPart)_workbookPart.GetPartById(replacementPartId)).Worksheet;

            MergeCells originalMergeCells;

            if (original.Elements<MergeCells>().Count() > 0)
                originalMergeCells = original.Elements<MergeCells>().First();
            else
                return;

            MergeCells replacementMergeCells;
            if (replacement.Elements<MergeCells>().Count() > 0)
                replacementMergeCells = replacement.Elements<MergeCells>().First();
            else
                replacementMergeCells = new MergeCells();

            foreach (MergeCell cellToMerge in originalMergeCells.Elements<MergeCell>())
            {
                MergeCell mergedCell = (MergeCell)cellToMerge.Clone();
                replacementMergeCells.Append(mergedCell);
            }

            replacement.Append(replacementMergeCells);
            replacement.Save();
        }

        private void EndWritingWithTemplate()
        {
            _reader.Close();
            _writer.Close();

            //HandleMergedCells();

            Sheet sheet = _workbookPart.Workbook.Sheets.Cast<Sheet>().Where(x => x.Id.Value.Equals(originalPartId)).FirstOrDefault();
            if (sheet != null)
                sheet.Id.Value = replacementPartId;

            WorksheetPart originalPart = (WorksheetPart)_workbookPart.GetPartById(originalPartId);
            _workbookPart.DeletePart(originalPart);
        }

        private Sheet GetRefSheet(string name, Sheets sheets)
        {
            var tmplate = sheets.Elements<Sheet>().FirstOrDefault(x => name.Contains(x.Name));

            if (tmplate != null)
                return sheets.Elements<Sheet>().LastOrDefault(x => x.Name.ToString().Contains(tmplate.Name));
            return
                null;
        }

        private string CreateSheet(string name, Sheets sheets)
        {
            WorksheetPart worksheetPart = _workbookPart.AddNewPart<WorksheetPart>();
            //var sheets = _workbookPart.Workbook.Sheets;
            uint index = (uint)sheets.Elements<Sheet>().Count() + 1;
            var id = _workbookPart.GetIdOfPart(worksheetPart);

            Sheet sheet = new Sheet
            {
                Id = id,
                Name = name,
                SheetId = index
            };

            var refSheet = GetRefSheet(name, sheets);


            if (refSheet == null)
                sheets.Append(sheet);
            else
                sheets.InsertAfter(sheet, refSheet);
            //refSheet.InsertAfterSelf(sheet);

            return sheet.Id;
        }

        public WorksheetPart GetOrCreateWorkSheetByName(string name)
        {
            //IEnumerable<Sheet> sheets = _workbookPart.Workbook.Sheets.Elements<Sheet>();
            Sheets sheets = _workbookPart.Workbook.Sheets;
            Sheet sheet = sheets.Elements<Sheet>().FirstOrDefault(x => x.Name.ToString().ToLower() == name.ToLower());
            string refId = "";

            if (sheet == null)
                refId = CreateSheet(name, sheets);
            else
                refId = sheet.Id;

            WorksheetPart worksheetPart = (WorksheetPart)(_workbookPart.GetPartById(refId));

            //WorksheetPart worksheetPart = _workbookPart.AddNewPart<WorksheetPart>();

            return worksheetPart;
        }

        private List<string> GetSharedStringList()
        {
            SharedStringTablePart shareStringPart;
            if (_workbookPart.GetPartsCountOfType<SharedStringTablePart>() > 0)
                shareStringPart = _workbookPart.GetPartsOfType<SharedStringTablePart>().First();
            else
                shareStringPart = _workbookPart.AddNewPart<SharedStringTablePart>();

            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            var items = shareStringPart.SharedStringTable.Elements<SharedStringItem>();

            return items.Select(x => x.InnerText).ToList();

        }

        private Border GenerateBorder()
        {
            Border border = new Border();

            LeftBorder leftBorder2 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color1 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder2.Append(color1);

            RightBorder rightBorder2 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color2 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder2.Append(color2);

            TopBorder topBorder2 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color3 = new Color() { Indexed = (UInt32Value)64U };

            topBorder2.Append(color3);

            BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color4 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder2.Append(color4);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border.Append(leftBorder2);
            border.Append(rightBorder2);
            border.Append(topBorder2);
            border.Append(bottomBorder2);
            border.Append(diagonalBorder2);

            return border;
        }

        private uint InsertBorder(Border border)
        {
            if (_workbookPart.WorkbookStylesPart == null)
                _workbookPart.AddNewPart<WorkbookStylesPart>();

            var stylePart = _workbookPart.WorkbookStylesPart;
            if (stylePart.Stylesheet == null)
                stylePart.Stylesheet = new Stylesheet();

            Borders borders = _workbookPart.WorkbookStylesPart.Stylesheet.Elements<Borders>().FirstOrDefault();
            if (borders == null)
            {
                borders = new Borders();
                borders.Count = 0;
            }
            
            borders.Append(border);
            
            return (uint)borders.Count++;
        }

        private uint InsertCellFormat(WorkbookPart workbookPart, CellFormat cellFormat)
        {

            CellFormats cellFormats = workbookPart.WorkbookStylesPart.Stylesheet.Elements<CellFormats>().FirstOrDefault();
            if (cellFormats == null)
            {
                cellFormats = new CellFormats();
                cellFormats.Count = 0;
            }
                

            cellFormats.Append(cellFormat);
            return (uint)cellFormats.Count++;
        }


        private void InsertData(DataTable data)
        {

            foreach (DataRow dataRow in data.Rows)
            {
                Row r = new Row();
                _writer.WriteStartElement(r);

                foreach (DataColumn col in data.Columns)
                {
                    Cell c = new Cell
                    {
                        DataType = CellValues.String,
                        CellValue = new CellValue(dataRow[col].ToString())
                    };


                    CellFormat format = new CellFormat();
                    format.BorderId = InsertBorder(GenerateBorder());
                    c.StyleIndex = InsertCellFormat(_workbookPart, format);

                    _writer.WriteElement(c);
                }
                _writer.WriteEndElement();
            }
            //_workbookPart.Workbook.Save();
        }

        public void AddDataWithoutTemplate(DataTable data)
        {
            InsertData(data);
        }

        public void RemoveSharedStringPart()
        {
            SharedStringTablePart shareStringPart;
            if (_workbookPart.GetPartsCountOfType<SharedStringTablePart>() > 0)
                shareStringPart = _workbookPart.GetPartsOfType<SharedStringTablePart>().First();
            else
                shareStringPart = _workbookPart.AddNewPart<SharedStringTablePart>();

            _workbookPart.DeletePart(shareStringPart);
            _workbookPart.Workbook.Save();

        }

        public void InsertRowAt(DataRow rowToInsert, DataColumnCollection columns, int refRow, WorksheetPart worksheetPart)
        {
            CellFormat format = new CellFormat();
            format.BorderId = InsertBorder(GenerateBorder());
            uint styleIndex = InsertCellFormat(_workbookPart, format);

            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            Row newRow = new Row { RowIndex = Convert.ToUInt32(refRow) };

            foreach (DataColumn col in columns)
            {
                Cell cell = new Cell
                {
                    DataType = CellValues.String,
                    CellValue = new CellValue(" \n" + rowToInsert[col].ToString() + " s")
                };

                cell.StyleIndex = styleIndex;
                newRow.AppendChild(cell);
            }

            Row refrenceRow = null;//sheetData.Descendants<Row>().Where(x => x.RowIndex != null && x.RowIndex.Value == refRow).First();
            IEnumerable<Row> rows = sheetData.Descendants<Row>().Where(x => x.RowIndex != null && x.RowIndex >= refRow);
            foreach (Row row in rows)
            {
                if (row.Elements<Cell>().Any(x => x.CellReference != null))
                {
                    if (refrenceRow == null)
                        refrenceRow = row;

                    uint newIndex = row.RowIndex.Value + 1;

                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        string cellReference = cell.CellReference.Value;
                        cell.CellReference = cellReference.Replace(row.RowIndex.Value.ToString(), newIndex.ToString());
                    }

                    row.RowIndex = new UInt32Value(newIndex);
                }
            }

            if (refrenceRow != null)
                sheetData.InsertBefore(newRow, refrenceRow);
            else
                sheetData.Append(newRow);
            worksheetPart.Worksheet.Save();
        }

        public void ReplaceTemplateValuesByTokenDom(Dictionary<string, string> tokens, WorksheetPart worksheetPart)
        {

            List<string> sharedStrings = GetSharedStringList();
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            IEnumerable<Row> rows = sheetData.Descendants<Row>().Where(x => x.RowIndex != null);
            foreach (Row row in rows)
            {
                if (row.Elements<Cell>().Any(x => x.CellValue != null))
                {
                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        if (cell.CellValue == null)
                            continue;

                        string text;

                        try
                        {
                            int index = Convert.ToInt32(cell.InnerText);
                            text = sharedStrings[index];
                        }
                        catch
                        {
                            text = cell.InnerText;
                        }

                        tokens.Keys.Where(x => text.IndexOf(x) != -1).ToList().ForEach(x => text = text.Replace(x, tokens[x]));

                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(text);
                    }
                }
            }

            worksheetPart.Worksheet.Save();
        }

        public void ReplaceTemplateValuesByToken(Dictionary<string, string> tokens, WorksheetPart workSheetPart)
        {
            List<string> sharedStrings = GetSharedStringList();

            StartWritingWithTemplate(workSheetPart);

            while (_reader.Read())
            {
                if (_reader.ElementType == typeof(SheetData))
                {
                    if (_reader.IsEndElement)
                        continue;

                    //Reading template sheet data
                    _writer.WriteStartElement(new SheetData());
                    _reader.Read();

                    while (true)
                    {
                        if (_reader.ElementType == typeof(SheetData) && _reader.IsEndElement)
                            break;

                        if (_reader.IsStartElement)
                        {
                            _writer.WriteStartElement(_reader);


                            if (_reader.ElementType.IsSubclassOf(typeof(OpenXmlLeafTextElement)))
                            {
                                string text;

                                try
                                {
                                    int index = Convert.ToInt32(_reader.GetText());
                                    text = sharedStrings[index];
                                }
                                catch
                                {
                                    text = _reader.GetText();
                                }


                                tokens.Keys.Where(x => text.IndexOf(x) != -1).ToList().ForEach(x => text = text.Replace(x, tokens[x]));

                                _writer.WriteString(text);
                            }
                        }
                        else if (_reader.IsEndElement)
                        {
                            _writer.WriteEndElement();
                        }
                        _reader.Read();
                    }

                    _writer.WriteEndElement();//close sheet
                    //_workbookPart.Workbook.Save();
                    //break;
                }
                else
                {
                    if (_reader.IsStartElement)
                    {
                        _writer.WriteStartElement(_reader);
                    }
                    else if (_reader.IsEndElement)
                    {
                        _writer.WriteEndElement();
                    }
                }
            }

            EndWritingWithTemplate();
        }

        public void AddDataWithTempate(DataTable data, WorksheetPart workSheetPart)
        {
            StartWritingWithTemplate(workSheetPart);

            while (_reader.Read())
            {
                if (_reader.ElementType == typeof(SheetData))
                {
                    if (_reader.IsEndElement)
                        continue;

                    //Reading template sheet data
                    _writer.WriteStartElement(new SheetData());
                    _reader.Read();

                    while (true)
                    {
                        if (_reader.ElementType == typeof(Row) && _reader.IsEndElement)
                            break;

                        if (_reader.IsStartElement)
                        {
                            _writer.WriteStartElement(_reader);
                            if (_reader.ElementType.IsSubclassOf(typeof(OpenXmlLeafTextElement)))
                            {
                                _writer.WriteString(_reader.GetText());
                            }
                        }
                        else if (_reader.IsEndElement)
                        {
                            _writer.WriteEndElement();
                        }
                        _reader.Read();
                    }
                    _writer.WriteEndElement(); // close head row


                    InsertData(data);

                    _writer.WriteEndElement();//close sheet
                    //_workbookPart.Workbook.Save();
                    break;
                }
                else
                {
                    if (_reader.IsStartElement)
                    {
                        _writer.WriteStartElement(_reader);
                    }
                    else if (_reader.IsEndElement)
                    {
                        _writer.WriteEndElement();
                    }
                }
            }

            EndWritingWithTemplate();
        }

        public WorksheetPart CopySheet(string from, string to)
        {
            var original = GetOrCreateWorkSheetByName(from);
            var copy = GetOrCreateWorkSheetByName(to);

            var writer = OpenXmlWriter.Create(copy);

            if (original.Worksheet == null)
            {
                original.Worksheet = new Worksheet(new SheetData());
            }

            var reader = OpenXmlReader.Create(original);


            while (reader.Read())
            {
                if (reader.ElementType == typeof(SheetData))
                {
                    if (reader.IsEndElement)
                        continue;

                    //Reading template sheet data
                    writer.WriteStartElement(new SheetData());
                    reader.Read();

                    while (true)
                    {
                        if (reader.ElementType == typeof(Row) && reader.IsEndElement)
                            break;

                        if (reader.IsStartElement)
                        {
                            writer.WriteStartElement(reader);
                            if (reader.ElementType.IsSubclassOf(typeof(OpenXmlLeafTextElement)))
                            {
                                writer.WriteString(reader.GetText());
                            }
                        }
                        else if (reader.IsEndElement)
                        {
                            writer.WriteEndElement();
                        }
                        reader.Read();
                    }
                    writer.WriteEndElement(); // close head row

                    writer.WriteEndElement();//close sheet
                    //_workbookPart.Workbook.Save();
                    //break;
                }
                else
                {
                    if (reader.IsStartElement)
                    {
                        writer.WriteStartElement(reader);
                    }
                    else if (reader.IsEndElement)
                    {
                        writer.WriteEndElement();
                    }
                }
            }

            reader.Close();
            writer.Close();
            return copy;
        }

    }
}
