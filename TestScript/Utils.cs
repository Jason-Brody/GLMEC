using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using TestScript.Attributes;
using Ex = Microsoft.Office.Interop.Excel;

namespace TestScript
{
    public static class Utils
    {

        public static float GetAmount(string val)
        {
            val.Replace(",", "");
            var pos = val.IndexOf('-');
            if (val.Contains("-") && val.IndexOf('-') == val.Length - 1)
            {
                val = "-" + val.Substring(0, val.Length - 1);
            }
            return float.Parse(val);
        }

        public static DataTable ReadStringToTable(string filePath, Func<string, string, List<string>> LineFunc)
        {
            string tempString = "";
            DataTable table = null;
            string headerRow = "";
            using (StreamReader sr = new StreamReader(filePath))
            {
                while (!sr.EndOfStream)
                {
                    tempString = sr.ReadLine();
                    var vals = LineFunc(tempString, headerRow);
                    if (vals != null && vals.Count() > 0)
                    {
                        if (table == null)
                        {
                            table = new DataTable();
                            Dictionary<string, int> cols = new Dictionary<string, int>();
                            headerRow = tempString;
                            for (int i = 0; i < vals.Count(); i++)
                            {
                                if (vals[i] == "")
                                {
                                    vals[i] = "Header_Temp_" + i.ToString();
                                }
                                while (cols.ContainsKey(vals[i]))
                                {
                                    vals[i] = vals[i] + "1";
                                }
                                cols.Add(vals[i], 0);

                                table.Columns.Add(vals[i]);
                            }
                        }
                        else
                        {
                            DataRow dr = table.NewRow();
                            for (int i = 0; i < vals.Count(); i++)
                            {
                                dr[i] = vals[i];
                            }
                            table.Rows.Add(dr);
                        }
                    }
                }
            }
            return table;
        }

        public static string FillNumber(string val, int length = 10)
        {
            if (val.Length < length)
            {
                while (val.Length < length)
                {
                    val = "0" + val;
                }

            }
            return val;
        }

        public static void Export<T>(this IEnumerable<T> Data, string fileName, string splitChar) where T : class
        {
            List<PropertyInfo> props = typeof(T).GetProperties().Where(p => (p.PropertyType == typeof(string) || p.PropertyType.IsValueType) && p.DeclaringType.IsPublic).ToList();


            using (StreamWriter sw = new StreamWriter(fileName, false))
            {
                string line = "";
                foreach (var prop in props)
                {
                    var attr = prop.GetCustomAttribute<AliasAttribute>();
                    if (attr != null)
                        line += attr.Name + splitChar;
                    else
                        line += prop.Name + splitChar;
                }
                line = line.Substring(0, line.Length - 1);
                sw.WriteLine(line);
                foreach (var item in Data)
                {
                    line = "";
                    foreach (var p in props)
                    {
                        var val = p.GetValue(item);
                        if (val == null)
                            val = "";

                        line += val.ToString() + splitChar;
                    }
                    line = line.Substring(0, line.Length - 1);
                    sw.WriteLine(line);
                }

            }
        }

        public static void ExportToExcel<T>(this IEnumerable<T> Data, string fileName, string sheetName, Action<object> otherAction) where T : class
        {
           

            List<PropertyInfo> props = typeof(T).GetProperties().Where(p => (p.PropertyType == typeof(string) || p.PropertyType.IsValueType) && p.DeclaringType.IsPublic).ToList();
            Ex.Application app = new Ex.Application();
            app.Visible = true;

            var wbs = app.Workbooks;
            var wb = wbs.Add();

            var sheet = wb.ActiveSheet as Ex.Worksheet;

            sheet.Name = sheetName;

            Ex.Range range = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1 + Data.Count(), props.Count]];

            try
            {

                object[,] datas = new object[Data.Count() + 1, props.Count];

                List<Tuple<string, int>> formualList = new List<Tuple<string, int>>();


                for (int i = 0; i < props.Count; i++)
                {
                    var attr = props[i].GetCustomAttribute<AliasAttribute>();
                    if (attr != null)
                        datas[0, i] = attr.Name;
                    else
                        datas[0, i] = props[i].Name;

                    var formulaAttr = props[i].GetCustomAttribute<ExcelFormulaAttribute>();
                    if (formulaAttr != null)
                    {
                        Tuple<string, int> item = new Tuple<string, int>(formulaAttr.Formula, i + 1);
                        formualList.Add(item);
                    }
                }

                for (int i = 0; i < Data.Count(); i++)
                {
                    for (int j = 0; j < props.Count; j++)
                    {
                        var val = props[j].GetValue(Data.ElementAt(i));
                        datas[i + 1, j] = val;
                    }
                }


               
                range.Value = datas;


                for (int i = 0; i < props.Count; i++)
                {
                    range = sheet.Cells[1, i + 1];
                    
                    var attr = props[i].GetCustomAttribute<ExcelHeaderStyleAttribute>();
                    if (attr != null)
                    {
                        range.Interior.Color = attr.BackgroundColor;
                        range.EntireColumn.NumberFormat = attr.NumberFormat;
                        range.ColumnWidth = attr.Width;
                        range.Font.Bold = attr.IsFontBold;
                        range.Font.Size = attr.FontSize;
                        range.WrapText = attr.IsTextWrap;
                        range.HorizontalAlignment = attr.HAlign;
                        range.VerticalAlignment = attr.VAlign;
                    }
                }



                if (formualList.Count > 0)
                {
                    foreach (var item in formualList)
                    {
                        range = sheet.Cells[2, item.Item2];
                        range.Formula = item.Item1;
                        range.AutoFill(sheet.Range[sheet.Cells[2, item.Item2], sheet.Cells[Data.Count() + 1, item.Item2]], Ex.XlAutoFillType.xlFillDefault);
                    }

                }


                otherAction(sheet);

                range = sheet.UsedRange;

                range.Borders[Ex.XlBordersIndex.xlEdgeLeft].LineStyle = Ex.XlLineStyle.xlContinuous;
                range.Borders[Ex.XlBordersIndex.xlEdgeTop].LineStyle = Ex.XlLineStyle.xlContinuous;
                range.Borders[Ex.XlBordersIndex.xlEdgeRight].LineStyle = Ex.XlLineStyle.xlContinuous;
                range.Borders[Ex.XlBordersIndex.xlEdgeBottom].LineStyle = Ex.XlLineStyle.xlContinuous;
                range.Borders.Color = ConsoleColor.Black;


                range = sheet.Cells[1, 1] as Ex.Range;
                range.Select();


                if (File.Exists(fileName))
                    File.Delete(fileName);
                wb.SaveAs(fileName);


                wb.Close();
                app.Quit();

            }

            catch (Exception ex)
            {
                throw (ex);
            }

            finally
            {
                Marshal.ReleaseComObject(range);
                range = null;
                Marshal.ReleaseComObject(sheet);
                sheet = null;
                Marshal.ReleaseComObject(wb);
                wb = null;
                Marshal.ReleaseComObject(wbs);
                wbs = null;
                Marshal.ReleaseComObject(app);
                app = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

        }

        //public static void ExportToExcel<T>(this IEnumerable<T> Data,string fileName,string sheetName)
        //{
        //    var doc = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);
        //    var wbPart = doc.AddWorkbookPart();
        //    wbPart.Workbook = new Workbook();
        //    doc.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
        //    WorksheetPart wsPart = wbPart.AddNewPart<WorksheetPart>();
        //    wsPart.Worksheet = new Worksheet(new SheetData());
        //    var sheet = new Sheet()
        //    {
        //        Id = wbPart.GetIdOfPart(wsPart),
        //        Name = sheetName,
        //        SheetId = 1
        //    };
        //    wbPart.Workbook.GetFirstChild<Sheets>().Append(sheet);

        //    SheetData sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
        //    Row headerRow = new Row();
        //    headerRow.RowIndex = 1;

        //    List<PropertyInfo> props = typeof(T).GetProperties().Where(p => (p.PropertyType == typeof(string) || p.PropertyType.IsValueType) && p.DeclaringType.IsPublic).ToList();

        //    for(int i=0;i<props.Count;i++)
        //    {
        //        Cell c = new Cell();
        //        c.DataType = new EnumValue<CellValues>(CellValues.String);
        //        c.CellReference = getColumnName(i + 1) + "1";
        //    }

        //    //var stylesPart = doc.WorkbookPart.AddNewPart<WorkbookStylesPart>();

        //    //stylesPart.Stylesheet = new Stylesheet();



        //}

        //private static string getColumnName(int columnIndex)
        //{
        //    int dividend = columnIndex;
        //    string columnName = String.Empty;
        //    int modifier;

        //    while (dividend > 0)
        //    {
        //        modifier = (dividend - 1) % 26;
        //        columnName = Convert.ToChar(65 + modifier).ToString() + columnName;
        //        dividend = (int)((dividend - modifier) / 26);
        //    }

        //    return columnName;
        //}


    }
}
