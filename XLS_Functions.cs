using Erebonia.Converters;
using Erebonia.Formats;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Yarhl.FileFormat;
using Yarhl.FileSystem;
using Yarhl.IO;
using Yarhl.Media.Text;

namespace Erebonia
{
    class XLS_Functions
    {
        static string fileName = string.Empty;

        static public void Export(string file, string lang)
        {
            fileName = file;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excel = new ExcelPackage(new FileInfo(file)))
            {
                ExcelWorksheet worksheet = excel.Workbook.Worksheets["Sheet1"];
                XLS diagXls = new XLS();
                XLS strXls = new XLS();
                var diagQty = 0;
                var strQty = 0;
                for (int i = 1; i < worksheet.Dimension.Rows; i++)
                {
                    for (int j = 1; j < worksheet.Dimension.Columns; j++)
                    {
                        if (worksheet.Cells[i, j].Value != null)
                        {
                            string value = worksheet.Cells[i, j].Value.ToString();
                            switch (value)
                            {
                                case "dialog":
                                    if (worksheet.Cells[i + 1, j].Value != null)
                                    {
                                        diagXls.Entries.Add(new XLS.Entry
                                        {
                                            ID = (uint)diagQty,
                                            Row = (uint)i,
                                            Column = (uint)j,
                                            Text = worksheet.Cells[i + 1, j].Value.ToString(),
                                        });
                                        diagQty++;
                                    }                                    
                                    break;
                                case "string":
                                    if (worksheet.Cells[i + 1, j].Value != null)
                                    {
                                        strXls.Entries.Add(new XLS.Entry
                                        {
                                            ID = (uint)strQty,
                                            Row = (uint)i,
                                            Column = (uint)j,
                                            Text = worksheet.Cells[i + 1, j].Value.ToString(),
                                        });
                                        strQty++;
                                    }
                                    break;
                            }
                        }                        
                    }
                }
                GeneratePo(diagXls, lang, true);
                GeneratePo(strXls, lang, false);
            }
        }

        static private void GeneratePo(XLS xls, string lang, bool mode)
        {
            var po = new Po
            {
                Header = new PoHeader("Trails in the Cold Steel", "tradusquare@gmail.com", lang)
                {
                    LanguageTeam = lang,
                }
            };

            foreach (var entry in xls.Entries)
            {
                uint asciiColumn = entry.Column + 64;
                if (!string.IsNullOrEmpty(entry.Text))
                {
                    po.Add(new PoEntry(entry.Text)
                    {
                        Context = entry.ID.ToString(),
                        ExtractedComments = $"{(char)asciiColumn}{entry.Row + 1}",
                    });
                }                
            }

            var node = NodeFactory.FromMemory("node");
            switch (mode)
            {
                case true:
                    node.TransformWith(new Po2BinaryEasy()
                    {
                        PoPassed = po
                    }).TransformWith(new Po2Binary()).Stream.WriteTo(Path.GetFileNameWithoutExtension(fileName) + "_dialog.po");
                    break;
                case false:
                    node.TransformWith(new Po2BinaryEasy()
                    {
                        PoPassed = po
                    }).TransformWith(new Po2Binary()).Stream.WriteTo(Path.GetFileNameWithoutExtension(fileName) + "_string.po");
                    break;
            }            
        }

        static public void Import(string poFile, string xlsFile)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var po = NodeFactory.FromFile(poFile).TransformWith(new Binary2Po()).GetFormatAs<Po>();

            using (var excel = new ExcelPackage(new FileInfo(xlsFile)))
            {
                ExcelWorksheet worksheet = excel.Workbook.Worksheets[0];
                foreach (var poEntry in po.Entries)
                {
                    var text = poEntry.Text == "<!empty>" ? string.Empty : poEntry.Text;
                    worksheet.Cells[poEntry.ExtractedComments].Value = text;
                }
                excel.Save();
            }            
        }
    }
}
