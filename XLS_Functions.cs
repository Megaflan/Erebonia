using Erebonia.Converters;
using Erebonia.Formats;
using OfficeOpenXml;
using SpreadsheetLight;
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
                ExcelWorksheet worksheet = excel.Workbook.Worksheets[0];
                XLS xls = new XLS();
                var c = 0;
                for (int i = 1; i < worksheet.Dimension.Rows; i++)
                {
                    for (int j = 1; j < worksheet.Dimension.Columns; j++)
                    {
                        if (worksheet.Cells[i, j].Value != null)
                        {
                            string value = worksheet.Cells[i, j].Value.ToString();
                            if (value == "dialog")
                            {
                                xls.Entries.Add(new XLS.Entry
                                {
                                    ID = (uint)c,
                                    Row = (uint)i,
                                    Column = (uint)j,
                                    Text = worksheet.Cells[i + 1, j].Value.ToString(),
                                });
                                c++;
                            }
                        }                        
                    }
                }
                GeneratePo(xls, lang);
            }
        }

        static private void GeneratePo(XLS xls, string lang)
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
                po.Add(new PoEntry(entry.Text)
                {
                    Context = entry.ID.ToString(),
                    ExtractedComments = $"{(char)asciiColumn}{entry.Row + 1}",
                });
            }

            var node = NodeFactory.FromMemory("node");
            node.TransformWith(new Po2BinaryEasy()
            {
                PoPassed = po
            }).TransformWith(new Po2Binary()).Stream.WriteTo(Path.GetDirectoryName(fileName) + "\\" + Path.GetFileNameWithoutExtension(fileName) + "_dialog.po");
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
