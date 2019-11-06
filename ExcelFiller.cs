using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;

namespace ExcelFiller
{
    public class ExcelFiller
    {
        private ExcelPackage pkg;
        private ExcelWorkbook wb;
        private ExcelWorksheet sheet;

        private Dictionary<string, ExcelNamedStyleXml> styles = new Dictionary<string, ExcelNamedStyleXml>();
        
        private void Process(string inputfile, string outputfile)
        {
            using (var reader = new StreamReader(inputfile))
            {
                using (pkg = new ExcelPackage())
                {
                    wb = pkg.Workbook;
                    sheet = null;
                    var line = reader.ReadLine();
                    while (line != null)
                    {
                        if (line != "" && !line.Trim().StartsWith("#"))
                        {
                            ExecuteCommand(ParseLine(line));
                        }
                        line = reader.ReadLine();
                    }

                    pkg.SaveAs(new FileInfo(outputfile));
                }

            }
        }

        private void ExecuteCommand(IReadOnlyList<string> command)
        {
            switch (command[0].ToLower())
            {
                case "create_sheet":
                    wb.Worksheets.Add(StringValue(command[1]));
                    sheet = wb.Worksheets[StringValue(command[1])];
                    break;

                case "select_sheet":
                    sheet = wb.Worksheets[StringValue(command[1])];
                    // TODO
                    break;

                case "set_cell_string":
                    sheet.Cells[StringValue(command[1])].Value = StringValue(command[2]);
                    break;

                case "set_cell_number":
                    sheet.Cells[StringValue(command[1])].Value = NumberValue(command[2]);
                    break;

                case "set_cell_formula":
                    sheet.Cells[StringValue(command[1])].Formula = FormulaValue(command[2]);
                    break;

                case "create_style":
                    styles[StringValue(command[1])] = wb.Styles.CreateNamedStyle(StringValue(command[1]));
                    break;

                case "set_style_fg_color":
                    styles[StringValue(command[1])].Style.Font.Color.SetColor(GetColor(command[2], command[3], command[4], command[5]));
                    break;

                case "set_style_bg_color":
                    styles[StringValue(command[1])].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    styles[StringValue(command[1])].Style.Fill.BackgroundColor.SetColor(GetColor(command[2], command[3], command[4], command[5]));
                    break;

                case "set_style_bold":
                    styles[StringValue(command[1])].Style.Font.Bold = true;
                    break;

                case "set_style_horz_align":
                    styles[StringValue(command[1])].Style.HorizontalAlignment = GetHorzAlign(command[2]);
                    break;

                case "set_style_vert_align":
                    styles[StringValue(command[1])].Style.VerticalAlignment = GetVertAlign(command[2]);
                    break;

                case "set_style_border":
                    switch (StringValue(command[2]).ToLower())
                    {
                        case "bottom": 
                            styles[StringValue(command[1])].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            break;
                        case "top": 
                            styles[StringValue(command[1])].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            break;
                        case "left": 
                            styles[StringValue(command[1])].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            break;
                        case "right": 
                            styles[StringValue(command[1])].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            break;
                    }
                    break;

                case "set_cell_style":
                    sheet.Cells[StringValue(command[1])].StyleName = StringValue(command[2]);
                    break;

                case "set_column_width":
                    if (StringValue(command[2]) == "auto")
                    {
                        sheet.Cells[StringValue(command[1])].AutoFitColumns();
                    }
                    else
                    {
                        sheet.Cells[StringValue(command[1])].AutoFitColumns(NumberValue(command[2]), NumberValue(command[2]));
                    }
                    break;

                case "set_filter":
                    sheet.Cells[StringValue(command[1])].AutoFilter = true;
                    break;
                
                case "freeze_pane":
                    // TODO
                    break;
                
                case "select_cell":
                    // TODO
                    break;
                
                case "calculate_sheet":
                    wb.Worksheets[StringValue(command[1])].Calculate();
                    break;
                
                default:
                    Console.WriteLine("Unknown command: " + command[0]);
                    break;
            }
        }

        private ExcelHorizontalAlignment GetHorzAlign(string align)
        {
            switch (align.ToLower())
            {
                case "left": return ExcelHorizontalAlignment.Left;
                case "right": return ExcelHorizontalAlignment.Right;
                case "center": return ExcelHorizontalAlignment.Center;
                case "justify": return ExcelHorizontalAlignment.Justify;
                case "fill": return ExcelHorizontalAlignment.Fill;
                case "distributed": return ExcelHorizontalAlignment.Distributed;
                default: return ExcelHorizontalAlignment.General;
            }
        }

        private ExcelVerticalAlignment GetVertAlign(string align)
        {
            switch (align.ToLower())
            {
                case "bottom": return ExcelVerticalAlignment.Bottom;
                case "center": return ExcelVerticalAlignment.Center;
                case "distributed": return ExcelVerticalAlignment.Distributed;
                case "justify": return ExcelVerticalAlignment.Justify;
                case "top": return ExcelVerticalAlignment.Top;
                default: return ExcelVerticalAlignment.Top;
            }
        }

        private Color GetColor(string r, string g, string b, string a)
        {
            return Color.FromArgb(
                IntValue(a),
                IntValue(r),
                IntValue(g),
                IntValue(b));
        }
        
        private string StringValue(string s)
        {
            if (s.StartsWith("\"") && s.EndsWith("\"")) return s.Substring(1, s.Length - 2);
            return s;
        }

        private double NumberValue(string s)
        {
            return double.Parse(s);
        }

        private int IntValue(string s)
        {
            return int.Parse(s);
        }

        private string FormulaValue(string s)
        {
            return StringValue(s).Substring(1);
        }
        
        private List<string> ParseLine(string line)
        {
            var tokens = new List<string>();
            var current = new StringBuilder();
            var in_string = false;
            var escape = false;
            for (var i = 0; i < line.Length; i++)
            {
                var ch = line[i];
                if (ch == ' ' && !in_string && current.Length > 0 )
                {
                    tokens.Add(current.ToString());
                    current.Clear();
                }
                else if (ch == '\"' && !in_string)
                {
                    current.Append(ch);
                    in_string = true;
                }
                else if (ch == '\"' && in_string && escape)
                {
                    current.Append(ch);
                    escape = false;
                }
                else if (ch == '\\' && in_string && !escape)
                {
                    escape = true;
                }
                else if (ch == '\"' && in_string && !escape)
                {
                    current.Append(ch);
                    in_string = false;
                }
                else
                {
                    current.Append(ch);
                }
            }
            if (current.Length > 0) tokens.Add(current.ToString());
            return tokens;
        }

        public static void Main(string[] args)
        {
            new ExcelFiller().Process(args[0], args[1]);
        }
    }
}