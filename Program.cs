using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;

namespace XlsxToConfig
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine(Environment.CurrentDirectory);
            var xlsxPath = string.Empty;
            var savePath = string.Empty;
            var section = string.Empty;
            if (args.Length < 1) {
                Console.Write("输入保存格式(lua, xml, json, ...)：");
                section = Console.ReadLine();
            } else {
                section = args[0];
            }

            // 读取配置文件
            var iniPath = Path.Combine(Environment.CurrentDirectory, "config.ini");
            if (File.Exists(iniPath)) {
                var convertor = new XlsxConvertor(iniPath, section);
                xlsxPath = convertor.readPath;
                savePath = convertor.savePath;

                if (!Directory.Exists(savePath)) {
                    Console.WriteLine("Save Path: {0} NOT exist!", savePath);
                } else {
                    savePath = Path.GetFullPath(savePath);
                    if (File.Exists(xlsxPath)) {
                        xlsxPath = Path.GetFullPath(xlsxPath);
                        ConvertXlsx(convertor, xlsxPath, Path.GetDirectoryName(xlsxPath), savePath);
                    } else if (Directory.Exists(xlsxPath)) {
                        xlsxPath = Path.GetFullPath(xlsxPath);
                        foreach (var path in Directory.GetFiles(xlsxPath, "*", SearchOption.AllDirectories)) {
                            ConvertXlsx(convertor, Path.GetFullPath(path), xlsxPath, savePath);
                        }
                    } else {
                        Console.WriteLine("Source Path: {0} NOT exist!", xlsxPath);
                    }
                }
            } else {
                Console.WriteLine("未找到配置文件!");
            }
                
            Console.WriteLine("Press ENTER to exit.");
            Console.ReadLine();
        }

        private class XlsxConvertor
        {
            /// <summary>
            /// 
            /// </summary>
            /// <param name="section"></param>
            /// <param name="key"></param>
            /// <param name="def"></param>
            /// <param name="retVal"></param>
            /// <param name="size"></param>
            /// <param name="filePath"></param>
            /// <returns>返回取得字符串缓冲区的长度</returns>
            [System.Runtime.InteropServices.DllImport("kernel32")]
            private static extern long GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

            private string m_ReadPath;
            private string m_SavePath;
            private string m_FileExt = "lua";

            private int m_PlatRowNum = 0;
            private int m_TypeRowNum = 1;
            private int m_HeaderRowNum = 2;
            private char[] m_ArrSeparator = new char[] { '|' };

            private string m_TableBegin = "{";
            private string m_TableEnd = "}";
            private string m_LineBegin = "{";
            private string m_LineEnd = "},";
            private string m_CellFmt = "{0}={1},";
            private string m_StrQoute = "\"";
            private string m_BoolTRUE = "true";
            private string m_BoolFALSE = "false";
            private string m_ArrBegin = "{";
            private string m_ArrElm = "{0},";
            private string m_ArrEnd = "}";
            private string m_NullValue = "nil";

            public string readPath { get { return m_ReadPath; } }
            public string savePath { get { return m_SavePath; } }
            public string fileExt { get { return m_FileExt; } }

            private static string ReadIniString(string section, string key, StringBuilder temp, string iniPath)
            {
                temp.Clear();
                GetPrivateProfileString(section, key, null, temp, temp.Capacity, iniPath);
                return temp.ToString();
            }

            private static int ReadIniInteger(string section, string key, int def, StringBuilder temp, string iniPath)
            {
                temp.Clear();
                GetPrivateProfileString(section, key, null, temp, temp.Capacity, iniPath);

                int ret;
                if (int.TryParse(temp.ToString(), out ret)) {
                    return ret;
                }
                return def;
            }

            public XlsxConvertor(string iniPath, string section)
            {
                var temp = new StringBuilder(1024);
                m_ReadPath = ReadIniString("GLOBAL", "xlsx", temp, iniPath);
                m_SavePath = ReadIniString(section, "save", temp, iniPath);
                m_FileExt = ReadIniString(section, "ext", temp, iniPath);
                m_PlatRowNum = ReadIniInteger(section, "PlatRow", -1, temp, iniPath);
                m_TypeRowNum = ReadIniInteger(section, "TypeRow", -1, temp, iniPath);
                m_HeaderRowNum = ReadIniInteger(section, "HeaderRow", -1, temp, iniPath);
                m_TableBegin = ReadIniString(section, "TableBegin", temp, iniPath);
                m_TableEnd = ReadIniString(section, "TableEnd", temp, iniPath);
                m_LineBegin = ReadIniString(section, "LineBegin", temp, iniPath);
                m_LineEnd = ReadIniString(section, "LineEnd", temp, iniPath);
                m_CellFmt = ReadIniString(section, "CellFmt", temp, iniPath);
                m_StrQoute = ReadIniString(section, "StrQoute", temp, iniPath);
                m_ArrBegin = ReadIniString(section, "ArrBegin", temp, iniPath);
                m_ArrElm = ReadIniString(section, "ArrElm", temp, iniPath);
                m_ArrEnd = ReadIniString(section, "ArrEnd", temp, iniPath);
                m_BoolTRUE = ReadIniString(section, "BoolTRUE", temp, iniPath);
                m_BoolFALSE = ReadIniString(section, "BoolFALSE", temp, iniPath);
                m_NullValue = ReadIniString(section, "NullValue", temp, iniPath);
            }
            
            private void BuildCell(StringBuilder strbld, string key, ICell cell, string cellType)
            {
                if (cell == null) return;

                switch (cellType) {
                    case "bool":
                        switch (cell.CellType) {
                            case CellType.Boolean:
                                strbld.AppendFormat(m_CellFmt, key, cell.BooleanCellValue ? m_BoolTRUE : m_BoolFALSE);
                                break;
                            case CellType.Numeric:
                                strbld.AppendFormat(m_CellFmt, key, cell.NumericCellValue != 0 ? m_BoolTRUE : m_BoolFALSE);
                                break;
                            default:
                                strbld.AppendFormat(m_CellFmt, key, !string.IsNullOrEmpty(cell.ToString()) ? m_BoolTRUE : m_BoolFALSE);
                                break;
                        }
                        break;
                    case "int":
                        switch (cell.CellType) {
                            case CellType.Boolean:
                                strbld.AppendFormat(m_CellFmt, key, cell.BooleanCellValue ? 1 : 0);
                                break;
                            case CellType.Numeric:
                                strbld.AppendFormat(m_CellFmt, key, cell);
                                break;
                            default:
                                var strValue = cell.ToString();
                                int value = 0;
                                if (int.TryParse(strValue, out value)) {
                                    strbld.AppendFormat(m_CellFmt, key, value);
                                }
                                break;
                        }
                        break;
                    case "string":
                        if (!string.IsNullOrEmpty(m_StrQoute)) {
                            strbld.AppendFormat(m_CellFmt, key, m_StrQoute + cell.ToString() + m_StrQoute);
                        } else {
                            strbld.AppendFormat(m_CellFmt, key, cell);
                        }
                        break;
                    case "int[]": {
                            var arrbld = new StringBuilder(m_ArrBegin);
                            foreach (var elm in cell.ToString().Split(m_ArrSeparator)) {
                                var value = 0;
                                if (int.TryParse(elm, out value)) {
                                    arrbld.AppendFormat(m_ArrElm, value);
                                } else {
                                    arrbld.AppendFormat(m_ArrElm, m_NullValue);
                                }
                            }
                            arrbld.Append(m_ArrEnd);
                            strbld.AppendFormat(m_CellFmt, key, arrbld.ToString());
                        }
                        break;
                    case "string[]": {
                            var arrbld = new StringBuilder(m_ArrBegin);
                            foreach (var elm in cell.ToString().Split(m_ArrSeparator)) {
                                arrbld.AppendFormat(m_ArrElm, string.IsNullOrEmpty(elm) ? m_NullValue : m_StrQoute + elm + m_StrQoute);
                            }
                            arrbld.Append(m_ArrEnd);
                            strbld.AppendFormat(m_CellFmt, key, arrbld.ToString());
                        }
                        break;
                    default: break;
                }
            }

            public string BuildTable(ISheet sheet)
            {
                var strbld = new StringBuilder();
                strbld.Append(m_TableBegin);
                var platforms = sheet.GetRow(m_PlatRowNum);
                var types = sheet.GetRow(m_TypeRowNum);
                var header = sheet.GetRow(m_HeaderRowNum);

                // TODO 检查格式

                var firstCellNum = header.FirstCellNum;
                var lastCellNum = header.LastCellNum;

                for (var i = sheet.FirstRowNum + m_HeaderRowNum + 1; i <= sheet.LastRowNum; ++i) {
                    var row = sheet.GetRow(i);
                    strbld.Append(m_LineBegin);
                    for (var j = firstCellNum; j < lastCellNum; ++j) {
                        var plat = platforms.GetCell(j);
                        if (plat == null || string.IsNullOrEmpty(plat.StringCellValue)) continue;

                        var type = types.GetCell(j);
                        if (type == null || string.IsNullOrEmpty(type.StringCellValue)) continue;

                        var cell = header.GetCell(j);
                        if (cell == null) continue;

                        BuildCell(strbld, cell.StringCellValue, row.GetCell(j), type.StringCellValue);
                    }
                    strbld.AppendLine(m_LineEnd);
                }
                strbld.Append(m_TableEnd);
                return strbld.ToString();
            }
        }

        private static void ConvertXlsx(XlsxConvertor convertor, string filePath, string readPath, string savePath)
        {
            var fileExt = Path.GetExtension(filePath).ToLower();
            FileStream fs = null;
            IWorkbook workbook = null;
            switch (fileExt) {
                case ".xlsx":
                    fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                    workbook = new XSSFWorkbook(fs);
                    break;
                case ".xls":
                    fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                    workbook = new HSSFWorkbook(fs);
                    break;
            }
            if (workbook == null) return;

            fs.Dispose();

            var fileName = filePath.Substring(readPath.Length + 1).ToLower();
            fileName = fileName.Substring(0, fileName.LastIndexOf('.'));
            for (var i = 0; i < workbook.NumberOfSheets; ++i) {
                var sheetName = workbook.GetSheetName(i);
                if (sheetName != null && sheetName.StartsWith("@")) {
                    sheetName = sheetName.Substring(1).ToLower();
                    var saveName = fileName + "_" + sheetName + "." + convertor.fileExt;
                    var saveFilePath = Path.Combine(savePath, saveName);
                    var saveFileDir = Path.GetDirectoryName(saveFilePath);
                    if (!Directory.Exists(saveFileDir)) Directory.CreateDirectory(saveFileDir);

                    using (var f = File.CreateText(Path.Combine(savePath, saveName))) {
                        f.Write(convertor.BuildTable(workbook.GetSheetAt(i)));
                    }
                }
            }
        }
    }
}
