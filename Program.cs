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

            private static char[] s_ArrSeparator = new char[] { '|' };

            class Scope
            {
                public readonly string Begin, End, Separator;
                public Scope(string b, string e, string s)
                {
                    Begin = b; End = e; Separator = s;
                }

                public Scope(string value)
                {
                    value = value.Replace("\\n", "\n").Replace("\\t", "\t");
                    var values = value.Split(s_ArrSeparator);
                    if (values.Length > 0) Begin = values[0];
                    if (values.Length > 1) Separator = values[1];
                    if (values.Length > 2) End = values[2];
                }
            }

            private string m_ReadPath;
            private string m_SavePath;
            private string m_FileExt = "lua";

            private int m_PlatRowNum = 0;
            private int m_TypeRowNum = 1;
            private int m_HeaderRowNum = 2;

            private Scope m_Table = new Scope("{\n", "\n}", null);
            private Scope m_Line = new Scope("{", "}", ",");
            private Scope m_Array = new Scope("{", "}", ",");
            private Scope m_Cell = new Scope("{0}={1}", null, ",");

            private string m_StrQoute = "\"";
            private string m_BoolTRUE = "true";
            private string m_BoolFALSE = "false";
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

            private static Scope ReadIniScope(string section, string key, StringBuilder temp, string iniPath)
            {                
                var tableValue = ReadIniString(section, key, temp, iniPath);
                if (!string.IsNullOrEmpty(tableValue)) {
                    return new Scope(tableValue);
                }
                return null;
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

                m_Table = ReadIniScope(section, "Table", temp, iniPath) ?? m_Table;
                m_Line = ReadIniScope(section, "Line", temp, iniPath) ?? m_Line;
                m_Array = ReadIniScope(section, "Array", temp, iniPath) ?? m_Array;
                m_Cell = ReadIniScope(section, "Cell", temp, iniPath) ?? m_Cell;

                m_StrQoute = ReadIniString(section, "StrQoute", temp, iniPath);
                m_BoolTRUE = ReadIniString(section, "BoolTRUE", temp, iniPath);
                m_BoolFALSE = ReadIniString(section, "BoolFALSE", temp, iniPath);
                m_NullValue = ReadIniString(section, "NullValue", temp, iniPath);
            }
            
            private void BuildCell(StringBuilder strbld, bool firstCell, string key, ICell cell, string cellType)
            {
                if (cell == null) return;
                
                if (!firstCell) strbld.Append(m_Cell.Separator);
                var cellFmt = m_Cell.Begin;
                switch (cellType) {
                    case "bool":
                        switch (cell.CellType) {
                            case CellType.Boolean:
                                strbld.AppendFormat(cellFmt, key, cell.BooleanCellValue ? m_BoolTRUE : m_BoolFALSE);
                                break;
                            case CellType.Numeric:
                                strbld.AppendFormat(cellFmt, key, cell.NumericCellValue != 0 ? m_BoolTRUE : m_BoolFALSE);
                                break;
                            default:
                                strbld.AppendFormat(cellFmt, key, !string.IsNullOrEmpty(cell.ToString()) ? m_BoolTRUE : m_BoolFALSE);
                                break;
                        }
                        break;
                    case "int":
                        switch (cell.CellType) {
                            case CellType.Boolean:
                                strbld.AppendFormat(cellFmt, key, cell.BooleanCellValue ? 1 : 0);
                                break;
                            case CellType.Numeric:
                                strbld.AppendFormat(cellFmt, key, cell);
                                break;
                            default:
                                var strValue = cell.ToString();
                                int.TryParse(strValue, out int value);
                                strbld.AppendFormat(cellFmt, key, value);
                                break;
                        }
                        break;
                    case "string":
                        if (!string.IsNullOrEmpty(m_StrQoute)) {
                            strbld.AppendFormat(cellFmt, key, m_StrQoute + cell.ToString() + m_StrQoute);
                        } else {
                            strbld.AppendFormat(cellFmt, key, cell);
                        }
                        break;
                    case "int[]": {
                            var arrbld = new StringBuilder(m_Array.Begin);
                            var array = cell.ToString().Split(s_ArrSeparator);
                            for(var i = 0; i < array.Length; ++i) {
                                var elm = array[i];
                                if (i > 0) arrbld.Append(m_Array.Separator);
                                if (int.TryParse(elm, out int value)) {
                                    arrbld.Append(value);
                                } else {
                                    if (array.Length == 1) break;
                                    arrbld.Append(m_NullValue);
                                }
                            }
                            arrbld.Append(m_Array.End);
                            strbld.AppendFormat(cellFmt, key, arrbld.ToString());
                        }
                        break;
                    case "string[]": {
                            var arrbld = new StringBuilder(m_Array.Begin);
                            var array = cell.ToString().Split(s_ArrSeparator);
                            for (var i = 0; i < array.Length; ++i) {
                                var elm = array[i];
                                if (i > 0) arrbld.Append(m_Array.Separator);
                                if (string.IsNullOrEmpty(elm)) {
                                    if (array.Length == 1) break;
                                    arrbld.Append(m_NullValue);
                                } else {
                                    arrbld.Append(m_StrQoute + elm + m_StrQoute);
                                }
                            }
                            arrbld.Append(m_Array.End);
                            strbld.AppendFormat(cellFmt, key, arrbld.ToString());
                        }
                        break;
                    default: strbld.Append(cellType); break;
                }
            }

            public string BuildTable(ISheet sheet)
            {
                var strbld = new StringBuilder();
                strbld.Append(m_Table.Begin);
                var platforms = sheet.GetRow(m_PlatRowNum);
                var types = sheet.GetRow(m_TypeRowNum);
                var header = sheet.GetRow(m_HeaderRowNum);

                // TODO 检查格式

                var firstCellNum = header.FirstCellNum;
                var lastCellNum = header.LastCellNum;

                var last = sheet.LastRowNum;
                for (var i = sheet.FirstRowNum + m_HeaderRowNum + 1; i <= last; ++i) {
                    var row = sheet.GetRow(i);
                    strbld.Append(m_Line.Begin);                    
                    var first = true;
                    for (var j = firstCellNum; j < lastCellNum; ++j) {
                        var plat = platforms.GetCell(j);
                        if (plat == null || string.IsNullOrEmpty(plat.StringCellValue)) continue;

                        var type = types.GetCell(j);
                        if (type == null || string.IsNullOrEmpty(type.StringCellValue)) continue;

                        var cell = header.GetCell(j);
                        if (cell == null) continue;

                        BuildCell(strbld, first, cell.StringCellValue, row.GetCell(j), type.StringCellValue);
                        first = false;
                    }
                    strbld.Append(m_Line.End);
                    if (i < last) strbld.AppendLine(m_Line.Separator);
                }
                strbld.Append(m_Table.End);
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
