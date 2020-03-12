using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace JingruiZhang.EPPlus
{
    /// <summary>
    /// EPPlus 处理 Excel 常用工具
    /// </summary>
    public class ZExcelHelper
    {
        /// <summary>
        /// 根据流还原 ExcelPackage 对象
        /// </summary>
        /// <param name="file_inputStream">只读取 Stream，不考虑其创建与释放</param>
        /// <returns>返回的对象目前案例外部未释放</returns>
        [Obsolete("建议直接使用 GetFirstWorkSheet")]
        public static ExcelPackage GetExcelPackage(Stream file_inputStream)
        {
            return new ExcelPackage(file_inputStream);
        }

        /// <summary>
        /// 根据流还原 ExcelWorkbook 对象
        /// </summary>
        /// <param name="file_inputStream">只读取 Stream，不考虑其创建与释放</param>
        /// <returns>返回的对象目前案例外部未释放</returns>
        [Obsolete("建议直接使用 GetFirstWorkSheet")]
        public static ExcelWorkbook GetExcelWorkBook(Stream file_inputStream)
        {
            var package = GetExcelPackage(file_inputStream);
            return package.Workbook;
        }

        /// <summary>
        /// 根据流还原 ExcelWorkbook 对象，并获取第一个 ExcelWorksheet
        /// </summary>
        /// <param name="file_inputStream">只读取 Stream，不考虑其创建与释放</param>
        /// <returns>返回的对象目前案例外部未释放</returns>
        public static ExcelWorksheet GetFirstWorkSheet(Stream file_inputStream)
        {
            ExcelWorkbook workBook = GetExcelWorkBook(file_inputStream);
            var worksheet = workBook.Worksheets.First();
            return worksheet;
        }

        /// <summary>
        /// 获取 worksheet 总列数
        /// </summary>
        /// <param name="sheet">ExcelWorksheet 对象</param>
        /// <returns>总列数（编辑过的列）</returns>
        public static int GetWorkSheetColumnCount(ExcelWorksheet sheet)
        {
            int cols = sheet.Dimension.End.Column;
            return cols;
        }

        /// <summary>
        /// 获取 worksheet 总行数
        /// </summary>
        /// <param name="sheet">ExcelWorksheet 对象</param>
        /// <returns>总列数（编辑过的列）</returns>
        public static int GetWorkSheetRowCount(ExcelWorksheet sheet)
        {
            int rows = sheet.Dimension.End.Row;
            return rows;
        }

        /// <summary>
        /// 读取某行某列的值
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row">从1开始，不是以0为起点</param>
        /// <param name="col">从1开始，不是以0为起点</param>
        /// <returns></returns>
        public static object GetValue(ExcelWorksheet worksheet, int row, int col)
        {
            return worksheet.Cells[row, col].Value;
        }

        /// <summary>
        /// 获取所有图片（注：ExcelDrawing.From.Row 是图片的左上角行数 - 1）
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        public static ExcelDrawings GetAllDrawings(ExcelWorksheet worksheet)
        {
            return worksheet.Drawings;
        }

        /// <summary>
        /// 解析表格成为内存对象集合（取值时会将类型的属性按字母正序，之后的序号对应每个列）
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet"></param>
        /// <param name="skipLineCount">跳过的行数，一般为1（只有一行表头）</param>
        /// <param name="takeColumnCount">取的列数，一般表格有多少列，取多少列</param>
        /// <returns></returns>
        public static List<T> AnalysisExcelToObjectList<T>(ExcelWorksheet worksheet, int skipLineCount, int takeColumnCount)
            where T : class, new()
        {
            List<T> toret = new List<T>();
            for (int i = 1 + skipLineCount; i <= worksheet.Dimension.End.Row; i++)
            {
                T retobj = new T();
                var prop = retobj.GetType().GetProperties();
                var proplist = prop.OrderBy(x => x.Name).ToList();
                if (proplist.Count < takeColumnCount)
                {
                    throw new Exception("指定类的属性个数小于 takeColumnCount");
                }
                for (int j = 1; j <= takeColumnCount; j++)
                {
                    var val = GetValue(worksheet, i, j);

                    // Contains("DateTime")
                    SafeSetValue(proplist[j - 1], retobj, val);
                    //proplist[j].SetValue(retobj, val);
                }
                toret.Add(retobj);
            }
            return toret;
        }

        private static void SafeSetValue<T>(PropertyInfo propertyInfo, T retobj, object val) where T : class, new()
        {
            if (propertyInfo.PropertyType.FullName.Contains("System.Double")
              && propertyInfo.PropertyType.FullName.Contains("Nullable"))
            {
                #region ...
                if (val == null)
                {
                    propertyInfo.SetValue(retobj, null);
                }
                else
                {
                    string valstr = val.ToString();
                    double db;
                    if (double.TryParse(valstr, out db))
                    {
                        propertyInfo.SetValue(retobj, db);
                    }
                    else
                    {
                        propertyInfo.SetValue(retobj, null);
                    }
                }
                #endregion
            }
            else if (propertyInfo.PropertyType.FullName.Contains("System.Double"))
            {
                #region ...
                if (val == null)
                {
                    propertyInfo.SetValue(retobj, 0);
                }
                else
                {
                    string valstr = val.ToString();
                    double db;
                    if (double.TryParse(valstr, out db))
                    {
                        propertyInfo.SetValue(retobj, db);
                    }
                    else
                    {
                        propertyInfo.SetValue(retobj, 0);
                    }
                }
                #endregion
            }
            else if (propertyInfo.PropertyType.FullName.Contains("System.DateTime")
                && propertyInfo.PropertyType.FullName.Contains("Nullable"))
            {
                #region ...
                if (val == null)
                {
                    propertyInfo.SetValue(retobj, null);
                }
                else
                {
                    string valstr = val.ToString();
                    DateTime dt;
                    if (DateTime.TryParse(valstr, out dt))
                    {
                        propertyInfo.SetValue(retobj, dt);
                    }
                    else
                    {
                        propertyInfo.SetValue(retobj, null);
                    }
                }
                #endregion
            }
            else if (propertyInfo.PropertyType.FullName.Contains("System.DateTime"))
            {
                #region ...
                if (val == null)
                {
                    propertyInfo.SetValue(retobj, DateTime.MinValue);
                }
                else
                {
                    string valstr = val.ToString();
                    DateTime dt;
                    if (DateTime.TryParse(valstr, out dt))
                    {
                        propertyInfo.SetValue(retobj, dt);
                    }
                    else
                    {
                        propertyInfo.SetValue(retobj, DateTime.MinValue);
                    }
                }
                #endregion
            }
            else if (propertyInfo.PropertyType.FullName.Contains("System.Int32")
                && propertyInfo.PropertyType.FullName.Contains("Nullable"))
            {
                #region ...
                if (val == null)
                {
                    propertyInfo.SetValue(retobj, null);
                }
                else
                {
                    string valstr = val.ToString();
                    int it;
                    if (int.TryParse(valstr, out it))
                    {
                        propertyInfo.SetValue(retobj, it);
                    }
                    else
                    {
                        propertyInfo.SetValue(retobj, null);
                    }
                }
                #endregion
            }
            else if (propertyInfo.PropertyType.FullName.Contains("System.Int32"))
            {
                #region ...
                if (val == null)
                {
                    propertyInfo.SetValue(retobj, 0);
                }
                else
                {
                    string valstr = val.ToString();
                    int it;
                    if (int.TryParse(valstr, out it))
                    {
                        propertyInfo.SetValue(retobj, it);
                    }
                    else
                    {
                        propertyInfo.SetValue(retobj, 0);
                    }
                }
                #endregion
            }
            else if (propertyInfo.PropertyType.FullName.Contains("System.String"))
            {
                #region ...
                if (val == null)
                {
                    propertyInfo.SetValue(retobj, null);
                }
                else
                {
                    propertyInfo.SetValue(retobj, val.ToString());
                }
                #endregion
            }
            else
            {
                propertyInfo.SetValue(retobj, val);
            }
        }
    }
}
