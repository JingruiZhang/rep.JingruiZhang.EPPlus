using OfficeOpenXml;
using System;

namespace JingruiZhang.EPPlus
{
    /// <summary>
    /// ExcelRange 相关扩展方法
    /// </summary>
    public static class ExcelRangeExtention
    {
        /// <summary>
        /// 为某一个 ExcelRange 指定含有多个字符串的字符串序列数据验证
        /// </summary>
        /// <param name="excelRange">worksheet.Cells索引出来的对象</param>
        /// <param name="showErrorMessage">当数据验证未通过时是否提示</param>
        /// <param name="validStrings">指定的可用的字符串集合</param>
        public static void AddListDataValidation(this ExcelRange excelRange, bool showErrorMessage, params string[] validStrings)
        {
            if (excelRange == null)
            {
                throw new NullReferenceException("excelRange 为空");
            }
            if (validStrings == null)
            {
                throw new ArgumentNullException("validStrings");
            }
            if (validStrings.Length == 0)
            {
                throw new Exception("数组 validStrings 不包含任何元素");
            }
            var datavalidation = excelRange.DataValidation.AddListDataValidation();
            for (int i = 0; i < validStrings.Length; i++)
            {
                datavalidation.Formula.Values.Add(validStrings[i]);
            }
            datavalidation.ShowErrorMessage = showErrorMessage;
        }
    }
}
