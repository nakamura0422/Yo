using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Yo.Models;

namespace Yo.Controllers
{
    /// <summary>
    /// Npoiを使ってExcelファイルを作成しますよ
    /// </summary>
    public class CreateExcel : Controller
    {
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public IWorkbook WriteWorkBook()
        {
            var book = WorkbookFactory.Create("ほげほげ.xlsx");
            //シート名からシート取得
            var sheet = book.GetSheet("newSheet");
            //セルに設定
            WriteCell(sheet, 0, 0, "0-0");
            WriteCell(sheet, 1, 1, "1-1");
            WriteCell(sheet, 0, 3, 100);
            WriteCell(sheet, 0, 4, DateTime.Today);
            //日付表示するために書式変更
            var style = book.CreateCellStyle();
            style.DataFormat = book.CreateDataFormat().GetFormat("yyyy/mm/dd");
            WriteStyle(sheet, 0, 4, style);

            return book;
        }

        /// <summary>
        /// セル設定(文字列用)
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columnIndex"></param>
        /// <param name="rowIndex"></param>
        /// <param name="value"></param>
        public static void WriteCell(ISheet sheet, int columnIndex, int rowIndex, string value)
        {
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellValue(value);
        }

        /// <summary>
        /// セル設定(数値用)
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columnIndex"></param>
        /// <param name="rowIndex"></param>
        /// <param name="value"></param>
        private void WriteCell(ISheet sheet, int columnIndex, int rowIndex, double value)
        {
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellValue(value);
        }

        /// <summary>
        /// セル設定(日付用)
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columnIndex"></param>
        /// <param name="rowIndex"></param>
        /// <param name="value"></param>
        private void WriteCell(ISheet sheet, int columnIndex, int rowIndex, DateTime value)
        {
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellValue(value);
        }

        /// <summary>
        /// 書式変更
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columnIndex"></param>
        /// <param name="rowIndex"></param>
        /// <param name="style"></param>
        private void WriteStyle(ISheet sheet, int columnIndex, int rowIndex, ICellStyle style)
        {
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.CellStyle = style;
        }


        /// <summary>
        /// ExcelファイルからBookを作成
        /// </summary>
        private IWorkbook CreateNewBook(string filePath)
        {
            IWorkbook book;
            var extension = Path.GetExtension(filePath);

            // HSSF => Microsoft Excel(xls形式)(excel 97-2003)
            // XSSF => Office Open XML Workbook形式(xlsx形式)(excel 2007以降)
            if (extension == ".xls")
                book = new HSSFWorkbook();
            else if (extension == ".xlsx")
                book = new XSSFWorkbook();
            else
                throw new ApplicationException("CreateNewBook: invalid extension");
            return book;
        }
    }
}
