using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace MyTemplate.Class
{
    /// <summary>
    /// MyStreamExcelクラス
    /// </summary>
    public class MyStreamExcel : IDisposable
    {
        XLWorkbook? _workBook;

        public XLWorkbook WorkBook
        {
            get { return _workBook; }
        }

        /// <summary>
        //　開放　
        /// </summary>
        public void Dispose()
        {
            _workBook?.Dispose();
            _workBook = null;
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="filePath"></param>
        /// <exception cref="Exception"></exception>
        public MyStreamExcel(string filePath)
        {
            try
            {
                _workBook = new XLWorkbook(filePath);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Error Open File.", ex);
            }
        }

        /// <summary>
        /// 指定行の指定列のデータを配列で取得
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="lineNum"></param>
        /// <param name="ranges"></param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException"></exception>
        public string[] ReadLine(string sheetName, int lineNum, string[] ranges)
        {
            try
            {
                var values = new List<string>();
                var workSheet = _workBook?.Worksheet(sheetName);

                foreach (string range in ranges)
                {
                    // rangeの値が固定セル「例:AA12」の場合
                    if (Regex.IsMatch(range, @"^[A-Za-z]+[0-9]+$"))
                    {
                        var cell = workSheet.Cell(range);
                        values.Add(cell.GetFormattedString());
                    }
                    else
                    {
                        var cell = workSheet.Cell(lineNum, range);
                        values.Add(cell.GetFormattedString());
                    }
                }
                return values.ToArray();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Error Read Line.", ex);
            }
        }

        /// <summary>
        /// 指定のSheetとRangeから値を取得（Range, RangeRow）
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="range"></param>
        /// <param name="rangeRow"></param>
        /// <returns></returns>
        public string Range(string sheetName, string range, int rangeRow)
        {
            return Range(sheetName, range + rangeRow);
        }

        /// <summary>
        /// 指定のSheetとRangeから値を取得（Range + RangeRow）
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="range"></param>
        /// <returns></returns>
        public string Range(string sheetName, string range)
        {
            var workSheet = _workBook?.Worksheet(sheetName);
            return workSheet.Cell(range).GetFormattedString();
        }

        /// <summary>
        /// データレコード最終行番号を取得
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public int LastRowNumber(string sheetName)
        {
            var lastRowUsed = _workBook?.Worksheet(sheetName).LastRowUsed();
            return lastRowUsed != null ? lastRowUsed.RowNumber() : 0;
        }

        /// <summary>
        /// Sheet一覧取得
        /// </summary>
        /// <returns></returns>
        public List<string> GetSheetLists()
        {
            List<string> sheeetLists = new List<string>();

            foreach (IXLWorksheet sheet in _workBook.Worksheets)
            {
                sheeetLists.Add(sheet.Name);
            }

            return sheeetLists;
        }

        /// <summary>
        /// シート名の存在確認
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="IsSelect"></param>
        /// <returns></returns>
        public string IsSheet(string sheetName, bool IsSelect = false)
        {
            List<string> sheetLists = GetSheetLists();

            // シート名が存在する場合
            if (sheetLists.Contains(sheetName))
            {
                return sheetName;
            }
            // シート名が存在しない場合
            else
            {
                // シート選択ダイアログを表示
                if (IsSelect)
                {
                    Forms.ExcelSheetList sheetList = new Forms.ExcelSheetList(this);
                    sheetList.ShowDialog();
                    return sheetList.GetSheetName();
                }
            }
            return string.Empty;
        }
    }
}
