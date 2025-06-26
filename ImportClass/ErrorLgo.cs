using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyTemplate.ImportClass
{
    /// <summary>
    /// エラーログクラス
    /// </summary>
    internal class ErrorLog: IDisposable
    {
        public DataTable _errorLogData;

        /// <summary>
        /// 開放
        /// </summary>
        public void Dispose()
        {
            _errorLogData.Dispose();
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public ErrorLog()
        {
            _errorLogData = new DataTable();
            _errorLogData.Columns.Add("@porcess", typeof(string));
            _errorLogData.Columns.Add("@data_name", typeof(string));
            _errorLogData.Columns.Add("@record_num", typeof(int));
            _errorLogData.Columns.Add("@message", typeof(string));
            _errorLogData.Columns.Add("@error_code", typeof(byte));
        }

        /// <summary>
        /// エラーログデータテーブル
        /// </summary>
        public DataView DefaultView { get { return _errorLogData.DefaultView; } }

        /// <summary>
        /// エラーログ件数
        /// </summary>
        public int Count { get { return _errorLogData.Rows.Count; } }

        /// <summary>
        /// エラーログフィールド
        /// </summary>
        /// <returns></returns>
        public List<string> ErrorLogField()
        {
            return new List<string> { "@porcess", "@data_name", "@record_num", "@message", "@error_code" };
        }

        /// <summary>
        /// エラーログ表題
        /// </summary>
        /// <returns></returns>
        public List<string> ErrorLogCaption()
        {
            return new List<string> { "処理名", "データ名", "行番号", "メッセージ", "エラーコード" };
        }

        /// <summary>
        /// エラーログ追加
        /// </summary>
        /// <param name="processName"></param>
        /// <param name="dataName"></param>
        /// <param name="recordNum"></param>
        /// <param name="message"></param>
        /// <param name="errorCode"></param>
        public void AddErrorLog(string processName, string dataName, int recordNum, string message, byte errorCode)
        {
            DataRow row = _errorLogData.NewRow();
            row["@porcess"] = processName;
            row["@data_name"] = dataName;
            row["@record_num"] = recordNum;
            row["@message"] = message;
            row["@error_code"] = errorCode;
            _errorLogData.Rows.Add(row);
        }
    }
}
