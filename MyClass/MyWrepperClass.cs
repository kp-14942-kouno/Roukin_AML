using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyTemplate.Class
{
    /// <summary>
    /// DataRowをラップするクラス
    /// </summary>
    public class DataRowWrapper
    {
        private readonly DataRow _row;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="row"></param>
        public DataRowWrapper(DataRow row)
        {
            _row = row;
        }

        /// <summary>
        /// 項目値を取得する
        /// </summary>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public object GetValue(string columnName)
        {
            return _row[columnName];
        }
    }

    /// <summary>
    /// DbDataReaderをラップするクラス
    /// </summary>
    public class DataReaderWrapper
    {
        private readonly DbDataReader _reader;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="reader"></param>
        public DataReaderWrapper(DbDataReader reader)
        {
            _reader = reader;
        }

        /// <summary>
        /// 項目値を取得する
        /// </summary>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public object GetValue(string columnName)
        {
            return _reader[columnName];
        }
    }
}
