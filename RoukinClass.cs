using DocumentFormat.OpenXml.Presentation;
using MyLibrary;
using MyLibrary.MyClass;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;

namespace MyTemplate
{
    public class FubiCheck : MyLibrary.MyLoading.Thread
    {
        private DataTable _dataTable;
        private DataTable _contoryCode;
        private DataTable _bussinessCode;


        public FubiCheck(DataTable dataTable)
        {
            _dataTable = dataTable;
        }

        public override int MultiThreadMethod()
        {
            try
            {
                // 各コードをテーブルで取得
                GetCodeData();
                // 実行
                Run();
            }
            catch (Exception ex)
            {
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error);
                Result = MyEnum.MyResult.Error;
            }
            Completed = true;
            return 0;
        }

        private void Run()
        {
            // ビートルDB
            using (var btl = new MyDbData("beetle"))
            { 
                foreach (DataRow row in _dataTable.Rows)
                {
                    // ビートルデータを取得
                    var prm = new Dictionary<string, object>
                    {
                        { "@qr_code", row["qr_code"].ToString() }
                    };
                    var btlData = btl.ExecuteQuery($"select * from T_DOC_INFO where [BARCODE_NO] = @qr_code;", prm);

                    // ビートルデータが存在しない場合はエラーで中断
                    if (btlData.Rows.Count == 0)
                    {
                        throw new Exception($"ビートルデータが見つかりません。QRコード: {row["qr_code"]}");
                    }

                    // 不備コード用
                    StringBuilder fubiCode = new StringBuilder();
                }
            }
        }

        /// <summary>
        /// 不備チェック　01　団体名
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_01(DataRow row, DataRow beatle, StringBuilder fubiCode)
        {
            string grpFlg = row["grpname_flg"].ToString().Trim();

            // 変更なしはOK
            if (grpFlg == "0") return;
            // 変更有で団体名漢字に値が入っている場合はOK
            if (grpFlg == "1" && !string.IsNullOrEmpty(row["grpname_kan"].ToString())) return;
            // 変更有無未選択で団体名漢字に値が入っている場合はOK
            if (string.IsNullOrEmpty(grpFlg) && !string.IsNullOrEmpty(row["grpname_kan"].ToString())) return;

            // それ以外は不備コードをセット
            fubiCode.Append("01");
        }

        /// <summary>
        /// 不備チェック　02　団体種別
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_02(DataRow row, DataRow beatle, StringBuilder fubiCode)
        {
            string grpType = row["grpType_1"].ToString().Trim() + row["grpType_2"].ToString().Trim();

            // 団体種類が選択ありで選択が1つならOK
            if (!string.IsNullOrEmpty(grpType) || grpType.Length == 1) return;

            // それ以外は不備コードをセット
            fubiCode.Append("02");
        }

        /// <summary>
        /// 不備チェック　03　事業内容 
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_03(DataRow row, DataRow beatle, StringBuilder fubiCode)
        {
            string bizCont1 = row["bixcont_1"].ToString().Trim();
            string bizCont2 = row["bixcont_2"].ToString().Trim();

            // 事業内容が未選択・複数選択以外はOK
            if (!(string.IsNullOrEmpty(bizCont1) && string.IsNullOrEmpty(bizCont2)) &&
                !(string.IsNullOrEmpty(bizCont1) == false && string.IsNullOrEmpty(bizCont2) == false)) return;

            // それ以外は不備コードをセット
            fubiCode.Append("03");
        }

        /// <summary>
        /// 不備チェック　04　設立年月日
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_04(DataRow row, DataRow beatle, StringBuilder fubiCode)
        {
            string estdateFlg = row["estdate_flg"].ToString().Trim();
            string estdate = row["est_date"].ToString().Trim();

            // 変更なしで設立年月日が空はOK
            if (estdateFlg == "0" && string.IsNullOrEmpty(estdate)) return;

            // 設立年月日（yyyymmdd）をyyyy/MM/dd形式に変換
            DateTime? formattedDate = MyLibrary.MyModules.MyUtilityModules.ParseDateString(estdate, "yyyyMMdd", MyEnum.CalenderType.Japanese);

            // 設立年月日がyyyy/MM/dd形式に変換できた場合はOK
            if (!string.IsNullOrEmpty(formattedDate.ToString()))
            {
                // 設立年月日が未来日でなければOK
                if (formattedDate <= DateTime.Now) return;
            }

            // それ以外は不備コードをセット
            fubiCode.Append("04");
        }

        /// <summary>
        /// 不備チェック　05　郵便番号
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_05(DataRow row, DataRow beatle, StringBuilder fubiCode)
        {
            string addrChgFlg = row["addrchg_flg"].ToString().Trim();
            string zipCode = row["new_zip"].ToString().Trim();

            // 住所変更なしで郵便番号が空はOK
            if (addrChgFlg=="0" && string.IsNullOrEmpty(zipCode)) return;

            // zipCodeが7桁の数字はOK
            if (System.Text.RegularExpressions.Regex.IsMatch(zipCode, @"^\d{7}$")) return;

            // それ以外は不備コードをセット
            fubiCode.Append("05");
        }

        /// <summary>
        /// 不備チェック　06　所在地
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_06(DataRow row, DataRow beatle, StringBuilder fubiCode)
        {
            string addrChgFlg = row["addrchg_flg"].ToString().Trim();
            string addr = row["new_pref"].ToString().Trim() +
                        row["new_city"].ToString().Trim() +
                        row["new_addrno"].ToString().Trim() +
                        row["new_bldg"].ToString().Trim();

            // 住所変更なしで住所が空はOK
            if (addrChgFlg == "0" && string.IsNullOrEmpty(addr)) return;

            // 住所があればOK
            if (string.IsNullOrEmpty(addr)) return;

            // それ以外は不備コードをセット
            fubiCode.Append("06");
        }

        /// <summary>
        /// 不備チェック　07　本店所在国
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_07(DataRow row, DataRow beatle, StringBuilder fubiCode)
        {

        }



        /// <summary>
        /// 各コード変換用テーブルを取得
        /// </summary>
        private void GetCodeData()
        {
            using (var db = new MyDbData("code"))
            {
                // 国コード
                _contoryCode = db.ExecuteQuery("select * from t_country_code order by  code;");
                // 業種コード
                _bussinessCode = db.ExecuteQuery("select * from t_bussiness_code order by code;");
            }
        }
    }
}
