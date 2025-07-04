using DocumentFormat.OpenXml.Presentation;
using MyLibrary;
using MyLibrary.MyClass;
using MyLibrary.MyModules;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Markup;

namespace MyTemplate
{

    /// <summary>
    /// WEBCASデータ変換出力クラス
    /// </summary>
    public class EpxWebcasData : MyLibrary.MyLoading.Thread
    {
        DataTable _table;
        string _fileDir = string.Empty; // 出力先ディレクトリ

        public EpxWebcasData(DataTable table, string fileDir)
        {
            _fileDir = fileDir;
            _table = table;
        }

        /// <summary>
        /// 排他処理
        /// </summary>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        public override int MultiThreadMethod()
        {
            try
            {
                Run();
                Result = MyEnum.MyResult.Ok;
            }
            catch(Exception ex)
            {
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error);
                Result = MyEnum.MyResult.Error;
            }
            Completed = true;
            return 0;
        }

        /// <summary>
        /// 実行
        /// </summary>
        private void Run()
        {
            string fileName = MyUtilityModules.AppSetting("roukin_setting", "exp_webcas_file_name", true, _table.Rows.Count);
            string delimiter = MyUtilityModules.AppSetting("roukin_setting", "exp_webcas_file_delimiter");
            string mojiCode = MyUtilityModules.AppSetting("roukin_setting", "exp_webcas_file_mojicode");

            ProgressValue = 0;
            ProgressMax = _table.Rows.Count;
            ProcessName = "WEBCASデータ変換出力";

            using (System.IO.StreamWriter writer = new System.IO.StreamWriter(System.IO.Path.Combine(_fileDir, fileName), false, MyUtilityModules.GetEncoding(mojiCode)))
            {
                StringBuilder builder = new StringBuilder();

                foreach (DataRow row in _table.Rows)
                {
                    ProgressValue++;

                    builder.Clear();

                    builder.Append(row["ctrl_num"].ToString().Trim() + delimiter);
                    builder.Append(row["answer_date"].ToString().Trim() + delimiter);
                    builder.Append(row["namechg_flg"].ToString().Trim() + delimiter);
                    builder.Append(row["lname_kanji"].ToString().Trim() + delimiter);
                    builder.Append(row["fname_kanji"].ToString().Trim() + delimiter);
                    builder.Append(row["lname_kana"].ToString().Trim() + delimiter);
                    builder.Append(row["fname_kana"].ToString().Trim() + delimiter);
                    builder.Append(row["addrchg_flg"].ToString().Trim() + delimiter);
                    builder.Append(row["zipcode"].ToString().Trim() + delimiter);
                    builder.Append(row["pref"].ToString().Trim() + delimiter);
                    builder.Append(row["city"].ToString().Trim() + delimiter);
                    builder.Append(row["addr_num"].ToString().Trim() + delimiter);
                    builder.Append(row["bldg_name"].ToString().Trim() + delimiter);
                    builder.Append(row["pref_kana"].ToString().Trim() + delimiter);
                    builder.Append(row["city_kana"].ToString().Trim() + delimiter);
                    builder.Append(row["addrnum_kan"].ToString().Trim() + delimiter);
                    builder.Append(row["bldgkana_nm"].ToString().Trim() + delimiter);
                    builder.Append(row["tel_fixed"].ToString().Trim() + delimiter);
                    builder.Append(row["tel_mobile"].ToString().Trim() + delimiter);
                    builder.Append(row["email"].ToString().Trim() + delimiter);
                    builder.Append(row["nationchg_flg"].ToString().Trim() + delimiter);
                    builder.Append(row["nation"].ToString().Trim() + delimiter);
                    builder.Append(row["region"].ToString().Trim() + delimiter);
                    builder.Append(row["nation_asia"].ToString().Trim() + delimiter);
                    builder.Append(row["nation_mide"].ToString().Trim() + delimiter);
                    builder.Append(row["nation_weur"].ToString().Trim() + delimiter);
                    builder.Append(row["nation_eeur"].ToString().Trim() + delimiter);
                    builder.Append(row["nation_nam"].ToString().Trim() + delimiter);
                    builder.Append(row["nation_cari"].ToString().Trim() + delimiter);
                    builder.Append(row["nation_latm"].ToString().Trim() + delimiter);
                    builder.Append(row["nation_afrc"].ToString().Trim() + delimiter);
                    builder.Append(row["nation_oce"].ToString().Trim() + delimiter);
                    builder.Append(row["nation_othr"].ToString().Trim() + delimiter);
                    builder.Append(row["name_alpha"].ToString().Trim() + delimiter);
                    builder.Append(row["visa"].ToString().Trim() + delimiter);
                    builder.Append(row["visa_limit"].ToString().Trim() + delimiter);
                    builder.Append(row["pep_flag"].ToString().Trim() + delimiter);
                    builder.Append(row["job"].ToString().Trim() + delimiter);
                    builder.Append(row["job_other"].ToString().Trim() + delimiter);
                    builder.Append(row["work_or_sch"].ToString().Trim() + delimiter);
                    builder.Append(row["worksch_kan"].ToString().Trim() + delimiter);
                    builder.Append(row["work_tel"].ToString().Trim() + delimiter);
                    builder.Append(row["industry"].ToString().Trim() + delimiter);
                    builder.Append(row["indus_othr"].ToString().Trim() + delimiter);
                    builder.Append(row["sidejob_flg"].ToString().Trim() + delimiter);
                    builder.Append(row["sidejob_typ"].ToString().Trim() + delimiter);
                    builder.Append(row["sideothr_tx"].ToString().Trim() + delimiter);
                    builder.Append(row["tx_purpose1"].ToString().Trim() + delimiter);
                    builder.Append(row["tx_purpose2"].ToString().Trim() + delimiter);
                    builder.Append(row["tx_purpose3"].ToString().Trim() + delimiter);
                    builder.Append(row["tx_purpose4"].ToString().Trim() + delimiter);
                    builder.Append(row["tx_purpose5"].ToString().Trim() + delimiter);
                    builder.Append(row["tx_purpose6"].ToString().Trim() + delimiter);
                    builder.Append(row["tx_othr_txt"].ToString().Trim() + delimiter);
                    builder.Append(row["tx_type"].ToString().Trim() + delimiter);
                    builder.Append(row["tx_freq"].ToString().Trim() + delimiter);
                    builder.Append(row["tx_amt_once"].ToString().Trim() + delimiter);
                    builder.Append(row["tx_over2m_f"].ToString().Trim() + delimiter);
                    builder.Append(row["tx_over2mfrq"].ToString().Trim() + delimiter);
                    builder.Append(row["tx_over2mam"].ToString().Trim() + delimiter);
                    builder.Append(row["src_salary"].ToString().Trim() + delimiter);
                    builder.Append(row["src_pension"].ToString().Trim() + delimiter);
                    builder.Append(row["src_insur"].ToString().Trim() + delimiter);
                    builder.Append(row["src_exec"].ToString().Trim() + delimiter);
                    builder.Append(row["src_bizin"].ToString().Trim() + delimiter);
                    builder.Append(row["src_inheri"].ToString().Trim() + delimiter);
                    builder.Append(row["src_invincm"].ToString().Trim() + delimiter);
                    builder.Append(row["src_saving"].ToString().Trim() + delimiter);
                    builder.Append(row["src_otherbk"].ToString().Trim() + delimiter);
                    builder.Append(row["src_home"].ToString().Trim() + delimiter);
                    builder.Append(row["src_loan"].ToString().Trim() + delimiter);
                    builder.Append(row["src_other"].ToString().Trim() + delimiter);
                    builder.Append(row["src_oth_txt"].ToString().Trim() + delimiter);

                    // 書き込み
                    writer.WriteLine(builder);
                }

                ResultMessage = $"WEBCASデータ：{ProgressValue}件の変換出力が完了しました。";
            }
        }
    }

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
