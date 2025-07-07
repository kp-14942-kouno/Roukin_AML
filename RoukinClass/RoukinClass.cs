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


}
