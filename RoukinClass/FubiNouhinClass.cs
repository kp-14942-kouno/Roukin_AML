using ICSharpCode.SharpZipLib.Zip;
using Microsoft.IdentityModel.Logging;
using MyLibrary;
using MyLibrary.MyClass;
using MyLibrary.MyModules;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using static MyTemplate.RoukinClass.RoukinModules;

namespace MyTemplate.RoukinClass
{
    /// <summary>
    /// 不備納品データ作成クラス
    /// </summary>
    public class FubiNouhinClass : MyLibrary.MyLoading.Thread
    {
        public const int OP_DATA = 1 << 0;      // 不備対象データ     
        public const int OP_CALL = 1 << 1;      // コール連携用画像

        private int _operation = 0; // 操作フラグ
        private DataTable _table = new DataTable(); // データテーブル
        private string _msg = string.Empty; // メッセージ

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="table"></param>
        /// <param name="operation"></param>
        /// <param name="msg"></param>
        public FubiNouhinClass(DataTable table, int operation, string msg)
        {
            _table = table;
            _operation = operation;
            _msg = msg;
        }

        /// <summary>
        /// 排他処理
        /// </summary>
        /// <returns></returns>
        public override int MultiThreadMethod()
        {
            //try
            {
                // 開始ログ出力
                MyLogger.SetLogger($"{_msg}作成開始", MyEnum.LoggerType.Info, false);

                using (var codeDb = new MyDbData("code"))
                {
                    // 実行呼出し
                    Run(codeDb);

                    // 結果メッセージ
                    ResultMessage = $"{_msg}全：{_table.Rows.Count} 件の作成完了";
                    // 終了ログ出力
                    MyLogger.SetLogger(ResultMessage, MyEnum.LoggerType.Info, false);

                    Result = MyLibrary.MyEnum.MyResult.Ok;
                }

            }
            //catch (Exception ex)
            //{
            //    MyLogger.SetLogger(ex, MyLibrary.MyEnum.LoggerType.Error);
            //    Result = MyLibrary.MyEnum.MyResult.Error;
            //}
            Completed = true;
            return 0;
        }

        /// <summary>
        /// 実行
        /// </summary>
        /// <param name="codeDb"></param>
        /// <exception cref="Exception"></exception>
        private void Run(MyDbData codeDb)
        {
            // 出力先パスを取得
            string expPath = MyUtilityModules.AppSetting("roukin_setting", "exp_root_path");
            // 金庫事務用ディレクトリ
            string safeBoxDir = MyUtilityModules.AppSetting("roukin_setting", "safe_box_admin_dir", true);
            // 不備対象者データファイル名
            string fubiName = MyUtilityModules.AppSetting("roukin_setting", "fubi_data_name", true);
            // 不備対象イメージ管理データファイル名
            string fubiImgName = MyUtilityModules.AppSetting("roukin_setting", "fubi_img_data_name", true);
            // 不備対象イメージデータのZIPファイル名
            string fubiZipName = MyUtilityModules.AppSetting("roukin_setting", "fubi_img_zip_name", true);
            // 申請書画像パス
            string imgPath = MyUtilityModules.AppSetting("roukin_setting", "img_root_path", true);
            // 不備状画像フォルダ名
            string fubiImgDir = MyUtilityModules.AppSetting("roukin_setting", "fubi_img_dir", true);

            // 金融機関コードを重複除外して取得
            var banks = _table.AsEnumerable().Select(x => x["bpo_bank_code"].ToString()).Distinct().ToList();
            // 作成日
            var date = DateTime.Now.ToString("yyyyMMdd");

            // 金融機関コードごとに処理
            foreach (string bank in banks)
            {
                // 金融機関コードから金融機関名を取得
                var bankData = codeDb.ExecuteQuery($"select * from t_financial_code where code = '{bank}'");

                // 金融機関名が見つからない場合は例外を投げる
                if (bankData.Rows.Count == 0)
                {
                    throw new Exception($"金融機関コードが見つかりません: {bank}");
                }

                // 金融機関名を取得
                string bankName = bankData.Rows[0]["financial_name"].ToString().Trim();
                // 支店番号を取得
                string branchNo = bankData.Rows[0]["branch_number"].ToString().Trim();

                bankData.Dispose(); // データテーブルを破棄

                // 出力先パス作成
                string expDir = System.IO.Path.Combine(expPath, safeBoxDir, bank + bankName);

                // 出力先作成
                System.IO.Directory.CreateDirectory(expDir);

                var table = _table.AsEnumerable()
                    .Where(x => x["bpo_bank_code"].ToString().Trim() == bank)
                    .CopyToDataTable();

                // 不備納品データの作成
                if ((_operation & OP_DATA) != 0)
                {
                    CreateData(table, bank, bankName, branchNo, expDir, imgPath, fubiName, fubiImgName, fubiZipName, date);
                }

                // コール連携用画像作成
                if ((_operation & OP_CALL) != 0)
                {
                    // コール連携用画像の作成処理をここに追加
                }
            }
        }

        /// <summary>
        /// 不備対象データ作成
        /// </summary>
        /// <param name="table"></param>
        /// <param name="bank"></param>
        /// <param name="bankName"></param>
        /// <param name="branchNo"></param>
        /// <param name="expDir"></param>
        /// <param name="imgPath"></param>
        /// <param name="fubiName"></param>
        /// <param name="fubiImgName"></param>
        /// <param name="fubiZipName"></param>
        /// <param name="date"></param>
        private void CreateData(DataTable table, string bank, string bankName, string branchNo, string expDir, string imgPath, string fubiName, string fubiImgName, string fubiZipName, string date)
        {
            int count = 0;
            string delimiter = ","; // 区切り文字
            using (var zipStream = new FileStream(Path.Combine(expDir, fubiZipName), FileMode.Create, FileAccess.Write))
            using (var zipOutput = new ZipOutputStream(zipStream))
            using (var fubiStrm = new StreamWriter(Path.Combine(expDir, fubiName), false, MyUtilityModules.GetEncoding(MyEnum.MojiCode.Sjis)))
            using (var imgStrm = new StreamWriter(Path.Combine(expDir, fubiImgName), false, MyUtilityModules.GetEncoding(MyEnum.MojiCode.Sjis)))
            {
                // 不備対象データのヘッダーを書き込む
                fubiStrm.WriteLine("金融機関コード,支店番号,顧客番号,カナ氏名,漢字氏名,案件毎番号," +
                                   string.Join(delimiter, Enumerable.Range(1, 42)
                                   .Select(i => $"不備事由{string.Concat(i.ToString().Select(c => c >= '0' && c <= '9' ? (char)('０' + (c - '0')) : c))}")));

                // 不備対象イメージ管理データのヘッダーを書き込む
                imgStrm.WriteLine("\"金融機関コード\",\"支店番号\",\"顧客番号\",\"案件毎番号\",\"代表書類区分\",\"イメージ毎番号\"");

                // データ分繰り返す
                foreach (DataRow row in table.Rows)
                {
                    ProgressValue++;
                    count++;

                    StringBuilder fubi = new StringBuilder();
                    StringBuilder img = new StringBuilder();

                    // 案件毎番号
                    string caseNo = bank + "_" + branchNo + "_" + date + "_" + count.ToString("D4");

                    // 不備納品データの作成
                    fubi.Append(row["bpo_bank_code"].ToString().Trim() + delimiter); // 金融機関コード
                    fubi.Append(branchNo + delimiter); // 支店番号
                    fubi.Append(row["bpo_cust_no"].ToString().Trim() + delimiter); // 顧客番号
                    fubi.Append(row["bpo_kana_name"].ToString().Trim() + delimiter); // カナ氏名
                    fubi.Append(row["bpo_org_kanji"].ToString().Trim() + delimiter); // 漢字氏名
                    fubi.Append(caseNo + delimiter); // 案件毎番号

                    // 不備コード
                    var fubiCode = row["fubi_code"].ToString().Trim().Split(";");
                    // 不備コードが存在する場合は不備コードを追加
                    for (int i = 0; i < 42; i++)
                    {
                        if (i < fubiCode.Count())
                        {
                            fubi.Append(fubiCode[i].Trim() + delimiter); // 不備コード
                        }
                        else
                        {
                            fubi.Append(delimiter); // 不備コードがない場合は空欄
                        }
                    }
                    // 書出し
                    fubiStrm.WriteLine(fubi.ToString());
                    fubi.Clear(); // StringBuilderをクリア

                    // 不備対象イメージ管理データ作成
                    // 画像ファイル枚数
                    int imgCount = int.Parse(row["img_num"].ToString());

                    // 画像ファイル枚数分レコードを作成する
                    for (int x = 1; x <= imgCount; x++)
                    {
                        img.Append(SetDc(row["bpo_bank_code"].ToString().Trim()) + delimiter); // 金融機関コード
                        img.Append(SetDc(branchNo) + delimiter); // 支店番号
                        img.Append(SetDc(row["bpo_cust_no"].ToString().Trim()) + delimiter); // 顧客番号
                        img.Append(SetDc(caseNo) + delimiter); // 案件毎番号
                        img.Append(SetDc("2") + delimiter); // 代表書類区分
                        img.Append(SetDc(x.ToString("D4"))); // イメージ毎番号

                        // 書出し
                        imgStrm.WriteLine(img.ToString());
                        img.Clear(); // StringBuilderをクリア

                        // 画像ファイルをZIPへ書出し
                        // 画像ファイルのパスを作成
                        string sourceFile = Path.Combine(imgPath, row["taba_num"].ToString(), $"{row["bpo_num"].ToString()}{(x == 1 ? "" : "_" + x)}.jpg");
                        // 書出し画像ファイル名
                        string imgName = caseNo + "_" + x.ToString("D4") + ".jpg";
                        // ZIPファイルに追加
                        var entry = new ZipEntry(imgName)
                        {
                            DateTime = DateTime.Now,
                            Size = new FileInfo(sourceFile).Length
                        };
                        zipOutput.PutNextEntry(entry);
                        // 画像ファイルをZIPに書き込む
                        using var inputStrem = File.OpenRead(sourceFile);
                        inputStrem.CopyTo(zipOutput);
                        zipOutput.CloseEntry(); // ZIPエントリを閉じる
                    }
                }
                zipStream.Flush(); // ZIPストリームをフラッシュ
            }
        }
    }
}
