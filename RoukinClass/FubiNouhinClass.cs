using DocumentFormat.OpenXml.Presentation;
using ICSharpCode.SharpZipLib.Zip;
using ImageMagick;
using Microsoft.IdentityModel.Logging;
using MyLibrary;
using MyLibrary.MyClass;
using MyLibrary.MyModules;
using Mysqlx.Expr;
using MyTemplate.MyClass;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Markup;
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
        private string _expPath = string.Empty; // 出力先パス

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="table"></param>
        /// <param name="operation"></param>
        /// <param name="msg"></param>
        public FubiNouhinClass(DataTable table, int operation, string msg, string expPath)
        {
            _table = table;
            _operation = operation;
            _msg = msg;
            _expPath = expPath;
        }

        /// <summary>
        /// 排他処理
        /// </summary>
        /// <returns></returns>
        public override int MultiThreadMethod()
        {

#if DEBUG
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
#else

            try
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
            catch (Exception ex)
            {
                MyLogger.SetLogger(ex, MyLibrary.MyEnum.LoggerType.Error);
                Result = MyLibrary.MyEnum.MyResult.Error;
            }
#endif
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
            ProgressBarType = MyLibrary.MyEnum.MyProgressBarType.Percent;
            ProcessName = $"不備納品作成中..."; // 処理名設定
            ProgressMax = _table.Rows.Count;
            ProgressValue = 0;

            // 金庫事務用ディレクトリ
            string safeBoxDir = MyUtilityModules.AppSetting("roukin_setting", "safe_box_admin_dir", true);
            // 不備対象者データファイル名
            string fubiName = MyUtilityModules.AppSetting("roukin_setting", "fubi_data_name", true);
            // 不備対象イメージ管理データファイル名
            string fubiImgName = MyUtilityModules.AppSetting("roukin_setting", "fubi_img_data_name", true);
            // 不備対象イメージデータのZIPファイル名
            string fubiZipName = MyUtilityModules.AppSetting("roukin_setting", "fubi_img_zip_name", true);
            // コンバート画像パス
            string imgPath = MyUtilityModules.AppSetting("roukin_setting", "img_root_path", true);
            // 申請書画像フォルダ名
            string imgDir = MyUtilityModules.AppSetting("roukin_setting", "img_dir", true);
            // 不備状画像フォルダ名
            string fubiImgDir = MyUtilityModules.AppSetting("roukin_setting", "fubi_img_dir", true);
            // コール連携用不備画像ZIPファイル名
            string callFubiZipName = MyUtilityModules.AppSetting("roukin_setting", "fubi_call_zip_name", true);
            // コール連携用ZIPのパスワード
            string callFubiZipPass = MyUtilityModules.AppSetting("roukin_setting", "fubi_call_zip_password", true);
            // 不備状画像納品ディレクトリパス
            string fubiNouhinImgDir = MyUtilityModules.AppSetting("roukin_setting", "fubi_nouhin_img_dir");

            // 金融機関コードを重複除外して取得
            var banks = _table.AsEnumerable().Select(x => x["bpo_bank_code"].ToString()).Distinct().ToList();
            // 作成日
            var date = DateTime.Now.ToString("yyyyMMdd");

            // コール用画像のZIPファイル作成
            using (var archive = new MyArchiveWriter(Path.Combine(_expPath, callFubiZipName), true, true))
            {
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
                    bankName = bankName.Replace("労金", ""); // 金融機関名から「労金」を削除
                                                           // 支店番号を取得
                    string branchNo = bankData.Rows[0]["branch_number"].ToString().Trim();

                    bankData.Dispose(); // データテーブルを破棄

                    // 出力先パス作成
                    string expDir = System.IO.Path.Combine(_expPath, safeBoxDir, bankName);
                    // 画像パス
                    string expImgDir = System.IO.Path.Combine(expDir, fubiNouhinImgDir);

                    // 出力先作成
                    System.IO.Directory.CreateDirectory(expDir);
                    System.IO.Directory.CreateDirectory(expImgDir);

                    // 対象の金融機関コードのデータを抽出
                    var table = _table.AsEnumerable()
                        .Where(x => x["bpo_bank_code"].ToString().Trim() == bank)
                        .CopyToDataTable();

                    // 不備納品データの作成
                    if ((_operation & OP_DATA) != 0)
                    {
                        CreateData(table, bank, bankName, branchNo, expDir, expImgDir, imgPath, imgDir, fubiImgDir, fubiName, fubiImgName, fubiZipName, date, archive, callFubiZipPass);
                    }
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
        private void CreateData(DataTable table, string bank, string bankName, string branchNo, string expDir, string expImgDir, string imgPath, string imgDir, string fubiImgDir, string fubiName, string fubiImgName, string fubiZipName, string date, MyArchiveWriter archive, string callFubiZipPass)
        {
            int count = 0;
            string delimiter = ","; // 区切り文字

            // 不備対象データのファイルストリームを作成
            using (var fsFubi = new FileStream(Path.Combine(expDir, fubiName), FileMode.CreateNew, FileAccess.Write, FileShare.None))
            using (var fubiStrm = new StreamWriter(fsFubi, MyUtilityModules.GetEncoding(MyEnum.MojiCode.Sjis)))
            {
                // 不備対象データのヘッダーを書き込む
                fubiStrm.WriteLine("金融機関コード,支店番号,顧客番号,カナ氏名,漢字氏名,案件毎番号," +
                                   string.Join(delimiter, Enumerable.Range(1, 42)
                                   .Select(i => $"不備事由{string.Concat(i.ToString().Select(c => c >= '0' && c <= '9' ? (char)('０' + (c - '0')) : c))}")));

                // データ分繰り返す
                foreach (DataRow row in table.Rows)
                {
                    ProgressValue++;
                    count++;

                    StringBuilder fubi = new StringBuilder();

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

                    // 申請書 +不備状作成（PDF化）
                    var pdfStream = ConvertImagesToPdf(row, imgPath, imgDir, fubiImgDir);
                    // 書出し画像ファイル名
                    string imgName = bank + "_" + branchNo + "_" + row["bpo_cust_no"].ToString().Trim() + ".PDF";

                    using var fs = new FileStream(Path.Combine(expImgDir, imgName), FileMode.CreateNew, FileAccess.Write, FileShare.None);
                    pdfStream.CopyTo(fs); // PDFストリームをファイルに書き込む

                    // コール用ZIPへ書出し
                    archive.WriteBytes(row["bpo_num"].ToString() + ".PDF", pdfStream.ToArray(), callFubiZipPass);
                }
            }
        }

        /// <summary>
        /// 申請書と不備状の画像をPDFに変換
        /// </summary>
        /// <param name="row"></param>
        /// <param name="imgPath"></param>
        /// <param name="fubiImgDir"></param>
        /// <returns></returns>
        private static MemoryStream ConvertImagesToPdf(DataRow row, string imgPath, string imgDir, string fubiImgDir)
        {
            // 申請書の画像枚数
            int imgCount = int.Parse(row["img_num"].ToString());
            // 不備状の画像枚数
            int fubiCount = int.Parse(row["fubi_img_num"].ToString());

            List<string> sourceFiles = new List<string>();

            // 申請書の画像ファイルのパスを作成
            for (int x = 1; x <= imgCount; x++)
            {
                // 画像ファイルのパスを作成
                string sourceFile = Path.Combine(imgPath, imgDir, row["taba_num"].ToString(), $"{row["bpo_num"].ToString()}{(x == 1 ? "" : "_" + ( x - 1))}.jpg");
                sourceFiles.Add(sourceFile);
            }

            // 不備状の画像ファイルのパスを作成
            for (int x = 1; x <= fubiCount; x++)
            {
                // 不備状の画像ファイルのパスを作成
                string sourceFile = Path.Combine(imgPath, fubiImgDir, row["taba_num"].ToString(), $"{row["bpo_num"].ToString()}{x}.tiff");
                sourceFiles.Add(sourceFile);
            }

            // MemoryStreamを作成
            var memoryStream = new MemoryStream();

            // 画像をPDFに変換
            using (var images = new MagickImageCollection())
            {
                foreach(var file in sourceFiles)
                {
                    var img = new MagickImage(file)
                    {
                        Format = MagickFormat.Pdf
                    };
                    images.Add(img);
                }
                images.Write(memoryStream, MagickFormat.Pdf);
            }
            memoryStream.Position = 0; // ストリームの位置を先頭に戻す
            return memoryStream;
        }
    }
}
