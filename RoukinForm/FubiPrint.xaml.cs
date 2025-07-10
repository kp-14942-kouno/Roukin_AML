using DocumentFormat.OpenXml.Presentation;
using MyLibrary;
using MyLibrary.MyClass;
using MyLibrary.MyModules;
using MyTemplate.ImportClass;
using MyTemplate.Report;
using MyTemplate.Report.Helpers;
using MyTemplate.Report.ViewModels;
using MyTemplate.Report.Views;
using MyTemplate.RoukinClass;
using Org.BouncyCastle.Crypto;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Printing;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using YamlDotNet.Serialization.BufferedDeserialization;

namespace MyTemplate.RoukinForm
{
    /// <summary>
    /// FubiPrint.xaml の相互作用ロジック
    /// </summary>
    public partial class FubiPrint : Window, IDisposable
    {
        // 不備状データ用テーブル
        private DataTable _table = new();
        // 不備文言用辞書
        Dictionary<string, string> _defectDic = new();

        /// <summary>
        /// リソースの開放
        /// </summary>
        public void Dispose()
        {
            // リソースの解放
            _table?.Dispose();
            _defectDic?.Clear();
        }

        /// <summary>
        /// デストラクタ
        /// </summary>
        ~FubiPrint()
        {
            // Disposeメソッドを呼び出す
            Dispose();
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public FubiPrint()
        {
            InitializeComponent();
            Owner = Application.Current.MainWindow;
            Modules.SetPrinterList(cmb_PrinterList);
            SetCount();
        }

        /// <summary>
        /// 件数表示
        /// </summary>
        private void SetCount()
        {
            tb_fubiCount.Text = $"{_table.Rows.Count.ToString()}件";
        }

        /// <summary>
        /// 不備文言取得
        /// </summary>
        /// <returns></returns>
        private bool GetDefectDic()
        {
            using (MyLibrary.MyLoading.Dialog dlg = new MyLibrary.MyLoading.Dialog())
            {
                var dic = new GetDefectDic();
                dlg.ThreadClass(dic);
                dlg.ShowDialog();

                if(dic.Result != MyLibrary.MyEnum.MyResult.Ok)
                {
                    return false;
                }
                _defectDic = dic.DefectDic;
            }
            return true;
        }

        /// <summary>
        /// 不備状データ読込みボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_InpFubiData_Click(object sender, RoutedEventArgs e)
        {
            // ファイル読込みプロパティ
            using (FileLoadProperties file = new FileLoadProperties())
            {
                // ファイル読込み設定
                if (!FileLoadClass.GetFileLoadSetting(4, file)) return;
                // ファイル読込み
                if (FileLoadClass.FileLoad(this, file) != MyLibrary.MyEnum.MyResult.Ok) return;

                _table = new DataTable();
                _table = file.LoadData;

                // テーブルに選択項目を追加
                var isSelectedColumn = new DataColumn("IsSelected", typeof(bool))
                {
                    DefaultValue = false
                };
                _table.Columns.Add(isSelectedColumn);

                _table = _table.AsEnumerable().OrderBy(x => x["taba_num"]).ThenBy(x => int.Parse(x["taba_count"].ToString())).CopyToDataTable();
                dg_FubiList.ItemsSource = _table.DefaultView;

                SetCount();
            }
        }

        /// <summary>
        /// Loadedイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // 不備文言を取得
            if (!GetDefectDic()) Close();
        }

        /// <summary>
        /// 閉じるボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Closedイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Closed(object sender, EventArgs e)
        {
            Dispose();
        }

        /// <summary>
        /// 全選択ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_AllSelect_Click(object sender, RoutedEventArgs e)
        {
            // 全行のIsSelected列をtrueに設定
            _table.AsEnumerable().ToList().ForEach(row => row["IsSelected"] = true);

            // DataGridを更新
            dg_FubiList.Items.Refresh();
        }

        /// <summary>
        /// 全解除ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_AllUnSelect_Click(object sender, RoutedEventArgs e)
        {
            // 全行のIsSelected列をtrueに設定
            _table.AsEnumerable().ToList().ForEach(row => row["IsSelected"] = false);

            // DataGridを更新
            dg_FubiList.Items.Refresh();
        }

        /// <summary>
        /// 不備状プレビューボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_Preview_Click(object sender, RoutedEventArgs e)
        {
            // 選択された行を取得
            var selectedRows = SelectedItem();
            // 選択された行がない場合は処理を終了
            if (selectedRows.Rows.Count == 0) return;

            // プレビューで選択できるのは1つのみ
            if (selectedRows.Rows.Count > 1)
            {
                MyMessageBox.Show("プレビューで選択できるのは１つのみです。");
                return;
            }
            // 選択された行の不備状のFixedDocumentを作成
            var document = FubiHelper.CreateFixedDocument(selectedRows.Rows[0], _defectDic);
            // FixedDocumentをReportPreivewに渡して表示
            var form = new ReportPreivew(document);
            form.ShowDialog();

            // 開放
            document = null;
        }

        /// <summary>
        /// DataGridの選択項目を取得
        /// </summary>
        /// <returns></returns>
        private DataTable SelectedItem()
        {
            var selectedRows = _table.AsEnumerable()
                .Where(row => row.Field<bool>("IsSelected"));

            if(!selectedRows.Any())
            {
                return new DataTable();
            }
            return selectedRows.CopyToDataTable();
        }

        /// <summary>
        /// 不備状印刷・画像作成ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_Print_Click(object sender, RoutedEventArgs e)
        {
            int oparation = 0;

            oparation |= chk_FubiPrint.IsChecked == true ? FubiPrintDocument.OP_PRINT : 0;
            oparation |= chk_FubiImage.IsChecked == true ? FubiPrintDocument.OP_IMAGE : 0;
            oparation |= chk_ExpData.IsChecked == true ? FubiPrintDocument.OP_BEETLE : 0;
            oparation |= chk_HikinukiPrit.IsChecked == true ? FubiPrintDocument.OP_HIKINUKI : 0;
            oparation |= chk_ExpMaching.IsChecked == true ? FubiPrintDocument.OP_MACHING : 0;

            // チェックが無ければ中断
            if (oparation == 0) return;

            ExpPrintAndData(oparation);
        }


        /// <summary>
        /// 印刷・画像作成・ビートルデータ作成を一括で行うメソッド
        /// </summary>
        /// <param name="operation"></param>
        private void ExpPrintAndData(int operation)
        {
            var print = string.Empty;
            var image = string.Empty;
            var beetle = string.Empty;

            if((operation & FubiPrintDocument.OP_PRINT) != 0)
            {
                print = "・印刷\r\n";
            }
            if ((operation & FubiPrintDocument.OP_IMAGE) != 0)
            {
                image = "・画像作成\r\n";
            }
            if ((operation & FubiPrintDocument.OP_BEETLE) != 0)
            {
                beetle = "・ビートルデータ作成\r\n";
            }
            if((operation & FubiPrintDocument.OP_HIKINUKI) != 0)
            {
                print += "・引抜きリスト\r\n";
            }
            if((operation & FubiPrintDocument.OP_MACHING) != 0)
            {
                image += "・マッチングリスト\r\n";
            }

            var message = $"{print}{image}{beetle}";

            // 選択された行を取得
            var selectedRows = SelectedItem();
            // 選択された行がない場合は処理を終了
            if (selectedRows.Rows.Count == 0) return;
            // 確認メッセージ
            if (MyMessageBox.Show($"選択された不備状の{message}を開始します。", "確認", MyEnum.MessageBoxButtons.YesNo, MyEnum.MessageBoxIcon.Info) != MyEnum.MessageBoxResult.Yes) return;

            // 選択されたプリンターを取得
            if (cmb_PrinterList.SelectedItem is not PrintQueue printer)
            {
                MyMessageBox.Show("プリンターが選択されていません。");
                return;
            }
            // プログレスダイアログ準備
            using (MyLibrary.MyLoading.Dialog dlg = new MyLibrary.MyLoading.Dialog(this))
            {
                // 不備状印刷・画像作成スレッドを実行
                var thread = new FubiPrintDocument(selectedRows, _defectDic, printer, operation, message);
                dlg.ThreadClass(thread);
                dlg.ShowDialog();
                if (thread.Result != MyLibrary.MyEnum.MyResult.Ok)
                {
                    MyMessageBox.Show($"{message}に失敗しました。");
                }
                else
                {
                    MyMessageBox.Show($"{message}が完了しました。");
                }
            }
        }
    }

    /// <summary>
    /// 不備文言データを取得するクラス
    /// </summary>
    public class GetDefectDic : MyLibrary.MyLoading.Thread
    {
        // 不備文言用辞書
        public Dictionary<string, string> DefectDic { get; private set; } = new Dictionary<string, string>();

        /// <summary>
        /// 排他処理
        /// </summary>
        /// <returns></returns>
        public override int MultiThreadMethod()
        {
            ProgressBarType = MyEnum.MyProgressBarType.None;
            ProcessName = "不備文言データ取得中";

            try
            {
                // 不備文言データを取得
                using (MyDbData db = new MyDbData("code"))
                {
                    using (DbDataReader reader = db.ExecuteReader("select * from t_fubi_code order by fubi_code"))
                    {
                        while (reader.Read())
                        {
                            DefectDic.Add(reader["fubi_code"].ToString(), reader["fubi_caption"].ToString());
                        }
                    }
                }
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
    }
}
