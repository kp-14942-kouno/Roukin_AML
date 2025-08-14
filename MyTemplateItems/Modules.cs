using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Win32;
using MyLibrary;
using MyLibrary.MyClass;
using MyLibrary.MyControls;
using Org.BouncyCastle.Crmf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Printing;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Xps;
using YamlDotNet.RepresentationModel;

namespace MyTemplate
{
    internal static class Modules
    {
        /// <summary>
        /// JavaScript実行
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="methodName"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        /// <exception cref="FileNotFoundException"></exception>
        internal static object ExcuteJavaScript(string filePath, string methodName, params object[] parameters)
        {
            if (File.Exists(filePath) == false)
            {
                throw new FileNotFoundException("File not found", filePath);
            }
            string script = File.ReadAllText(filePath, Encoding.UTF8);
            var engine = new Jint.Engine();
            engine.Execute(script);
            return engine.Invoke(methodName, parameters).ToObject();
        }

        /// <summary>
        /// フォルダ選択ダイアログ
        /// </summary>
        /// <param name="initialPath"></param>
        /// <returns></returns>
        public static string MyFolderDialog(string? initialPath = null, bool isInitialRetern = true)
        {
            // 初期パスが存在すればそのまま返す
            if(isInitialRetern && System.IO.Path.Exists(initialPath) == true) return initialPath;

            // OpenFolderDialogのインスタンスを作成
            var dlg = new OpenFolderDialog()
            {
                Title = "フォルダ選択",
                InitialDirectory = initialPath,
                ShowHiddenItems=false
            };

            // ダイアログを表示し、ユーザーがフォルダを選択したかどうかを確認
            if (dlg.ShowDialog() != true)
            {
                // ユーザーがキャンセルした場合は空文字列を返す
                return string.Empty;
            }

            // 選択されたフォルダのフルパスを返す
            return Path.GetFullPath(dlg.FolderName);
        }

        /// <summary>
        /// ファイル選択ダイアログ
        /// </summary>
        /// <param name="initialPath"></param>
        /// <param name="extensions"></param>
        /// <returns></returns>
        public static string MyFileDialog(string? initialPath = null, string[] extensions = null)
        {
            OpenFileDialog dlg = new OpenFileDialog
            {
                Title = "ファイル選択",
                InitialDirectory = initialPath,
                Filter = extensions != null ? string.Join("|", extensions.Select(ext => $"{ext.ToUpper()}|*.{ext.ToLower()}")) + "|All Files|*.*" : "All Files|*.*",
                Multiselect = false,
            };

            // ファイル選択キャンセルは空文字列を返す
            if (dlg.ShowDialog() != true)
            {
                return string.Empty;
            }
            // 選択されたファイルのフルパスを返す
            return dlg.FileName.ToString();
        }

        /// <summary>
        /// コンボボックスにテーブル一覧をセット
        /// </summary>
        /// <param name="combobox"></param>
        public static void SetTableList(ComboBox combobox)
        {
            combobox.DisplayMemberPath = "table_caption";
            combobox.SelectedValuePath = "table_id";

            using MyDbData db = new MyDbData("setting");
            var table = db.ExecuteQuery("select * from t_table_setting where search_flg=1 order by table_id");
            combobox.ItemsSource = table.DefaultView;
        }

        /// <summary>
        /// プリンターの一覧をコンボボックスにセット
        /// </summary>
        /// <param name="combobox"></param>
        public static void SetPrinterList(ComboBox combobox)
        {
            combobox.Items.Clear();

            // LocalPrintServerを使用してローカルプリンターの一覧を取得
            var server = new LocalPrintServer();
            var queues = server.GetPrintQueues(new[] { EnumeratedPrintQueueTypes.Local, EnumeratedPrintQueueTypes.Connections });
            // PrintQueueのリストを取得
            combobox.ItemsSource = queues;
            
            // 既定のプリンターを取得し設定する
            var defaultPrinter = new System.Drawing.Printing.PrinterSettings().PrinterName;
            if (queues.Any(q => q.Name == defaultPrinter))
            {
                combobox.SelectedItem = queues.First(q => q.Name == defaultPrinter);
            }
            else
            {
                // デフォルトプリンターがリストにない場合は最初のプリンターを選択
                combobox.SelectedIndex = 0;
            }
        }

        /// <summary>
        /// FixedDocumentをプリンターで印刷
        /// </summary>
        /// <param name="document"></param>
        /// <param name="printer"></param>
        /// <param name="size"></param>
        /// <param name="orientation"></param>
        /// <param name="duplexing"></param>
        /// <param name="inputBin"></param>
        public static void FixedDocumentPrint(FixedDocument document, PrintQueue printer, Report.ParperSize size = Report.ParperSize.A4, 
                                           PageOrientation orientation = PageOrientation.Portrait, Duplexing duplexing = Duplexing.OneSided,
                                                InputBin inputBin = InputBin.AutoSelect)
        {
            var pt = printer.DefaultPrintTicket;


            // 印刷チケットの作成
            pt.PageMediaSize = Report.ParperSizeHelper.PageMediaSize(size); // 用紙サイズ
            pt.PageOrientation = orientation;   // 印刷方向
            pt.Duplexing = duplexing;   // 片面・両面印刷設定
            pt.InputBin = inputBin; // 給紙方法

            printer.DefaultPrintTicket = pt;

            // 印刷実行
            XpsDocumentWriter writer = PrintQueue.CreateXpsDocumentWriter(printer);
            writer.Write(document, pt);
        }

        /// <summary>
        /// IMEモード設定
        /// </summary>
        /// <param name="textBox"></param>
        /// <param name="mode"></param>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        public static void ImeMode(TextBox textBox, byte mode)
        {
            InputMethod.SetPreferredImeState(textBox, InputMethodState.On);

            switch (mode)
            {
                case 0: // 前半混在（文字数）
                case 1: // 全角
                case 11: // 前半混在（byte）
                    InputMethod.SetPreferredImeConversionMode(textBox, ImeConversionModeValues.Native | ImeConversionModeValues.FullShape);
                    break;
                case 2: // 全角数値
                case 3: // 全角英数
                    InputMethod.SetPreferredImeConversionMode(textBox, ImeConversionModeValues.Alphanumeric | ImeConversionModeValues.FullShape);
                    break;
                case 4: // 全角カナ
                    InputMethod.SetPreferredImeConversionMode(textBox, ImeConversionModeValues.Katakana | ImeConversionModeValues.FullShape);
                    break;
                case 5: // 半角
                case 6: // 半角数値
                case 7: // 半角英数
                case 9: // date(yyyymmdd)
                case 10: // date(yyyy/mm/dd)
                    InputMethod.SetPreferredImeConversionMode(textBox, ImeConversionModeValues.Alphanumeric);
                    break;
                case 8: // 半角カナ
                    InputMethod.SetPreferredImeConversionMode(textBox, ImeConversionModeValues.Katakana | ImeConversionModeValues.Native);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(mode), "Invalid IME mode specified.");
            }
        }
    }
}
