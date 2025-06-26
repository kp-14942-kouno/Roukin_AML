using DocumentFormat.OpenXml.Drawing.Diagrams;
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
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
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
        public static string MyFolderDialog(string initialPath = null)
        {
            var dlg = new OpenFolderDialog()
            {
                Title = "フォルダ選択",
                InitialDirectory = initialPath,
                ShowHiddenItems=false
            };

            if(dlg.ShowDialog() != true)
            {
                return string.Empty;
            }

            return Path.GetFullPath(dlg.FolderName);
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
