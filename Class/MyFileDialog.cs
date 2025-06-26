using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MyLibrary;
using MyTemplate.Forms;
using System.Windows;
using System.Data;
using MyLibrary.MyClass;
using Microsoft.Win32;

namespace MyTemplate.Class
{
    internal class MyFileDialog
    {
        private List<string> _filePaths;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public MyFileDialog()
        {
            _filePaths = new List<string>();
        }

        public List<string> FilePaths { get { return _filePaths; } }

        /// <summary>
        /// ファイル選択ダイアログ表示（単一選択）
        /// </summary>
        /// <param name="initialPath"></param>
        /// <param name="extensions"></param>
        /// <param name="regPatterns"></param>
        /// <returns></returns>
        public MyEnum.MyResult Single(string initialPath = "", string[] extensions = null, string[] regPatterns = null)
        {
            try
            {
                OpenFileDialog dlg = new OpenFileDialog
                {
                    Title = "ファイル選択",
                    InitialDirectory = initialPath,
                    Filter = extensions != null ? string.Join("|", extensions.Select(ext => $"{ext.ToUpper()}|*.{ext.ToLower()}")) + "|All Files|*.*" : "All Files|*.*",
                    Multiselect = false,
                };

                if(dlg.ShowDialog() != true)
                {
                    return MyEnum.MyResult.Cancel;
                }

                // 選択されたファイル名を取得
                string file = dlg.FileName.ToString();

                // 正規表現チェック
                if (regPatterns != null)
                {
                    foreach (string reg in regPatterns)
                    {
                        if (System.Text.RegularExpressions.Regex.IsMatch(System.IO.Path.GetFileName(file), reg))
                        {
                            _filePaths = new List<string> { dlg.FileName };
                            break;
                        }
                        else
                        {
                            System.Windows.MessageBox.Show("ファイル名が不正です。");
                            return MyEnum.MyResult.Cancel;
                        }
                    }
                }
                else
                {
                    _filePaths = new List<string> { dlg.FileName };
                }
                return MyEnum.MyResult.Ok;
            }
            catch (Exception ex)
            {
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error);
                return MyEnum.MyResult.Error;
            }
        }

        /// <summary>
        /// フォルダ選択ダイアログ表示（複数選択）
        /// </summary>
        /// <param name="window"></param>
        /// <param name="initialPath"></param>
        /// <param name="regPatterns"></param>
        /// <returns></returns>
        public MyEnum.MyResult Mulit(Window window = null, string initialPath = "", string[] regPatterns = null)
        {
            // 初期フォルダ
            initialPath = Modules.MyFolderDialog(initialPath);
            // 初期フォルダが選択されていない場合はキャンセル
            if (initialPath == string.Empty)
            {
                return MyEnum.MyResult.Cancel;
            }

            // フォルダ選択ダイアログ表示
            var multi = new MultiSelect(window, initialPath, regPatterns);
            multi.ShowDialog();
            // 選択されたファイル名を取得
            _filePaths = multi.GetFilePaths();
            return MyEnum.MyResult.Ok;
        }

        /// <summary>
        /// ドラッグ＆ドロップダイアログ表示（複数選択）
        /// </summary>
        /// <param name="window"></param>
        /// <param name="regPatterns"></param>
        /// <returns></returns>
        public MyEnum.MyResult DragDrop(Window window, string[] regPatterns = null)
        {
            MyEnum.MyResult ressult;
            // ドラッグ＆ドロップダイアログ表示
            var drag = new Forms.DragDrop(window, regPatterns);
            ressult = drag.ShowDialog();
            // 選択されたファイル名を取得
            _filePaths = drag.GetFilePaths();
            return ressult;
        }
    }
}
