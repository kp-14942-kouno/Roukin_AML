using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using MyLibrary;

namespace MyTemplate.Forms
{
    /// <summary>
    /// MultiSelect.xaml の相互作用ロジック
    /// </summary>
    public partial class MultiSelect : Window
    {
        private string _initialPath;
        private string[] _fileNameRegs;

        private MyLibrary.MyEnum.MyResult _result = MyLibrary.MyEnum.MyResult.Cancel;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="window"></param>
        /// <param name="initialPath"></param>
        /// <param name="fileNameRegs"></param>
        public MultiSelect(Window window, string initialPath, string[] fileNameRegs)
        {
            InitializeComponent();

            this.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            this.Owner = window;

            _initialPath = initialPath;
            _fileNameRegs = fileNameRegs;

            Reload();
        }

        public new MyEnum.MyResult ShowDialog()
        {
            base.ShowDialog();
            return _result;
        }

        /// <summary>
        /// Reload処理
        /// </summary>
        private void Reload()
        {
            // ファイルリストの初期化
            dg_FileList.Items.Clear();

            // 指定フォルダ内のファイルを取得
            foreach (string filePath in System.IO.Directory.GetFiles(_initialPath))
            {
                string fileName = System.IO.Path.GetFileName(filePath);

                foreach (string fileNameReg in _fileNameRegs)
                {
                    // 正規表現
                    if (Regex.IsMatch(fileName, fileNameReg))
                    {
                        // 追加
                        dg_FileList.Items.Add(new { file_path = filePath, file_name = fileName });
                    }
                }
            }
            dg_FileList.Items.Refresh();
        }

        /// <summary>
        /// Reloadボタンクリックイベント処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Reload_Click(object sender, RoutedEventArgs e)
        {
            Reload();
        }

        /// <summary>
        /// 削除ボタンクリックイベント処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            // 選択行を削除
            if (dg_FileList.SelectedIndex >= 0)
            {
                dg_FileList.Items.RemoveAt(dg_FileList.SelectedIndex);
            }
        }

        /// <summary>
        /// OKボタンクリックイベント処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OK_Click(object sender, RoutedEventArgs e)
        {
            // ファイルが無い場合はCancel
            if (dg_FileList.Items.Count == 0)
            {
                _result = MyLibrary.MyEnum.MyResult.Cancel;
            }
            else
            {
                _result = MyLibrary.MyEnum.MyResult.Ok;
            }
            this.Close();
        }

        /// <summary>
        /// CANCELボタンクリックイベント処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CANCEL_Click(object sender, RoutedEventArgs e)
        {
            _result = MyLibrary.MyEnum.MyResult.Cancel;
            this.Close();
        }

        /// <summary>
        /// ファイルパスをList<string>で返す
        /// </summary>
        /// <returns></returns>
        public List<string?> GetFilePaths()
        {
            // DataGridのfile_path項目をList化
            return dg_FileList.Items.Cast<dynamic>().Select(item => item.file_path as string).ToList();
        }
    }
}
