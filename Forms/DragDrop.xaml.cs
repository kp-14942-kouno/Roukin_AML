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

namespace MyTemplate.Forms
{
    /// <summary>
    /// DragDrop.xaml の相互作用ロジック
    /// </summary>
    public partial class DragDrop : Window
    {
        private Brush _color;
        private string[] _fileNameRegs;
        private MyLibrary.MyEnum.MyResult _result = MyLibrary.MyEnum.MyResult.Cancel;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="fileNameRegs"></param>
        public DragDrop(Window window, string[] fileNameRegs)
        {
            InitializeComponent();

            // ウィンドウの中央に表示
            this.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            this.Owner = window;

            _color = dg_FileList.Background;
            _fileNameRegs = fileNameRegs;
        }

        /// <summary>
        /// ダイアログ表示
        /// </summary>
        /// <returns></returns>
        public new MyLibrary.MyEnum.MyResult ShowDialog()
        {
            base.ShowDialog();
            return _result;
        }

        /// <summary>
        /// DragEnterイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DataGrid_DragEnter(object sender, DragEventArgs e)
        {
            // ファイルがDragされたら背景色を変える
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false))
            {
                e.Effects = DragDropEffects.All;
                dg_FileList.Background = Brushes.LightPink;
            }
        }

        /// <summary>
        /// DragLeaveイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DataGrid_DragLeave(object sender, DragEventArgs e)
        {
            // Dragが終了したら背景色を元に戻す
            dg_FileList.Background = _color;
        }

        /// <summary>
        /// Dropイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DataGrid_Drop(object sender, DragEventArgs e)
        {
            string[] filePaths = e.Data.GetData(DataFormats.FileDrop) as string[];

            foreach (string filePath in filePaths)
            {
                // ファイルのみ処理
                if (!System.IO.File.Exists(filePath))
                {
                    continue;
                }
                // 拡張子を取得
                string extention = System.IO.Path.GetExtension(filePath).ToLower();
                // ファイル名
                string fileName = System.IO.Path.GetFileName(filePath);

                // ファイルパラメータ分繰返す
                foreach (string fileNameReg in _fileNameRegs)
                {
                    // 正規表現
                    if (!Regex.IsMatch(fileName, fileNameReg))
                    {
                        continue;
                    }
                    // 登録済み
                    if (dg_FileList.Items.Cast<dynamic>().Any(item => item.file_path == filePath)) continue;

                    // 追加
                    dg_FileList.Items.Add(new { file_path = filePath, file_name = fileName });
                    dg_FileList.Items.Refresh();
                }
            }
            // ファイル数を表示
            this.Count.Content = dg_FileList.Items.Count;
            // 背景色を元に戻す
            this.dg_FileList.Background = _color;
        }

        /// <summary>
        /// 削除ボタンクリックイベント
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
        /// OKボタンイベント
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
        /// Cancelボタンイベント
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
