using MyLibrary;
using MyLibrary.MyModules;
using MyTemplate.ImportClass;
using MyTemplate.RoukinClass;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
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

namespace MyTemplate.RoukinForm
{
    /// <summary>
    /// Nouhin.xaml の相互作用ロジック
    /// </summary>
    public partial class Nouhin : Window
    {
        private DataTable _web = new();
        public Nouhin()
        {
            InitializeComponent();
        }

        private void bt_InsWebData_Click(object sender, RoutedEventArgs e)
        {
            using (var web = new FileLoadProperties())
            {
                if (!FileLoadClass.GetFileLoadSetting(6, web)) return;

                if (FileLoadClass.FileLoad(this, web) != MyLibrary.MyEnum.MyResult.Ok) return;

                _web = web.LoadData;
            }
        }

        private void bt_ExpNouhin_Click(object sender, RoutedEventArgs e)
        {
            string expPath = MyUtilityModules.AppSetting("roukin_setting", "exp_root_path");
            string expDir = MyUtilityModules.AppSetting("roukin_setting", "nouhin_dir");
            
            // 出力先指定
            expPath = MyTemplate.Modules.MyFolderDialog(System.IO.Path.Combine(expPath, expDir));
            // 出力先が指定されていない場合は処理を中止
            if (string.IsNullOrEmpty(expPath)) return;
            // 確認
            if (MyMessageBox.Show("納品データを出力します。よろしいですか？", "確認", 
                MyEnum.MessageBoxButtons.YesNo, MyEnum.MessageBoxIcon.None) != MyEnum.MessageBoxResult.Yes) return;

            // ローディングダイアログを表示
            using (var dlg = new MyLibrary.MyLoading.Dialog(this))
            {
                var exp = new NouhinClass(new DataTable(), _web, expPath);
                dlg.ThreadClass(exp);
                dlg.ShowDialog();

                if (exp.Result != MyEnum.MyResult.Ok) return;

                MyMessageBox.Show(exp.ResultMessage);
            }

        }
    }
}
