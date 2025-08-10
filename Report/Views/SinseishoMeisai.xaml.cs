using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace MyTemplate.Report.Views
{
    /// <summary>
    /// SinseishoMeisai.xaml の相互作用ロジック
    /// </summary>
    public partial class SinseishoMeisai : UserControl
    {
        public SinseishoMeisai(List<Models.Meisai> items, string bankName, int page, int pages)
        {
            InitializeComponent();

            tb_BankName.Text = $"{bankName.Replace("労金","")}労働金庫様　団体確認書の納品明細";
            ItemList.ItemsSource = items;
            tb_Pages.Text = $"Page. {page} / {pages}";
        }
    }
}
