using DocumentFormat.OpenXml.Bibliography;
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
    /// HikinukiPage.xaml の相互作用ロジック
    /// </summary>
    public partial class HikinukiPage : UserControl
    {
        public HikinukiPage(List<Models.Hikinuki> items, string taba, int page, int pages)
        {
            InitializeComponent();

            ItemList.ItemsSource = items;
            tb_Tabe.Text = taba;
        }
    }
}
