using DocumentFormat.OpenXml.Spreadsheet;
using MyTemplate.Report.Models;
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
    /// FubiFuchakuSiwake.xaml の相互作用ロジック
    /// </summary>
    public partial class FubiFuchakuSiwake : UserControl
    {
        public FubiFuchakuSiwake(List<Models.Siwake> items, string bankCode, string finantialName, int page, int pages, BitmapImage qrCode)
        {
            InitializeComponent();

            // 1ページ目にはQRコードを表示
            if (page == 1)
            {
                img_QR.Source = qrCode;
            }
            
            tb_BankCode.Text = bankCode;
            tb_BankName.Text = finantialName;
            ItemList.ItemsSource = items;

            // ページ番号の表示
            tb_Pages.Text = $"Page. {page} / {pages}";
        }
    }
}
