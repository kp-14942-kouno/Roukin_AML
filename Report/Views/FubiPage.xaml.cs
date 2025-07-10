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
    /// FubiPage.xaml の相互作用ロジック
    /// </summary>
    public partial class FubiPage : UserControl
    {
        public string QrCode => tb_QrCode.Text;

        public FubiPage(PersonModel model, PersonItems items, int pages)
        {
            InitializeComponent();

            tb_QrCode.Text = items.qr_code;
            tb_PostNum.Text = model.post_num;  
            tb_Addr.Text = model.addr;
            tb_Name.Text = model.name;
            img_QR.Source = items.qr_image;

            FubiList.ItemsSource = items.fubi;

            if(pages > 1)
            {
                tb_Pages.Text = $"{items.page} / {pages}";
            }
        }
    }
}
