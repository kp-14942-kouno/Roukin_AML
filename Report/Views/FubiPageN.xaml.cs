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
    /// FubiPageN.xaml の相互作用ロジック
    /// </summary>
    public partial class FubiPageN : UserControl
    {
        public string QrCode { get; private set; } = string.Empty;

        public FubiPageN(PersonItems items, int pages)
        {
            InitializeComponent();

            QrCode = items.qr_code;
            FubiList.ItemsSource = items.fubi;
            imb_QR.Source = items.qr_image;
            tb_Pages.Text = $"{items.page} / {pages}";
        }
    }
}
