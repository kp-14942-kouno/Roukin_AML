using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Xps;
using System.Windows.Xps.Packaging;

namespace MyTemplate.Report
{
    /// <summary>
    /// ReportPreivew.xaml の相互作用ロジック
    /// </summary>
    public partial class ReportPreivew : Window
    {
        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="document"></param>
        public ReportPreivew(FixedDocument document)
        {
            InitializeComponent();

            this.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            this.Owner = Application.Current.MainWindow;

            dv_Preview.Document = document;

            // FixedDocumentを受け取り、DataContextに設定する
            //var doc = new Report.Models.DocumentModel
            //{
            //    Document = document
            //};

            //this.DataContext = doc;
        }
    }
}
