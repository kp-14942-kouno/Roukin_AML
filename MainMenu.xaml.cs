using DocumentFormat.OpenXml.Drawing.Diagrams;
using MyLibrary;
using MyLibrary.MyModules;
using MyTemplate.ImportClass;
using MyTemplate.Report.ViewModels;
using MyTemplate.Report.Views;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Diagnostics;
using System.IO;
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
using System.Windows.Shapes;

namespace MyTemplate
{
    /// <summary>
    /// MainMenu.xaml の相互作用ロジック
    /// </summary>
    public partial class MainMenu : Window
    {
        public MainMenu()
        {
            InitializeComponent();
        }

        private void FixedDocumentAsJpeg(FixedDocument document)
        {
            string dir = @"D:\04.VisualStudio\ろうきん";

            foreach(var pageContent in document.Pages)
            {
                if(pageContent is PageContent pc)
                {
                    var fixedPage = pc.GetPageRoot(false);
                    if(fixedPage == null) continue;

                    var child = fixedPage.Children[0];
                    string qrCode = string.Empty;

                    if (child is FubiPage view)
                    {
                        if (view.DataContext is PersonViewModel person)
                        {
                            qrCode = person.Item.qr_code;
                        }

                    }
                    else if (child is FubiPageN viewN)
                    {
                        if (viewN.DataContext is PersonViewModel person)
                        {
                            qrCode = person.Item.qr_code;
                        }
                    }

                    const double dpi = 200;
                    double width = fixedPage.Width;
                    double height = fixedPage.Height;

                    var pixelWidth = (int)(width * dpi / 96);
                    var pixelHeight = (int)(height * dpi / 96);

                    var renderBitmap = new RenderTargetBitmap(pixelWidth, pixelHeight, dpi, dpi, PixelFormats.Pbgra32);
                    fixedPage.Measure(new Size(width, height));
                    fixedPage.Arrange(new Rect(new Size(width, height)));
                    renderBitmap.Render(fixedPage);

                    var encoder = new JpegBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(renderBitmap));

                    string outputFile = System.IO.Path.Combine(dir, $"{qrCode}.jpg");
                    using(var stream = new FileStream(outputFile, FileMode.Create))
                    {
                        encoder.Save(stream);
                    }
                }
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var defectDic = new Dictionary<string, string>();

            using (MyDbData db = new MyDbData("setting"))
            {
                using (DbDataReader reader = db.ExecuteReader("select * from t_fubi_code order by fubi_code"))
                {
                    while (reader.Read())
                    {
                        defectDic.Add(reader["fubi_code"].ToString(), reader["fubi_caption"].ToString());
                    }
                }
            }

            // ファイル読込みプロパティ
            using (FileLoadProperties file = new FileLoadProperties())
            {
                // ファイル読込み設定
                if (!FileLoadClass.GetFileLoadSetting(50, file)) return;
                // ファイル読込み
                if (FileLoadClass.FileLoad(this, file) != MyLibrary.MyEnum.MyResult.Ok) return;

                var docment = Report.Helpers.FubiHelper.CreateFixedDocument(file.LoadData, defectDic);

                if (docment.Pages.Count != 0)
                {

                    dv_Report.Document = docment;
                    //FixedDocumentAsJpeg(docment);


                    /*
                    // 印刷ダイアログを表示
                    PrintDialog printDialog = new PrintDialog();
                    if (printDialog.ShowDialog() == true)
                    {
                        // 印刷設定を適用
                        printDialog.PrintDocument(docment.DocumentPaginator, "Fubi Report");
                    }
                    */
                }
                else
                {
                    MessageBox.Show("印刷するデータがありません。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }
}
