using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Markup;

namespace MyTemplate.Report.Helpers
{
    public static class FubiHelper
    {

        public static FixedDocument CreateFixedDocument(DataTable dataTable, Dictionary<string, string> defectDic)
        {
            var document = new FixedDocument();
            var size = ParperSize.A4.ToSSize(); // 用紙サイズを指定
            
            const int firstPageLimit= 5; // 1ページ目の上限
            const int otherPageLimit = 10; // 2ページ目以降の上限

            foreach (DataRow row in dataTable.Rows)
            {
                List<List<string>> pages = new();
                List<string> currentPage = new();
                int currentLine = 0; // 現在の行数
                int lineLimit = firstPageLimit; // 行数の上限を初期化

                // PersonModelのインスタンスを作成
                var person = new Report.Models.PersonModel
                {
                    qr_code = row["qr_code"].ToString(),
                    name = $"{row["name"].ToString()}　御中",
                    post_num = row["post_num"].ToString(),
                    addr = row["addr"].ToString(),
                    fubi_code = row["fubi_code"].ToString()
                };

                // ページの作成
                foreach (var code in person.fubi_ary)
                {
                    var lines = defectDic[code].ToString().Split(@"\n");
                    int neededLines = lines.Length + 1;

                    if(currentLine + neededLines > lineLimit)
                    {
                        pages.Add(currentPage);
                        currentPage = new List<string>();
                        currentLine = 0;
                        lineLimit = otherPageLimit; // 2ページ目以降は行数の上限を変更
                    }

                    currentPage.AddRange(lines);
                    currentPage.Add(""); // 空行を追加
                    currentLine += neededLines;
                }

                if(currentPage.Count > 0)
                {
                    pages.Add(currentPage);
                }

                for(int i=0; i < pages.Count; i++)
                {
                    var personPage = new Report.ViewModels.PersonViewModel
                    {
                        Person = person,
                        Item = new Models.ItemModule()
                    };

                    // QRコード値
                    personPage.Item.qr_code = personPage.Person.qr_code + (i + 1);
                    // QRコード生成
                    personPage.Item.Qr = QrCoderHelper.GenerateQrCode(personPage.Item.qr_code, 50);
                    // ページごとにテキストをリセット
                    personPage.Item.fubi = string.Empty;

                    // Viewの選択
                    // 1ページ目はFubiPage、2ページ目以降はFubiPageNを使用
                    UserControl view = i switch
                    {
                        0 => new Report.Views.FubiPage(),
                        _ => new Report.Views.FubiPageN()
                    };

                    // 不備内容の作成
                    foreach(var line in pages[i])
                    {
                        personPage.Item.fubi += line + "\n";
                    }

                    // 最後の行に次ページの案内を追加
                    if (i < pages.Count - 1)
                    {
                        personPage.Item.fubi += "\n" + "※次頁もご覧ください";
                    }

                    // ページ数の設定
                    personPage.Item.pages = $"Page. {i + 1} / {pages.Count}";

                    // ViewModelのDataContextに設定
                    view.DataContext = personPage;

                    // レイアウトの設定
                    view.Width = size.Width;
                    view.Height = size.Height;
                    view.Measure(size);
                    view.Arrange(new System.Windows.Rect(size));
                    view.UpdateLayout();

                    // FixedPageにViewを追加
                    var fixedPage = new System.Windows.Documents.FixedPage
                    {
                        Width = size.Width,
                        Height = size.Height
                    };
                    fixedPage.Children.Add(view);

                    // FixedPageをPageContentに追加
                    var pageContent = new System.Windows.Documents.PageContent();
                    ((IAddChild)pageContent).AddChild(fixedPage);

                    // PageContentをFixedDocumentに追加
                    document.Pages.Add(pageContent);
                }
            }
            return document;
        } 
    }
}
