using Acornima.Ast;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Markup;

namespace MyTemplate.Report.Helpers
{
    public static class FubiHelper
    {
        /// <summary>
        /// DataTableの各行をPageContentに変換し、1つのFixedDocumentにまとめる
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="defectDic"></param>
        /// <returns></returns>
        public static FixedDocument CreateFixedDocument(DataTable dataTable, Dictionary<string, string> defectDic)
        {
            var document = new FixedDocument();

            foreach (DataRow row in dataTable.Rows)
            {
                var pageContents = CreatePageContent(row, defectDic);
                foreach(PageContent page in pageContents)
                {
                    // PageContentをFixedDocumentに追加
                    document.Pages.Add(page);
                }
            }

            return document;
        }

        /// <summary>
        /// DataRowを受け取り、ページごとの内容を作成する
        /// </summary>
        /// <param name="dataRow"></param>
        /// <param name="defectDic"></param>
        /// <returns></returns>
        private static List<PageContent> CreatePageContent(DataRow dataRow, Dictionary<string, string> defectDic)
        {
            var pageContents = new List<PageContent>();

            var size = ParperSize.A4.ToSSize(); // 用紙サイズを指定

            const int firstPageLimit = 29; // 1ページ目の上限
            const int otherPageLimit = 56; // 2ページ目以降の上限

            List<List<string>> pages = new();
            List<string> currentPage = new();
            int currentLine = 0; // 現在の行数
            int lineLimit = firstPageLimit; // 行数の上限を初期化

            // 郵便番号が７桁の数値の場合、ハイフンを挿入して「123-4567」の形式に変換
            var post_num = dataRow["post_num"].ToString().Trim();
            if (post_num.Length == 7 && int.TryParse(post_num, out _))
            {
                post_num = post_num.Insert(3, "-");
            }

            // PersonModelのインスタンスを作成
            var person = new Report.Models.PersonModel
            {
                qr_code = dataRow["qr_code"].ToString(),
                name = $"{dataRow["name"].ToString()}　御中",
                post_num = post_num,
                addr = dataRow["addr"].ToString(),
                fubi_code = dataRow["fubi_code"].ToString()
            };

            // ページの作成
            var fubi_ary = person.fubi_code.Split(',');
            foreach (var code in fubi_ary)
            {
                var lines = defectDic[code].ToString().Split(@"\n");
                int neededLines = lines.Length + 1;

                if (currentLine + neededLines > lineLimit)
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

            if (currentPage.Count > 0)
            {
                pages.Add(currentPage);
            }

            // ページごとにPersonViewModelを作成
            for (int i = 0; i < pages.Count; i++)
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
                UserControl view = i == 0
                    ? new Views.FubiPage()
                    : new Views.FubiPageN();

                // 不備内容の作成
                var sb = new StringBuilder();
                foreach (var line in pages[i])
                {
                    sb.AppendLine(line);
                }

                // 最後の行に次ページの案内を追加
                if (i < pages.Count - 1)
                {
                    sb.AppendLine();
                    sb.AppendLine("※次頁もご覧ください");
                }
                personPage.Item.fubi = sb.ToString();
                sb.Clear();

                // ページ数
                personPage.Item.page = i + 1;
                personPage.Item.max_page = pages.Count;
                personPage.Item.pages = personPage.Item.max_page > 1 ? $"Page. {personPage.Item.page} / {personPage.Item.max_page}" : "";

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

                pageContents.Add(pageContent);

                personPage = null;
                fixedPage = null;
                pageContent = null;
            }
            return pageContents;
        }

        /// <summary>
        /// DataRowを受け取り、FixedDocumentを作成する
        /// </summary>
        /// <param name="dataRow"></param>
        /// <param name="defectDic"></param>
        /// <returns></returns>
        public static FixedDocument CreateFixedDocument(DataRow dataRow, Dictionary<string, string> defectDic)
        {
            var document = new FixedDocument();

            var pageContents = CreatePageContent(dataRow, defectDic);
            foreach (PageContent page in pageContents)
            {
                // PageContentをFixedDocumentに追加
                document.Pages.Add(page);
            }

            return document;
        }
    }
}
