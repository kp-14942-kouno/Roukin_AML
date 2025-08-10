using Acornima.Ast;
using DocumentFormat.OpenXml.Wordprocessing;
using MyTemplate.Report.Models;
using MyTemplate.Report.Views;
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
        public static FixedDocument CreateFixedDocument(DataTable dataTable, List<Report.Models.DefectModel> defectModels, List<Report.Models.BankModel> bankModel)
        {
            var document = new FixedDocument();

            foreach (DataRow row in dataTable.Rows)
            {
                var pageContents = CreatePageContent(row, defectModels, bankModel);
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
        private static List<PageContent> CreatePageContent(DataRow dataRow, List<Report.Models.DefectModel> defectModels, List<Report.Models.BankModel> bankModel)
        {
            var pageContents = new List<PageContent>();

            var size = ParperSize.A4.ToSSize(); // 用紙サイズを指定

            const int firstPageLimit = 24; // 1ページ目の上限
            const int otherPageLimit = 55; // 2ページ目以降の上限

            var person = new Models.PersonModel();

            // 郵便番号が７桁の数値の場合、ハイフンを挿入して「123-4567」の形式に変換
            var post_num = dataRow["bpo_zip_code"].ToString().Trim();
            if (post_num.Length == 7 && int.TryParse(post_num, out _))
            {
                post_num = post_num.Insert(3, "-");
            }
            person.name = $"{dataRow["bpo_org_kanji"].ToString()}　御中";
            person.post_num = post_num;
            person.addr = dataRow["bpo_address"].ToString();

            // 金融機関名を取得
            var bankName = bankModel.FirstOrDefault(x => x.code == dataRow["bpo_bank_code"].ToString())?.financial_name;
            person.bank_name = bankName;

            var pages = new List<PersonItems>();
            var fubiAry = dataRow["fubi_code"].ToString().Split(';');

            var pageNum = 1;
            var currentLine = 0; // 現在の行数
            var lineLimit = firstPageLimit; // 行数の上限を初期化
            var personItem = new PersonItems();

            personItem.fubi.Add(new FubiText { Fubi = "" }); // 開始行を1段下げる

            foreach (var code in fubiAry)
            {
                // 不備文言を取得
                var line = defectModels.ToList()
                    .Where(x => x.fubi_code == code)
                    .Select(x => x.fubi_caption).ToList()
                    .FirstOrDefault();
                    
                // 改行で配列化
                var lines = line.ToString().Split(@"\n", StringSplitOptions.None);

                // 必要な行数を計算
                int neededLines = lines.Length + 1;

                if (currentLine + neededLines > lineLimit)
                {
                    // QR
                    personItem.qr_code = $"{dataRow["bpo_num"].ToString()}{pageNum}";
                    personItem.qr_image = QrCoderHelper.GenerateQrCode(personItem.qr_code, 50);

                    pages.Add(personItem);
                    personItem = new PersonItems(); // 新しいPersonItemsを作成
                    personItem.fubi.Add(new FubiText { Fubi = "" }); // 開始行を1段下げる
                    pageNum++;
                    currentLine = 0;
                    lineLimit = otherPageLimit; // 2ページ目以降は行数の上限を変更
                }
                // 不備文言
                personItem.fubi.Add(new FubiText { Fubi = line.ToString().Replace(@"\n", Environment.NewLine) + Environment.NewLine });

                // ページ数
                personItem.page = pageNum;
                // 現在行
                currentLine += neededLines;
            }

            // 最後の行に次ページの案内を追加
            if (pageNum > pages.Count - 1)
            {
                personItem.qr_code = $"{dataRow["bpo_num"].ToString()}{pageNum}";
                personItem.qr_image = QrCoderHelper.GenerateQrCode(personItem.qr_code, 50);
                pages.Add(personItem);
            }

            for (int i = 0; i< pages.Count; i++)
            {
                if (i < pages.Count - 1)
                {
                    // 最後の行に次ページの案内を追加
                    pages[i].fubi.Add(new FubiText { Fubi = "※次頁もご覧ください" });
                }

                UserControl view = i == 0
                    ? new Views.FubiPage(person, pages[i], pages.Count)
                    : new Views.FubiPageN(pages[i], pages.Count);


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

            }

            return pageContents;
        }

        /// <summary>
        /// DataRowを受け取り、FixedDocumentを作成する
        /// </summary>
        /// <param name="dataRow"></param>
        /// <param name="defectDic"></param>
        /// <returns></returns>
        public static FixedDocument CreateFixedDocument(DataRow dataRow, List<Report.Models.DefectModel> defectModels, List<Report.Models.BankModel> bankModel)
        {
            var document = new FixedDocument();

            var pageContents = CreatePageContent(dataRow, defectModels, bankModel);
            foreach (PageContent page in pageContents)
            {
                // PageContentをFixedDocumentに追加
                document.Pages.Add(page);
            }

            return document;
        }
    }
}
