using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;

namespace MyTemplate.Report.Helpers
{
    public static class FubiFuchakuHelper
    {
        /// <summary>
        /// FixedDocumentを作成する
        /// </summary>
        /// <param name="table"></param>
        /// <param name="taba"></param>
        /// <returns></returns>
        public static FixedDocument CreateFixedDocument(DataTable table, string code, string financialName)
        {
            // 引抜リストのデータを変換
            List<Models.Siwake> siwakeList = ConvTable(table);

            // 用紙サイズ
            var size = MyTemplate.Report.ParperSize.A4.ToSSize();

            // 引抜リストのデータを10件ずつのページに分割
            var pages = ReportModules.ChunkBy(siwakeList, 10);

            var count = 0;

            // ページごとに分割してFixedDocumentを作成
            FixedDocument fiexedDoc = new FixedDocument();
            foreach (var pageData in pages)
            {
                count++;
                // ページ作成
                var page = new Report.Views.FubiFuchakuSiwake(pageData, code, financialName, count, pages.Count());

                // ページのサイズを設定
                FixedPage fixedPage = new FixedPage { Height = size.Height, Width = size.Width };
                fixedPage.Children.Add(page);
                // ページのレイアウトを更新
                page.Measure(size);
                page.Arrange(new System.Windows.Rect(size));
                page.UpdateLayout();

                // ページのコンテンツをFixedPageに追加
                var pageContent = new PageContent();
                ((System.Windows.Markup.IAddChild)pageContent).AddChild(fixedPage);
                // ページをFixedDocumentに追加
                fiexedDoc.Pages.Add(pageContent);
            }

            return fiexedDoc;
        }

        /// <summary>
        /// 仕分けリストのデータをリストに変換
        /// 前行の束番号が同じならセットしない
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        private static List<Models.Siwake> ConvTable(DataTable table)
        {
            List<Models.Siwake> siwakeList = new List<Models.Siwake>();
            string previousTaba = string.Empty;

            foreach (DataRow row in table.Rows)
            {
                string tabaNum = string.Empty;

                if (previousTaba != row["taba_num"].ToString())
                {
                    tabaNum = row["taba_num"].ToString();
                    previousTaba = tabaNum;
                }

                var siwake = new Models.Siwake
                {
                    bpo_num = row["bpo_num"].ToString(),
                    taba_num = tabaNum,
                    group_name = row["bpo_org_kanji"].ToString()
                };
                siwakeList.Add(siwake);
            }
            return siwakeList;
        }
    }
}
