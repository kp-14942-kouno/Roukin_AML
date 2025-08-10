using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;

namespace MyTemplate.Report.Helpers
{
    public static class SinseishoMeisaiHelper
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
            List<Models.Meisai> meisaiList = ConvTable(table);

            // 用紙サイズ
            var size = MyTemplate.Report.ParperSize.A4.ToSSize();

            // 引抜リストのデータを10件ずつのページに分割
            var pages = ReportModules.ChunkBy(meisaiList, 35);

            var count = 0;

            // ページごとに分割してFixedDocumentを作成
            FixedDocument fiexedDoc = new FixedDocument();
            foreach (var pageData in pages)
            {
                count++;
                // ページ作成
                var page = new Report.Views.SinseishoMeisai(pageData, financialName, count, pages.Count());

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
        /// 明細リストのデータをリストに変換
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        private static List<Models.Meisai> ConvTable(DataTable table)
        {
            var meisaiList = new List<Models.Meisai>();
            var num = 0;

            foreach (DataRow row in table.Rows)
            {
                num++;

                var meisai = new Models.Meisai
                {
                    no = num.ToString(),
                    bpo_num = row["bpo_num"].ToString(),
                    bpo_persona_cd = row["bpo_persona_cd"].ToString(),
                    bpo_cust_no = row["bpo_cust_no"].ToString(),
                    bpo_org_kanji = row["bpo_org_kanji"].ToString()
                };
                meisaiList.Add(meisai);
            }
            return meisaiList;
        }
    }
}
