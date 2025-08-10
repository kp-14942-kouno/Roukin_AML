using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;

namespace MyTemplate.Report.Helpers
{
    /// <summary>
    /// 引抜リスト印刷用ヘルパークラス
    /// </summary>
    public static class HikinukiHelper
    {
        /// <summary>
        /// FixedDocumentを作成する
        /// </summary>
        /// <param name="table"></param>
        /// <param name="taba"></param>
        /// <returns></returns>
        public static FixedDocument CreateFixedDocument(DataTable table, string taba)
        {
            // 引抜リストのデータを変換
            List<Models.Hikinuki> hikinukiList = ConvTable(table);
            
            // 用紙サイズ
            var size = MyTemplate.Report.ParperSize.A4.ToSSize();

            // 引抜リストのデータを10件ずつのページに分割
            var pages = ReportModules.ChunkBy(hikinukiList, 40);

            var count = 0;

            // ページごとに分割してFixedDocumentを作成
            FixedDocument fiexedDoc = new FixedDocument();
            foreach (var pageData in pages)
            {
                count++;
                // ページ作成
                var page = new Report.Views.HikinukiPage(pageData, taba, count, pages.Count());

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
        /// 引抜リストのデータをリストに変換
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        private static List<Models.Hikinuki> ConvTable(DataTable table)
        {
            List<Models.Hikinuki> hikinukiList = new List<Models.Hikinuki>();
            foreach (DataRow row in table.Rows)
            {
                var hikinuki = new Models.Hikinuki
                {
                    bpo_num = row["bpo_num"].ToString(),
                    group_name = row["bpo_org_kanji"].ToString()
                };
                hikinukiList.Add(hikinuki);
            }
            return hikinukiList;
        }
    }
}
