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
    public static class HikinukiHelper
    {
        public static FixedDocument CreateFixedDocument(DataTable table, string taba)
        {
            List<Models.Hikinuki> hikinukiList = ConvTable(table);
            
            var size = MyTemplate.Report.ParperSize.A4.ToSSize(); // 用紙サイズを指定
            var pages = ChunkBy(hikinukiList, 10);
            var count = 0;

            FixedDocument fiexedDoc = new FixedDocument();
            foreach (var pageData in pages)
            {
                count++;
                var page = new Report.Views.HikinukiPage(pageData, taba, count, pages.Count());

                FixedPage fixedPage = new FixedPage { Height = size.Height, Width = size.Width };
                fixedPage.Children.Add(page);

                page.Measure(size);
                page.Arrange(new System.Windows.Rect(size));
                page.UpdateLayout();

                var pageContent = new PageContent();
                ((System.Windows.Markup.IAddChild)pageContent).AddChild(fixedPage);

                fiexedDoc.Pages.Add(pageContent);
            }

            return fiexedDoc;
        }

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

        public static List<List<T>> ChunkBy<T>(List<T> source, int chunkSize)
        {
            if (chunkSize <= 0)
                throw new ArgumentException("Chunk size must be greater than zero.", nameof(chunkSize));
            return source.Select((x, i) => new { Index = i, Value = x })
                         .GroupBy(x => x.Index / chunkSize)
                         .Select(g => g.Select(x => x.Value).ToList())
                         .ToList();
        }
    }
}
