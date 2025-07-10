using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;
using System.Windows.Media.Imaging;

namespace MyTemplate.Report.Models
{
    public class PersonModel
    {
        public string name { get; set; }
        public string post_num { get; set; }
        public string addr { get; set; }
    }

    public class PersonItems
    {
        public string qr_code { get; set; }
        public int page { get; set; }                           // ページ数
        public int pages { get; set; }                          // 最大ページ数
        public List<FubiText> fubi { get; set; } = new();       // 不備内容のテキスト
        public BitmapImage qr_image { get; set; }                     // QRコード画像
    }

    public class FubiText
    {
        public string Fubi { get; set; } // 不備内容のテキスト
    }

    public class Hikinuki
    {
        public string bpo_num { get; set; } // BPO番号
        public string group_name { get; set; } // 団体名
    }
}
