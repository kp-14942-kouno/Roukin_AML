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
        public string qr_code { get; set; }
        public string name { get; set; }
        public string post_num { get; set; }
        public string addr { get; set; }
        public string fubi_code { get; set; }
    }

    public class ItemModule
    {
        public string qr_code { get; set; }     // QRコードの値
        public int page { get; set; }           // ページ数
        public int max_page { get; set; }       // 最大ページ数
        public string pages { get; set; }// 印刷用ページ数
        public string fubi { get; set; }        // 不備内容のテキスト
        public BitmapImage Qr { get; set; }     // QRコード画像
    }
}
