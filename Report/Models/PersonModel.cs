using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
        public string title { get { return "お取引目的等の確認書"; } }
        public string caption
        {
            get
            {
                var text = "平素は格別のお引き立てを賜り、厚くお礼を申しあげます。\r\n" +
                           "さて、このたびはお取引目的等の確認書をご送付いただき、ありがとうございました。\r\n" +
                           "ご送付いただきましたお取引目的等の確認書のご記入内容に一部不足等がございましたので、一旦ご返却\r\n" +
                           "させていただきます。恐れ入りますが、以下の不足内容等をご確認のうえ追加記入・訂正等をいただき、同封の\r\n" +
                           "返信用封筒にてご返送をお願いいたします。";

                return text;
            }
        }

        public string post_num_fix { get
            {
                // 郵便番号が７桁の数値の場合、ハイフンを挿入して「123-4567」の形式に変換
                if (post_num.Length == 7 && int.TryParse(post_num, out _))
                {
                    return post_num.Insert(3, "-");
                }
                else
                {
                    return post_num; // それ以外はそのまま返す
                }
            }
        }
        public string[] fubi_ary { get { return fubi_code.Split(','); }} // 文字列をカンマで分割して配列に変換
    }

    public class ItemModule
    {
        public string qr_code { get; set; }     // QRコードの値
        public string page { get; set; }        // ページ数
        public string fubi { get; set; }        // 不備内容のテキスト
        public string pages { get; set; }       // 印刷用ページ数
        public BitmapImage Qr { get; set; }     // QRコード画像
    }
}
