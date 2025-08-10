using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;
using System.Windows.Media.Imaging;

namespace MyTemplate.Report.Models
{
    /// <summary>
    /// 団体情報モデル
    /// </summary>
    public class PersonModel
    {
        public string name { get; set; }
        public string post_num { get; set; }
        public string addr { get; set; }
        public string bank_name { get; set; }               // 金融機関名
    }

    /// <summary>
    /// 不備情報モデル
    /// </summary>
    public class PersonItems
    {
        public string qr_code { get; set; }
        public int page { get; set; }                       // ページ数
        public int pages { get; set; }                      // 最大ページ数
        public List<FubiText> fubi { get; set; } = new();   // 不備内容のテキスト
        public BitmapImage qr_image { get; set; }           // QRコード画像
    }

    /// <summary>
    /// 不備文言モデル
    /// </summary>
    public class FubiText
    {
        public string Fubi { get; set; } // 不備内容のテキスト
    }

    /// <summary>
    /// 不備文言モデル
    /// </summary>
    public class DefectModel
    {
        public string fubi_code { get; set; } // 不備コード
        public string fubi_caption { get; set; } // 不備名
    }

    /// <summary>
    /// 引抜リストモデル
    /// </summary>
    public class Hikinuki
    {
        public string bpo_num { get; set; } // BPO番号
        public string group_name { get; set; } // 団体名
    }

    /// <summary>
    /// 仕分けリストモデル
    /// </summary>
    public class Siwake 
    {
        public string taba_num { get; set; } // 束番号
        public string bpo_num { get; set; } // BPO番号
        public string group_name { get; set; } // 団体名
}

    /// <summary>
    /// 金融機関情報モデル
    /// </summary>
    public class BankModel
    {
        public string code { get; set; } // 金融機関コード
        public string branch_number { get; set; } // 支店番号
        public string financial_name { get; set; } // 金融機関名
    }

    /// <summary>
    /// 申請書明細モデル
    /// </summary>
    public class Meisai
    {
        public string no { get; set; } // 通番
        public string bpo_num { get; set; } // BPO番号
        public string bpo_persona_cd { get; set; } // 人格コード
        public string bpo_org_kanji { get; set; } // 団体名
        public string bpo_cust_no { get; set; } // 顧客番号
    }
}
