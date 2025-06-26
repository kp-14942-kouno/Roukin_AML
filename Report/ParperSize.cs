using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyTemplate.Report
{
    /// <summary>
    /// 用紙サイズを表す列挙型
    /// </summary>
    public enum ParperSize
    {
        A3 = 0,
        A4,
        A5,
        B3,
        B4,
        B5
    }

    /// <summary>
    /// 用紙サイズ（mm）をピクセルに変換する拡張メソッド
    /// </summary>
    public static class ParperSizeExtensions
    {
        public static System.Windows.Size ToSSize(this ParperSize size)
        {
            // 用紙サイズ（mm)をピクセルに変換するための定数
            return size switch
            {
                ParperSize.A3 => new System.Windows.Size(1123, 1587), // A3: 297mm x 420mm
                ParperSize.A4 => new System.Windows.Size(795, 1123),  // A4: 210mm x 297mm
                ParperSize.A5 => new System.Windows.Size(561, 795),   // A5: 148mm x 210mm
                ParperSize.B3 => new System.Windows.Size(1375, 1947), // B3: 364mm x 515mm
                ParperSize.B4 => new System.Windows.Size(972, 1375),  // B4: 257mm x 364mm
                ParperSize.B5 => new System.Windows.Size(688, 972),   // B5: 182mm x 257mm
                _ => throw new ArgumentOutOfRangeException(nameof(size), size, null)
            };
        }
    }
}
