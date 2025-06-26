using System.Globalization;
using System.Windows.Data;

namespace MyTemplate.Class
{
    /// <summary>
    /// Dictionaryの値からキーを逆引きして表示用に変換する汎用IValueConverter
    /// 
    /// 主にWPFのバインディングで、値を対応する表示名（キー）に変換したい場合に利用
    /// 使い方:
    /// - LookupDictionary プロパティ、または ConverterParameter で
    ///   Dictionary&lt;string, object&gt; を指定
    /// - Convertメソッドは、値から対応するキー（表示名）を返す
    /// - ConvertBackは未実装
    /// </summary>
    public class ReverseLockupConverter : IValueConverter
    {
        /// 逆引きに使用するDictionary
        /// ConverterParameterで渡された場合はそちらが優先
        //public Dictionary<string ,object> LookupDictionary {  get; set; }

        /// <summary>
        /// 値からDictionaryのキー（表示名）を逆引きして返す
        /// </summary>
        /// <param name="value">バインディングされた値（例: コードやID）</param>
        /// <param name="targetType">ターゲットの型</param>
        /// <param name="parameter">ConverterParameterで渡されたDictionary（任意）</param>
        /// <param name="culture">カルチャ情報</param>
        /// <returns>対応するキー（表示名）、該当しない場合は値のToString()</returns>
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            // ConverterParameter優先、なければプロパティのDictionaryを使用
            var dict = parameter as Dictionary<string, object>; //?? LookupDictionary;
            if (dict != null && value != null)
            {
                // 値からキーを逆引き
                var key = dict.FirstOrDefault(x => x.Value?.ToString() == value.ToString()).Key;
                if (!string.IsNullOrEmpty(key)) return key;
            }
            // 該当しない場合は値をそのまま返す
            return value?.ToString() ?? "";
        }

        /// <summary>
        /// 逆変換は未実装
        /// </summary>
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    /// <summary>
    /// 複数値（値＋辞書）を受け取り、値から辞書のキー（表示名）を逆引きするIMultiValueConverterの実装例。
    /// </summary>
    public class ReverseLookupMultiConverter : IMultiValueConverter
    {
        /// <summary>
        /// values[0]: 変換したい値（例: コードやID）
        /// values[1]: Dictionary<string, object>（例: AuthorityDict）
        /// </summary>
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values.Length < 2) return values.FirstOrDefault()?.ToString() ?? "";

            var value = values[0];
            var dict = values[1] as Dictionary<string, object>;
            if (dict != null && value != null)
            {
                var key = dict.FirstOrDefault(x => x.Value?.ToString() == value.ToString()).Key;
                if (!string.IsNullOrEmpty(key)) return key;
            }
            return value?.ToString() ?? "";
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
