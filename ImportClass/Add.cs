using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.InkML;
using MyLibrary;
using MyLibrary.MyClass;
using MyLibrary.MyModules;
using MyTemplate.Class;

namespace MyTemplate.ImportClass
{
    /// <summary>
    /// 追加情報クラス
    /// </summary>
    internal static class Add
    {
        /// <summary>
        /// 追加情報取得処理
        /// </summary>
        /// <param name="addSettings"></param>
        /// <param name="window"></param>
        /// <returns></returns>
        public static MyEnum.MyResult Run(FileLoadProperties load, System.Windows.Window window = null)
        {
            try
            {
                MyStandardCheck check = new MyStandardCheck();

                foreach (AddSetting add in load.AddSettings)
                {
                    string? value = string.Empty;
                    string msg = string.Empty;
                    int errNo = 0;

                    switch (add.add_type)
                    {
                        // 値
                        case 0:
                            value = add.add_value.ToString();
                            break;

                        // 入力値
                        case 1:
                            bool res;
                            do
                            {
                                res = true;
                                value = MyInputBox.Show(window, add.column_caption, "追加情報");
                                // 戻り値がnullはキャンセル
                                if (value == null) return MyEnum.MyResult.Cancel;

                                // 入力値チェック
                                (errNo, msg) = check.GetResult(value, add.column_type, add.column_length, add.column_null, add.column_fix, add.column_reg);
                                // エラーは中断
                                if (!string.IsNullOrEmpty(msg))
                                {
                                    MyMessageBox.Show(msg);
                                    res = false;
                                }
                            }
                            while (!res);
                                
                            // トリム
                            if (add.column_trim == 1) value = value.Trim();

                            // 外部メソッド呼び出し
                            if (!string.IsNullOrEmpty(add.method_name))
                                load.js.Invoke(add.method_name, value);

                            add.value = value;
                            continue;

                        // フォーム値
                        case 2:
                            if (window != null)
                            {
                                object control = window.FindName(add.add_value.ToString());
                                if (control != null)
                                {
                                    switch (control)
                                    {
                                        case TextBox textBox:
                                            value = textBox.Text;
                                            break;
                                        case ComboBox comboBox:
                                            value = comboBox.SelectedValue.ToString();
                                            break;
                                        case CheckBox checkBox:
                                            value = checkBox.IsChecked.ToString();
                                            break;
                                        case RadioButton radioButton:
                                            value = radioButton.IsChecked.ToString();
                                            break;
                                        case TextBlock textBlock:
                                            value = textBlock.Text;
                                            break;
                                    }
                                }
                            }
                            break;

                        // DateTime
                        case 3:
                            value = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                            break;
                        
                        // yyyyMMdd
                        case 4:
                            value = DateTime.Now.ToString("yyyyMMdd");
                            break;

                        // yyyy/MM/dd
                        case 5:
                            value = DateTime.Now.ToString("yyyy/MM/dd");
                            break;

                        default:
                            continue;
                    }

                    // トリム
                    if(add.column_trim == 1) value = value.Trim();

                    // 値のチェック
                    (errNo, msg) = check.GetResult(value, add.column_type, add.column_length, add.column_null, add.column_fix, add.column_reg);

                    // エラーは中断
                    if (!string.IsNullOrEmpty(msg))
                    {
                        MyMessageBox.Show($"追加情報項目「{add.column_caption}」\r\n{msg}");
                        return MyEnum.MyResult.Cancel;
                    }

                    // 外部メソッド呼び出し
                    if (!string.IsNullOrEmpty(add.method_name))
                        value = load.js.Invoke(add.method_name, value).ToString();

                    // 値をセット
                    add.value = value;
                }
            }
            catch(Exception ex)
            {
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error);
                return MyEnum.MyResult.Error;
            }
            return MyEnum.MyResult.Ok;
        }
    }
}
