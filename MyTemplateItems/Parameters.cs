namespace MyTemplate
{
    /// <summary>
    /// ログイン設定
    /// </summary>
    public class LoginParam
    {
        public int id { get; set; }
        public string schema { get; set; } = string.Empty;
        public string table_name { get; set; } = string.Empty;
        public int user_length_min { get; set; }
        public int user_length_max { get; set; }
        public int pass_length_min { get; set; }
        public int pass_length_max { get; set; }
        public int user_name_length_min { get; set; }
        public int user_name_length_max { get; set; }
    }
}
