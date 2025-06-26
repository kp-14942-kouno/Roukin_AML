using MyLibrary.MyModules;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static MyTemplate.UserInfo;

namespace MyTemplate
{
    /// <summary>
    /// ユーザー情報
    /// </summary>
    public static class UserInfo
    {
        // ユーザー名
        public static string UserName { get; set; } = string.Empty;
        // 権限
        public static int Authority { get; set; }

        // 権限名称に変換
        public static string AuthorityName { get
            {
                var autoritys = MyUtilityModules.AppSettings("authority");
                var authority = autoritys.FirstOrDefault(x => Equals(x.Value.ToString(), UserInfo.Authority.ToString()));
                return authority.Key;
            } }
    }
}
