using DocumentFormat.OpenXml.Vml;
using Jint;
using MyLibrary.MyModules;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MyTemplate.Class
{
    /// <summary>
    /// Jint(JavaScript)クラス
    /// </summary>
    internal class MyJScriptClass: IDisposable
    {
        Engine _engine;

        /// <summary>
        /// 開放
        /// </summary>
        public void Dispose()
        {
            _engine.Dispose();
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="nodeKey"></param>
        public MyJScriptClass(string nodeKey)
        {
            _engine = new Engine();

            var dir = MyUtilityModules.AppSetting("script", "dir");
            var js = MyUtilityModules.AppSetting("script", nodeKey);
            var path = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, dir, js);
            var jsCode = System.IO.File.ReadAllText(path);
            _engine.Execute(jsCode);
        }

        /// <summary>
        /// 実行
        /// </summary>
        /// <param name="method"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public object Invoke(string method, params object[] parameters)
        {
            return _engine.Invoke(method, parameters).ToObject();
        }
    }
}
