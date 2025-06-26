using MyTemplate.Report.Models;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

namespace MyTemplate.Report.ViewModels
{
    public class PersonViewModel
    {
        public PersonModel Person { get; set; }

        public ItemModule Item { get; set; } // 不備内容のモジュール
    }
}
