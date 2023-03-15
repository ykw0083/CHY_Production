using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FT_ADDON
{
    class Clipboard : STASafe
    {
        private Clipboard() { }

        public static void SetText(string text)
        {
            Clipboard clip = new Clipboard();
            clip.AddAction(() => System.Windows.Forms.Clipboard.SetText(text));
            clip.Execute();
        }
    }
}
