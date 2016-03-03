using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using DevExpress.LookAndFeel;

namespace DcBatteryChoose
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            DevExpress.Skins.SkinManager.EnableFormSkins();
            UserLookAndFeel.Default.SetSkinStyle("Office 2010 Blue");//设置主题样式
            
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FrmMain());
        }
    }
}
