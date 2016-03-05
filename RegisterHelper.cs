using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;
namespace DcBatteryChoose
{
    class RegisterHelper
    {
        public const int useDays = 30;
        public static void createReg(string regName)
        {
            RegistryKey key = Registry.CurrentUser;
            RegistryKey software = key.CreateSubKey("SOFTWARE\\" + regName);
            software.SetValue("startDate", DateTime.Today.ToShortDateString());
            software.SetValue("endDate", DateTime.Today.AddDays(useDays).ToShortDateString());
            key.Close();
        }

        public static bool IsRegeditItemExist(string regName)
        {
            string[] subkeyNames;
            RegistryKey hkml = Registry.CurrentUser;
            RegistryKey software = hkml.OpenSubKey("SOFTWARE");
            //RegistryKey software = hkml.OpenSubKey("SOFTWARE", true);  
            subkeyNames = software.GetSubKeyNames();
            //取得该项下所有子项的名称的序列，并传递给预定的数组中  
            foreach (string keyName in subkeyNames)
            //遍历整个数组  
            {
                if (keyName == regName)
                //判断子项的名称  
                {
                    hkml.Close();
                    return true;
                }
            }
            hkml.Close();
            return false;
        }

        public static void ReadReg(string regName, out DateTime startDate, out DateTime endDate)
        {
            RegistryKey key = Registry.CurrentUser.OpenSubKey("software\\" + regName);
            startDate = DateTime.Parse(key.GetValue("startDate").ToString());
            endDate = DateTime.Parse(key.GetValue("endDate").ToString());
            key.Close();
        }
    }
}
