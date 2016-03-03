using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Excel = Microsoft.Office.Interop.Excel;

namespace DcBatteryChoose
{
    class ExcelHelper
    {
        public static void SaveExcel(string originName, string destName, LoadCollection load, BarChooseInfo barChooseInfo, List<ChargeLoadInfo> lstLoadInfo)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory;
            Excel.Application application = new Excel.Application();
            Excel.Workbook workBook = application.Workbooks.Open(originName);
            
            Excel.Worksheet sheet = workBook.Sheets[1];
            Excel.Worksheet sheet2 = workBook.Sheets[2];
            try
            {
            sheet.Cells[1, 4] = load.StdVol;
            sheet.Cells[1, 7] = load.BarNum;

            for (int i = 0; i < lstLoadInfo.Count; i++)
            {
                ChargeLoadInfo info = lstLoadInfo[i];
                sheet.Cells[6 + i, 1] = info.Num;
                sheet.Cells[6 + i, 2] = info.LoadName;
                sheet.Cells[6 + i, 3] = info.LoadDesc;
                sheet.Cells[6 + i, 4] = info.EqCap;
                sheet.Cells[6 + i, 5] = info.LoadRate;
                sheet.Cells[6 + i, 6] = info.ContinueTime;
                sheet.Cells[6 + i, 7] = info.CalCap;
                sheet.Cells[6 + i, 8] = info.CalCur;
                sheet.Cells[6 + i, 9] = info.Ijc;
                sheet.Cells[6 + i, 10] = info.I1;
                sheet.Cells[6 + i, 11] = info.I2;
                sheet.Cells[6 + i, 12] = info.I3;
                sheet.Cells[6 + i, 13] = info.I4;
                sheet.Cells[6 + i, 14] = info.I5;
                sheet.Cells[6 + i, 15] = info.I6;
                sheet.Cells[6 + i, 16] = info.Ichm;
            }

            sheet2.Cells[3, 2] = barChooseInfo.StdVol;
            sheet2.Cells[3, 4] = barChooseInfo.SingelVol;
            sheet2.Cells[4, 2] = barChooseInfo.CalNum;
            sheet2.Cells[4, 4] = barChooseInfo.ChooseNum;

            sheet2.Cells[6, 3] = barChooseInfo.CtrlMax;
            sheet2.Cells[7, 3] = barChooseInfo.PowerMax;
            sheet2.Cells[8, 3] = barChooseInfo.AllMax;

            sheet2.Cells[6, 5] = barChooseInfo.Ctrl;
            sheet2.Cells[7, 5] = barChooseInfo.Power;
            sheet2.Cells[8, 5] = barChooseInfo.All;

            sheet2.Cells[10, 3] = barChooseInfo.CtrlFinishMin;
            sheet2.Cells[11, 3] = barChooseInfo.PowerFinishMin;
            sheet2.Cells[12, 3] = barChooseInfo.AllFinishMin;

            sheet2.Cells[10, 5] = barChooseInfo.CtrlFinish;
            sheet2.Cells[11, 5] = barChooseInfo.PowerFinish;
            sheet2.Cells[12, 5] = barChooseInfo.AllFinish;

            sheet2.Cells[15, 2] = barChooseInfo.KK1.ToString("F1");
            sheet2.Cells[15, 4] = barChooseInfo.KK2.ToString("F1");
            sheet2.Cells[15, 7] = barChooseInfo.KK3.ToString("F1");
            sheet2.Cells[15, 11] = barChooseInfo.KK4.ToString("F1");

            sheet2.Cells[17, 2] = barChooseInfo.CapRate1;
            sheet2.Cells[17, 4] = barChooseInfo.CapRate21;
            sheet2.Cells[17, 5] = barChooseInfo.CapRate22;
            sheet2.Cells[17, 7] = barChooseInfo.CapRate31;
            sheet2.Cells[17, 8] = barChooseInfo.CapRate32;
            sheet2.Cells[17, 9] = barChooseInfo.CapRate33;
            sheet2.Cells[17, 11] = barChooseInfo.CapRate4;

            sheet2.Cells[19, 2] = barChooseInfo.I1;
            sheet2.Cells[19, 4] = barChooseInfo.I21;
            sheet2.Cells[19, 5] = barChooseInfo.I22;
            sheet2.Cells[19, 7] = barChooseInfo.I31;
            sheet2.Cells[19, 8] = barChooseInfo.I32;
            sheet2.Cells[19, 9] = barChooseInfo.I33;
            sheet2.Cells[19, 11] = barChooseInfo.I4;

            sheet2.Cells[20, 2] = barChooseInfo.CalCap1;
            sheet2.Cells[20, 4] = barChooseInfo.CalCap2;
            sheet2.Cells[20, 7] = barChooseInfo.CalCap3;
            sheet2.Cells[20, 11] = barChooseInfo.CalCap4;

            sheet2.Cells[21, 2] = barChooseInfo.MaxCap;
            sheet2.Cells[22, 2] = barChooseInfo.ChooseCap;

            workBook.Saved = true;

                workBook.SaveAs(destName);
            }
            catch
            {
            }
            workBook.Close();
            application.Quit();
            PublicMethod.Kill(application);//调用kill当前excel进程  

            releaseObject(sheet);
            releaseObject(sheet2);
            releaseObject(workBook);

            releaseObject(application);
        }
        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }

    }
    public class PublicMethod
    {
        [System.Runtime.InteropServices.DllImport("User32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);

        public static void Kill(Microsoft.Office.Interop.Excel.Application excel)
        {
            try
            {
                IntPtr t = new IntPtr(excel.Hwnd);

                int k = 0;

                GetWindowThreadProcessId(t, out k);

                System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);

                p.Kill();
            }
            catch
            { }
        }
    }
}
