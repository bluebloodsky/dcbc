using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DcBatteryChoose
{
    class ChargeLoadInfo
    {
        public string Num { get; set; }

        public string LoadName { get; set; }

        public string LoadDesc { get; set; }

        public string EqCap { get; set; }

        public string LoadRate { get; set; }

        public string ContinueTime { get; set; }

        public string CalCap { get; set; }

        public string CalCur { get; set; }

        public string Ijc { get; set; }

        public string I1 { get; set; }

        public string I2 { get; set; }

        public string I3 { get; set; }

        public string I4 { get; set; }

        public string I5 { get; set; }

        public string I6 { get; set; }

        public string Ichm { get; set; }

        public int ID { get; set; }

        public int ParentID { get; set; }
    }
    class LoadCollection
    {
        public double StdVol { get; set; }
        public int BarNum { get; set; } 
        public List<ChargeLoadInfo> LstLoadInfo{ get; set; }

        public void AddChargeLoadInfo(ChargeLoadInfo info)
        {
            if (null == LstLoadInfo)
            {
                LstLoadInfo = new List<ChargeLoadInfo>();
            }
            LstLoadInfo.Add(info);
        
        }
    }

    class BarChooseInfo
    {
        public double StdVol { get; set; }
        public double SingelVol { get; set; }
        public double CalNum { get; set; }
        public int ChooseNum { get; set; }

        public double CtrlMax { get; set; }
        public double PowerMax { get; set; }
        public double AllMax { get; set; }

        public double Ctrl { get; set; }
        public double Power { get; set; }
        public double All { get; set; }

        public double CtrlFinishMin { get; set; }
        public double PowerFinishMin { get; set; }
        public double AllFinishMin { get; set; }

        public double CtrlFinish { get; set; }
        public double PowerFinish { get; set; }
        public double AllFinish { get; set; }

        public double KK1 { get; set; }
        public double KK2 { get; set; }
        public double KK3 { get; set; }
        public double KK4 { get; set; }

        public double CapRate1 { get; set; }
        public double CapRate21 { get; set; }
        public double CapRate22 { get; set; }
        public double CapRate31 { get; set; }
        public double CapRate32 { get; set; }
        public double CapRate33 { get; set; }
        public double CapRate4 { get; set; }

        public double I1 { get; set; }

        public double I21 { get; set; }
        public double I22 { get; set; }
        public double I31 { get; set; }
        public double I32 { get; set; }
        public double I33 { get; set; }
        public double I4 { get; set; }

        public double CalCap1 { get; set; }
        public double CalCap2 { get; set; }
        public double CalCap3 { get; set; }
        public double CalCap4 { get; set; }

        public double MaxCap { get; set; }

        public double ChooseCap { get; set; }

    }
}
