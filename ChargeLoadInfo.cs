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
        public float StdVol { get; set; }
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
        public float StdVol { get; set; }
        public float SingelVol { get; set; }
        public float CalNum { get; set; }
        public int ChooseNum { get; set; }

        public float CtrlMax { get; set; }
        public float PowerMax { get; set; }
        public float AllMax { get; set; }

        public float Ctrl { get; set; }
        public float Power { get; set; }
        public float All { get; set; }

        public float CtrlFinishMin { get; set; }
        public float PowerFinishMin { get; set; }
        public float AllFinishMin { get; set; }

        public float CtrlFinish { get; set; }
        public float PowerFinish { get; set; }
        public float AllFinish { get; set; }

        public float KK1 { get; set; }
        public float KK2 { get; set; }
        public float KK3 { get; set; }
        public float KK4 { get; set; }

        public float CapRate1 { get; set; }
        public float CapRate21 { get; set; }
        public float CapRate22 { get; set; }
        public float CapRate31 { get; set; }
        public float CapRate32 { get; set; }
        public float CapRate33 { get; set; }
        public float CapRate4 { get; set; }

        public float I1 { get; set; }

        public float I21 { get; set; }
        public float I22 { get; set; }
        public float I31 { get; set; }
        public float I32 { get; set; }
        public float I33 { get; set; }
        public float I4 { get; set; }

        public float CalCap1 { get; set; }
        public float CalCap2 { get; set; }
        public float CalCap3 { get; set; }
        public float CalCap4 { get; set; }

        public float MaxCap { get; set; }

        public float ChooseCap { get; set; }

    }
}
