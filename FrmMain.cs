using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DevExpress.XtraTreeList;
using System.IO;
using DevExpress.XtraTreeList.Nodes;

namespace DcBatteryChoose
{
    public partial class FrmMain : XtraForm
    {
        LoadCollection loadCollection;
        BarChooseInfo barChooseInfo;
        public FrmMain()
        {
            InitializeComponent();
        }

        private void FrmMain_Load(object sender, EventArgs e)
        {
            this.KeyPreview = true;

            loadCollection = new LoadCollection();
            barChooseInfo = new BarChooseInfo();

            string[] strRoot = new string[4] { "动力回路合计", "控制负荷合计", "随机负荷合计","随机负荷合计" };

            string[] strChildren1 = new string[5] {"电厂事故照明","事故长明灯","燃机动力电源(应急润滑油泵/液压盘车泵)"
            ,"交流不停电电源装置","断路器合/跳闸"};
            string[] strChildren2 = new string[6] { "电气经常性负荷","试验保护屏","火灾报警控制箱"
                ,"天然气处理系统PLC","燃机组用控制电源正常负荷","燃机组用控制电源事故负荷"};
            string[] strChildren3 = new string[1] { "恢复供电断路器合闸" };

            for (int i = 0; i < strRoot.Length; i++)
            {
                ChargeLoadInfo loadInfo = new ChargeLoadInfo();
                loadInfo.ID = i + 1;
                loadInfo.ParentID = 0;
                loadInfo.Num = "";
                loadInfo.LoadName = strRoot[i];
                loadCollection.AddChargeLoadInfo(loadInfo);            
            }

            for (int i = 0; i < strChildren1.Length; i++)
            {
                ChargeLoadInfo loadInfo = new ChargeLoadInfo();
                loadInfo.ID = loadCollection.LstLoadInfo.Count + 1;
                loadInfo.ParentID = 1;
                loadInfo.LoadName = strChildren1[i];
                loadInfo.Num = (i + 1).ToString();

                loadCollection.AddChargeLoadInfo(loadInfo);  
            }

            for (int i = 0; i < strChildren2.Length; i++)
            {
                ChargeLoadInfo loadInfo = new ChargeLoadInfo();
                loadInfo.ID = loadCollection.LstLoadInfo.Count + 1;
                loadInfo.ParentID = 2;
                loadInfo.LoadName = strChildren2[i];
                loadInfo.Num = (i + 1).ToString();

                loadCollection.AddChargeLoadInfo(loadInfo);  
            }

            for (int i = 0; i < strChildren3.Length; i++)
            {
                ChargeLoadInfo loadInfo = new ChargeLoadInfo();
                loadInfo.ID = loadCollection.LstLoadInfo.Count + 1;
                loadInfo.ParentID = 3;
                loadInfo.LoadName = strChildren3[i];
                loadInfo.Num = (i + 1).ToString();

                loadCollection.AddChargeLoadInfo(loadInfo);  
            }
            /*
            ChargeLoadInfo powerLoad = new ChargeLoadInfo();
            powerLoad.ID = 1;
            powerLoad.ParentID = 0;
            powerLoad.LoadName = "动力回路合计";
            powerLoad.Num = "";
            loadCollection.AddChargeLoadInfo(powerLoad);

            ChargeLoadInfo ControlLoad = new ChargeLoadInfo();
            ControlLoad.ID = 2;
            ControlLoad.ParentID = 0;
            ControlLoad.LoadName = "控制负荷合计";
            ControlLoad.Num = "";
            loadCollection.AddChargeLoadInfo(ControlLoad);

            ChargeLoadInfo randomLoad = new ChargeLoadInfo();
            randomLoad.ID = 3;
            randomLoad.ParentID = 0;
            randomLoad.LoadName = "随机负荷合计";
            randomLoad.Num = "";
            loadCollection.AddChargeLoadInfo(randomLoad);

            ChargeLoadInfo totLoad = new ChargeLoadInfo();
            totLoad.ID = 4;
            totLoad.ParentID = 0;
            totLoad.LoadName = "合计";
            totLoad.Num = "";
            loadCollection.AddChargeLoadInfo(totLoad);


            ChargeLoadInfo childLoad = new ChargeLoadInfo();
            childLoad.ID = 5;
            childLoad.ParentID = 1;
            childLoad.LoadName = "电厂事故照明";
            childLoad.Num = "1";
            loadCollection.AddChargeLoadInfo(childLoad);

            ChargeLoadInfo childLoad1 = new ChargeLoadInfo();
            childLoad1.ID = 6;
            childLoad1.ParentID = 1;
            childLoad1.LoadName = "事故长明灯";
            childLoad1.Num = "2";
            loadCollection.AddChargeLoadInfo(childLoad1);

            ChargeLoadInfo childLoad2 = new ChargeLoadInfo();
            childLoad2.ID = 7;
            childLoad2.ParentID = 1;
            childLoad2.LoadName = "燃机动力电源(应急润滑油泵/液压盘车泵)";
            childLoad2.Num = "3";
            loadCollection.AddChargeLoadInfo(childLoad2);

            ChargeLoadInfo childLoad4 = new ChargeLoadInfo();
            childLoad4.ID = 8;
            childLoad4.ParentID = 1;
            childLoad4.LoadName = "交流不停电电源装置";
            childLoad4.Num = "5";
            loadCollection.AddChargeLoadInfo(childLoad4);

            ChargeLoadInfo childLoad5 = new ChargeLoadInfo();
            childLoad5.ID = 9;
            childLoad5.ParentID = 1;
            childLoad5.LoadName = "断路器合/跳闸";
            childLoad5.Num = "6";
            loadCollection.AddChargeLoadInfo(childLoad5);

            ChargeLoadInfo childLoad6 = new ChargeLoadInfo();
            childLoad6.ID = 10;
            childLoad6.ParentID = 2;
            childLoad6.LoadName = "电气经常性负荷";
            childLoad6.Num = "1";
            loadCollection.AddChargeLoadInfo(childLoad6);


            ChargeLoadInfo childLoad7 = new ChargeLoadInfo();
            childLoad7.ID = 11;
            childLoad7.ParentID = 2;
            childLoad7.LoadName = "试验保护屏";
            childLoad7.Num = "2";
            loadCollection.AddChargeLoadInfo(childLoad7);

            ChargeLoadInfo childLoad8 = new ChargeLoadInfo();
            childLoad8.ID = 12;
            childLoad8.ParentID = 2;
            childLoad8.LoadName = "火灾报警控制箱";
            childLoad8.Num = "3";
            loadCollection.AddChargeLoadInfo(childLoad8);

            ChargeLoadInfo childLoad7 = new ChargeLoadInfo();
            childLoad7.ID = 11;
            childLoad7.ParentID = 2;
            childLoad7.LoadName = "试验保护屏";
            childLoad7.Num = "2";
            loadCollection.AddChargeLoadInfo(childLoad7);

            */
            treeList.DataSource = loadCollection.LstLoadInfo;
        }

        private void treeList_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Insert)
            {
                TreeList tree = sender as TreeList;
                int nodeId = tree.FocusedNode.Id;
                if (nodeId != 3)
                {
                    TreeListNode parentNode = tree.FocusedNode;
                    String num = "1";
                    if (nodeId > 3)
                        parentNode = parentNode.ParentNode;
                    if(parentNode.Nodes.Count > 0)
                    {
                        String lastNum = parentNode.LastNode.GetDisplayText("Num");
                        num = (MyMethod.str2int(lastNum) + 1).ToString();
                    }

                    tree.AppendNode(new object[] { num }, parentNode.Id);
                }
            }
            else if (e.KeyCode == Keys.Delete)
            { 
                TreeList tree = sender as TreeList;
                int nodeId = tree.FocusedNode.Id;
                if (nodeId > 3)
                {
                    tree.DeleteNode(tree.FocusedNode);
                }
            }
        }

        private List<ChargeLoadInfo> sortLoads()
        {
            List<ChargeLoadInfo> sortList = new List<ChargeLoadInfo>();
            foreach(ChargeLoadInfo load in loadCollection.LstLoadInfo)
            {
                int i = 0;
                while (i < sortList.Count)
                {
                    if (sortList[i].ID == load.ParentID)
                        break;
                    else if (sortList[i].ParentID == load.ParentID 
                        && (sortList[i].Num.Length > load.Num.Length || string.Compare(sortList[i].Num ,load.Num) > 0) )
                        break;
                    i++;
                }
                sortList.Insert(i, load);
            
            }
            return sortList;
        
        }

        private void FrmMain_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.S && e.Control)
            {
                barChooseInfo.All = MyMethod.str2f(tbAll.Text);
                barChooseInfo.AllFinish = MyMethod.str2f(tbAllFinish.Text);
                barChooseInfo.AllFinishMin = MyMethod.str2f(tbAllFinishMin.Text);
                barChooseInfo.AllMax = MyMethod.str2f(tbAllMax.Text);
                barChooseInfo.CalCap1 = MyMethod.str2f(tbCalCap1.Text);
                barChooseInfo.CalCap2 = MyMethod.str2f(tbCalCap2.Text);
                barChooseInfo.CalCap3 = MyMethod.str2f(tbCalCap3.Text);
                barChooseInfo.CalCap4 = MyMethod.str2f(tbCalCap4.Text);
                barChooseInfo.CalNum = MyMethod.str2f(tbCalNum.Text);
                barChooseInfo.CapRate1 = MyMethod.str2f(tbCapRate1.Text);

                barChooseInfo.CapRate21 = MyMethod.str2f(tbCapRate21.Text);
                barChooseInfo.CapRate22 = MyMethod.str2f(tbCapRate22.Text);
                barChooseInfo.CapRate31 = MyMethod.str2f(tbCapRate31.Text);
                barChooseInfo.CapRate32 = MyMethod.str2f(tbCapRate32.Text);
                barChooseInfo.CapRate33 = MyMethod.str2f(tbCapRate33.Text);
                barChooseInfo.CapRate4 = MyMethod.str2f(tbCapRate4.Text);
                barChooseInfo.ChooseCap = MyMethod.str2f(tbChooseCap.Text);
                barChooseInfo.ChooseNum = MyMethod.str2int(tbChooseNum.Text);
                barChooseInfo.Ctrl = MyMethod.str2f(tbCtrl.Text);
                barChooseInfo.CtrlFinish = MyMethod.str2f(tbCtrlFinish.Text);

                barChooseInfo.CtrlFinishMin = MyMethod.str2f(tbCtrlFinishMin.Text);
                barChooseInfo.CtrlMax = MyMethod.str2f(tbCtrlMax.Text);
                barChooseInfo.I1 = MyMethod.str2f(tbI1.Text);
                barChooseInfo.I21 = MyMethod.str2f(tbI21.Text);
                barChooseInfo.I22 = MyMethod.str2f(tbI22.Text);
                barChooseInfo.I31 = MyMethod.str2f(tbI31.Text);
                barChooseInfo.I32 = MyMethod.str2f(tbI32.Text);
                barChooseInfo.I33 = MyMethod.str2f(tbI33.Text);
                barChooseInfo.I4 = MyMethod.str2f(tbI4.Text);
                barChooseInfo.KK1 = MyMethod.str2f(tbKk1.Text);

                barChooseInfo.KK2 = MyMethod.str2f(tbKk2.Text);
                barChooseInfo.KK3 = MyMethod.str2f(tbKk3.Text);
                barChooseInfo.KK4 = MyMethod.str2f(tbKk4.Text);
                barChooseInfo.MaxCap = MyMethod.str2f(tbMaxCap.Text);
                barChooseInfo.Power = MyMethod.str2f(tbPower.Text);
                barChooseInfo.PowerFinish = MyMethod.str2f(tbPowerFinish.Text);
                barChooseInfo.PowerFinishMin = MyMethod.str2f(tbPowerFinishMin.Text);
                barChooseInfo.PowerMax = MyMethod.str2f(tbPowerMax.Text);
                barChooseInfo.SingelVol = MyMethod.str2f(tbSingleVol.Text);
                barChooseInfo.StdVol = MyMethod.str2f(tbSysVol.Text);                
                
                loadCollection.StdVol = MyMethod.str2f(tbStdVol.EditValue.ToString());
                loadCollection.BarNum = MyMethod.str2int(tbBarNum.Text);
                byte[] template = Properties.Resources.model;//Excel资源去掉后缀名
                FileStream stream = new FileStream("a.tmp", FileMode.Create);
                stream.Write(template, 0, template.Length);
                stream.Close();
                stream.Dispose();

                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "xls files(*.xls)|*.xls";
                sfd.FileName = "直流系统蓄电池选择";
                sfd.AddExtension = true;
                sfd.RestoreDirectory = true;
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    ExcelHelper.SaveExcel(System.AppDomain.CurrentDomain.BaseDirectory + "\\a.tmp",sfd.FileName , loadCollection, barChooseInfo ,sortLoads());
                }
                File.Delete("a.tmp");
            }
        }

        private void treeList_ShowingEditor(object sender, CancelEventArgs e)
        {
            TreeList tree = sender as TreeList;
            if (tree.FocusedNode.Level == 0)
            {
                e.Cancel = true;            
            }
        }

        private void treeList_CellValueChanged(object sender, CellValueChangedEventArgs e)
        {
            int column = e.Column.ColumnHandle;
            if (column == 3 || column == 4)
            {
                TreeList tree = sender as TreeList;
                TreeListNode node = tree.FocusedNode;
                float calCap = MyMethod.obj2f(node.GetValue(3));
                float rate = MyMethod.obj2f(node.GetValue(4));
                node.SetValue(6, calCap * rate);
                float std_vol = MyMethod.str2f(tbStdVol.Text);
                node.SetValue(7, calCap * rate * 1000 / std_vol);            
            }
            else if (column > 7)
            {
                TreeList tree = sender as TreeList;
                TreeListNode parentNode = tree.FocusedNode.ParentNode;
                float val = 0;
                foreach (TreeListNode child in parentNode.Nodes)
                { 
                    val += MyMethod.obj2f(child.GetValue(column));
                }
                parentNode.SetValue(column , val);

                float powerTot = MyMethod.obj2f(tree.FindNodeByID(0).GetValue(column));
                float ctrlTot = MyMethod.obj2f(tree.FindNodeByID(1).GetValue(column));                
                float randomTot = MyMethod.obj2f(tree.FindNodeByID(2).GetValue(column));

                float tot;
                int num = MyMethod.str2int(tbBarNum.Text);
                if(num == 1)
                    tot = powerTot + ctrlTot + randomTot ;
                else
                    tot = 0.6f * powerTot + ctrlTot + randomTot ;
                tree.FindNodeByID(3).SetValue(column , tot);

                if (column == 9) //j40
                {
                    tbI1.Text = tot.ToString("f3");
                    tbI21.Text = tot.ToString("f3");
                    tbI31.Text = tot.ToString("f3");
                    calc("tbI1");
                }
                else if (column == 10) //k40
                {
                    tbI22.Text = tot.ToString("f3");
                    tbI32.Text = tot.ToString("f3");
                    calc("tbI2");
                }
                else if (column == 11) //l40
                {
                    tbI33.Text = tot.ToString("f3");
                    calc("tbI3");
                }
                else if (column == 15) //P40
                {
                    tbI4.Text = tot.ToString("f3");
                    calc("tbI4");
                }
            }
           

        }

        private void tb_num_EditValueChanged(object sender, EventArgs e)
        {
            float stdVol = MyMethod.str2f(tbSysVol.Text);

            float singleVol = MyMethod.str2f(tbSingleVol.Text);
            tbCalNum.Text = (stdVol * 1.05 / singleVol).ToString("f2");


            float numChoose = MyMethod.str2f(tbChooseNum.Text);

            tbCtrlMax.Text = (1.1 * stdVol / numChoose).ToString("f2");
            tbPowerMax.Text = (1.125 * stdVol / numChoose).ToString("f2");
            tbAllMax.Text = (1.1 * stdVol / numChoose).ToString("f2");

            tbCtrlFinishMin.Text = (0.85 * stdVol / numChoose).ToString("f2");
            tbPowerFinishMin.Text = (0.875 * stdVol / numChoose).ToString("f2");
            tbAllFinishMin.Text = (0.875 * stdVol / numChoose).ToString("f2");

        }

        private void tb_stage_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit tb = sender as TextEdit;
            calc(tb.Name);
        }
        private void calc(string name)
        {
            if (name == "tbCapRate1" || name == "tbI1")
            {
                float kk1 = MyMethod.str2f(tbKk1.Text);
                float capRate1 = MyMethod.str2f(tbCapRate1.Text);
                float i1 = MyMethod.str2f(tbI1.Text);
                tbCalCap1.Text = (kk1 * i1 / capRate1).ToString("f0");
            }
            if (name == "tbCapRate21"
                || name == "tbCapRate22"
                || name == "tbI1"
                || name == "tbI2")
            {
                float kk2 = MyMethod.str2f(tbKk2.Text);
                float capRate21 = MyMethod.str2f(tbCapRate21.Text);
                float capRate22 = MyMethod.str2f(tbCapRate22.Text);
                float i21 = MyMethod.str2f(tbI21.Text);
                float i22 = MyMethod.str2f(tbI22.Text);
                tbCalCap2.Text = (kk2 * (i21 / capRate21 + (i22 - i21) / capRate22)).ToString();
            }
            if (name == "tbCapRate31"
                || name == "tbCapRate32"
                || name == "tbCapRate33"
                || name == "tbI1"
                || name == "tbI2"
                || name == "tbI3"
                )
            {
                float kk3 = MyMethod.str2f(tbKk3.Text);
                float capRate31 = MyMethod.str2f(tbCapRate31.Text);
                float capRate32 = MyMethod.str2f(tbCapRate32.Text);
                float capRate33 = MyMethod.str2f(tbCapRate33.Text);
                float i31 = MyMethod.str2f(tbI31.Text);
                float i32 = MyMethod.str2f(tbI32.Text);
                float i33 = MyMethod.str2f(tbI33.Text);
                tbCalCap3.Text = (kk3 * (i31 / capRate31 + (i32 - i31) / capRate32 + (i33 - i32) / capRate33)).ToString();
            }
            else if (name == "tbCapRate4" || name == "tbI4")
            {
                float kk4 = MyMethod.str2f(tbKk4.Text);
                float capRate4 = MyMethod.str2f(tbCapRate4.Text);
                float i4 = MyMethod.str2f(tbI4.Text);
                tbCalCap4.Text = (kk4 * i4 / capRate4).ToString("f0");
            }
            {
                float maxCap = MyMethod.str2f(tbCalCap1.Text);
                float calCap2 = MyMethod.str2f(tbCalCap2.Text);
                float calCap3 = MyMethod.str2f(tbCalCap3.Text);
                float calCap4 = MyMethod.str2f(tbCalCap4.Text);

                if(calCap2 + calCap4 > maxCap)
                {
                    maxCap = calCap2 + calCap4;
                }
                if( calCap3 + calCap4 > maxCap)
                {
                    maxCap = calCap3 + calCap4;
                }
                tbMaxCap.Text = maxCap.ToString();

            }

        }
    }
}
