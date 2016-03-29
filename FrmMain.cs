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
            test();
            DevExpress.XtraSplashScreen.SplashScreenManager.ShowForm(typeof(SplashScreen1));  
            InitializeComponent();          
        }

        private void FrmMain_Load(object sender, EventArgs e)
        {
            this.KeyPreview = true;

            loadCollection = new LoadCollection();
            barChooseInfo = new BarChooseInfo();

            string[] strRoot = new string[4] { "动力回路合计", "控制负荷合计", "随机负荷合计","合计" };

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
            treeList.DataSource = loadCollection.LstLoadInfo;

            DevExpress.XtraSplashScreen.SplashScreenManager.CloseForm();
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
                barChooseInfo.All = MyMethod.str2double(tbAll.Text);
                barChooseInfo.AllFinish = MyMethod.str2double(tbAllFinish.Text);
                barChooseInfo.AllFinishMin = MyMethod.str2double(tbAllFinishMin.Text);
                barChooseInfo.AllMax = MyMethod.str2double(tbAllMax.Text);
                barChooseInfo.CalCap1 = MyMethod.str2double(tbCalCap1.Text);
                barChooseInfo.CalCap2 = MyMethod.str2double(tbCalCap2.Text);
                barChooseInfo.CalCap3 = MyMethod.str2double(tbCalCap3.Text);
                barChooseInfo.CalCap4 = MyMethod.str2double(tbCalCap4.Text);
                barChooseInfo.CalNum = MyMethod.str2double(tbCalNum.Text);
                barChooseInfo.CapRate1 = MyMethod.str2double(tbCapRate1.Text);

                barChooseInfo.CapRate21 = MyMethod.str2double(tbCapRate21.Text);
                barChooseInfo.CapRate22 = MyMethod.str2double(tbCapRate22.Text);
                barChooseInfo.CapRate31 = MyMethod.str2double(tbCapRate31.Text);
                barChooseInfo.CapRate32 = MyMethod.str2double(tbCapRate32.Text);
                barChooseInfo.CapRate33 = MyMethod.str2double(tbCapRate33.Text);
                barChooseInfo.CapRate4 = MyMethod.str2double(tbCapRate4.Text);
                barChooseInfo.ChooseCap = MyMethod.str2double(tbChooseCap.Text);
                barChooseInfo.ChooseNum = MyMethod.str2int(tbChooseNum.Text);
                barChooseInfo.Ctrl = MyMethod.str2double(tbCtrl.Text);
                barChooseInfo.CtrlFinish = MyMethod.str2double(tbCtrlFinish.Text);

                barChooseInfo.CtrlFinishMin = MyMethod.str2double(tbCtrlFinishMin.Text);
                barChooseInfo.CtrlMax = MyMethod.str2double(tbCtrlMax.Text);
                barChooseInfo.I1 = MyMethod.str2double(tbI1.Text);
                barChooseInfo.I21 = MyMethod.str2double(tbI21.Text);
                barChooseInfo.I22 = MyMethod.str2double(tbI22.Text);
                barChooseInfo.I31 = MyMethod.str2double(tbI31.Text);
                barChooseInfo.I32 = MyMethod.str2double(tbI32.Text);
                barChooseInfo.I33 = MyMethod.str2double(tbI33.Text);
                barChooseInfo.I4 = MyMethod.str2double(tbI4.Text);
                barChooseInfo.KK1 = MyMethod.str2double(tbKk1.Text);

                barChooseInfo.KK2 = MyMethod.str2double(tbKk2.Text);
                barChooseInfo.KK3 = MyMethod.str2double(tbKk3.Text);
                barChooseInfo.KK4 = MyMethod.str2double(tbKk4.Text);
                barChooseInfo.MaxCap = MyMethod.str2double(tbMaxCap.Text);
                barChooseInfo.Power = MyMethod.str2double(tbPower.Text);
                barChooseInfo.PowerFinish = MyMethod.str2double(tbPowerFinish.Text);
                barChooseInfo.PowerFinishMin = MyMethod.str2double(tbPowerFinishMin.Text);
                barChooseInfo.PowerMax = MyMethod.str2double(tbPowerMax.Text);
                barChooseInfo.SingelVol = MyMethod.str2double(tbSingleVol.Text);
                barChooseInfo.StdVol = MyMethod.str2double(tbSysVol.Text);                
                
                loadCollection.StdVol = MyMethod.str2double(tbStdVol.EditValue.ToString());
                loadCollection.BarNum = MyMethod.str2int(tbBarNum.Text);

                byte[] template = Properties.Resources.model;//Excel资源去掉后缀名

                Stream stream = new MemoryStream(template);

                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "xls files(*.xls)|*.xls";
                sfd.FileName = "直流系统蓄电池选择";
                sfd.AddExtension = true;
                sfd.RestoreDirectory = true;
                if (sfd.ShowDialog() == DialogResult.OK)
                {

                    ExcelNPIOHelper helper = new ExcelNPIOHelper(sfd.FileName);
                    helper.ExcelUpdate(stream, loadCollection, barChooseInfo, sortLoads());
                    helper.Dispose();
                }

                /*
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
                 * */
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
                double calCap = MyMethod.obj2double(node.GetValue(3));
                double rate = MyMethod.obj2double(node.GetValue(4));
                node.SetValue(6, (calCap * rate).ToString("0.000#"));
                double std_vol = MyMethod.str2double(tbStdVol.Text);
                node.SetValue(7, (calCap * rate * 1000 / std_vol).ToString("f4"));            
            }
            else if (column > 7)
            {
                TreeList tree = sender as TreeList;
                TreeListNode parentNode = tree.FocusedNode.ParentNode;
                double val = 0;
                foreach (TreeListNode child in parentNode.Nodes)
                { 
                    val += MyMethod.obj2double(child.GetValue(column));
                }
                parentNode.SetValue(column , val);

                double powerTot = MyMethod.obj2double(tree.FindNodeByID(0).GetValue(column));
                double ctrlTot = MyMethod.obj2double(tree.FindNodeByID(1).GetValue(column));                
                double randomTot = MyMethod.obj2double(tree.FindNodeByID(2).GetValue(column));

                double tot;
                int num = MyMethod.str2int(tbBarNum.Text);
                if(num == 1)
                    tot = powerTot + ctrlTot + randomTot ;
                else
                    tot = 0.6f * powerTot + ctrlTot + randomTot ;
                tree.FindNodeByID(3).SetValue(column, tot.ToString("f4"));

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
            double stdVol = MyMethod.str2double(tbSysVol.Text);

            double singleVol = MyMethod.str2double(tbSingleVol.Text);
            tbCalNum.Text = (stdVol * 1.05 / singleVol).ToString("f2");


            double numChoose = MyMethod.str2double(tbChooseNum.Text);

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
                double kk1 = MyMethod.str2double(tbKk1.Text);
                double capRate1 = MyMethod.str2double(tbCapRate1.Text);
                double i1 = MyMethod.str2double(tbI1.Text);
                tbCalCap1.Text = (kk1 * i1 / capRate1).ToString("f0");
            }
            if (name == "tbCapRate21"
                || name == "tbCapRate22"
                || name == "tbI1"
                || name == "tbI2")
            {
                double kk2 = MyMethod.str2double(tbKk2.Text);
                double capRate21 = MyMethod.str2double(tbCapRate21.Text);
                double capRate22 = MyMethod.str2double(tbCapRate22.Text);
                double i21 = MyMethod.str2double(tbI21.Text);
                double i22 = MyMethod.str2double(tbI22.Text);
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
                double kk3 = MyMethod.str2double(tbKk3.Text);
                double capRate31 = MyMethod.str2double(tbCapRate31.Text);
                double capRate32 = MyMethod.str2double(tbCapRate32.Text);
                double capRate33 = MyMethod.str2double(tbCapRate33.Text);
                double i31 = MyMethod.str2double(tbI31.Text);
                double i32 = MyMethod.str2double(tbI32.Text);
                double i33 = MyMethod.str2double(tbI33.Text);
                tbCalCap3.Text = (kk3 * (i31 / capRate31 + (i32 - i31) / capRate32 + (i33 - i32) / capRate33)).ToString();
            }
            else if (name == "tbCapRate4" || name == "tbI4")
            {
                double kk4 = MyMethod.str2double(tbKk4.Text);
                double capRate4 = MyMethod.str2double(tbCapRate4.Text);
                double i4 = MyMethod.str2double(tbI4.Text);
                tbCalCap4.Text = (kk4 * i4 / capRate4).ToString("f0");
            }
            {
                double maxCap = MyMethod.str2double(tbCalCap1.Text);
                double calCap2 = MyMethod.str2double(tbCalCap2.Text);
                double calCap3 = MyMethod.str2double(tbCalCap3.Text);
                double calCap4 = MyMethod.str2double(tbCalCap4.Text);

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
        private void test()
        {
            if (!RegisterHelper.IsRegeditItemExist("wjz_dcbc"))
            {
                RegisterHelper.createReg("wjz_dcbc");
            }
            else
            {
                DateTime startDate, endDate;
                RegisterHelper.ReadReg("wjz_dcbc", out startDate, out endDate);
                if (startDate > DateTime.Now || endDate < DateTime.Now || startDate.AddDays(RegisterHelper.useDays) < endDate)
                {
                    XtraMessageBox.Show("软件已过期，请联系开发人员获取正式版本", "提示");
                    System.Environment.Exit(0);
                }
            }
        }

        private void treeList_ValidatingEditor(object sender, DevExpress.XtraEditors.Controls.BaseContainerValidateEditorEventArgs e)
        {
            TreeList tree = sender as TreeList;
            int column = tree.FocusedColumn.ColumnHandle;

            if (column > 2)
            { 
                 if (e.Value != null
                     && !System.Text.RegularExpressions.Regex.IsMatch(e.Value.ToString(), @"^[+-]?\d*[.]?\d*$"))//e.Value就是你编辑后的值，，自己写验证。。。
                {
                    e.ErrorText = "只能输入浮点数！";
                    e.Valid = false;
                    return;
                }
            
            }
        }
    }
}
