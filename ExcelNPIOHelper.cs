using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;
using System.IO;
using System.Data;
using NPOI.HSSF.UserModel;

namespace DcBatteryChoose
{
    class ExcelNPIOHelper:IDisposable
    {
        private string fileName = null; //文件名
        private IWorkbook workbook = null;
        private FileStream fs = null;
        private bool disposed;

        public ExcelNPIOHelper(string fileName)
        {
            this.fileName = fileName;
            disposed = false;
        }

        /// <summary>
        /// 将DataTable数据导入到excel中
        /// </summary>
        /// <param name="data">要导入的数据</param>
        /// <param name="isColumnWritten">DataTable的列名是否要导入</param>
        /// <param name="sheetName">要导入的excel的sheet的名称</param>
        /// <returns>导入数据行数(包含列名那一行)</returns>
        public int DataTableToExcel(DataTable data, string sheetName, bool isColumnWritten)
        {
            int i = 0;
            int j = 0;
            int count = 0;
            ISheet sheet = null;

            fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            workbook = new HSSFWorkbook();

            try
            {
                if (workbook != null)
                {
                    sheet = workbook.CreateSheet(sheetName);
                }
                else
                {
                    return -1;
                }

                if (isColumnWritten == true) //写入DataTable的列名
                {
                    IRow row = sheet.CreateRow(0);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Columns[j].ColumnName);
                    }
                    count = 1;
                }
                else
                {
                    count = 0;
                }

                for (i = 0; i < data.Rows.Count; ++i)
                {
                    IRow row = sheet.CreateRow(count);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Rows[i][j].ToString());
                    }
                    ++count;
                }
                workbook.Write(fs); //写入到excel
                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return -1;
            }
        }

        /// <summary>
        /// 将excel中的数据导入到DataTable中
        /// </summary>
        /// <param name="sheetName">excel工作薄sheet的名称</param>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
        /// <returns>返回的DataTable</returns>
        public DataTable ExcelToDataTable(string sheetName, bool isFirstRowColumn)
        {
            ISheet sheet = null;
            DataTable data = new DataTable();
            int startRow = 0;
            try
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                workbook = new HSSFWorkbook(fs);

                if (sheetName != null)
                {
                    sheet = workbook.GetSheet(sheetName);
                    if (sheet == null) //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                    {
                        sheet = workbook.GetSheetAt(0);
                    }
                }
                else
                {
                    sheet = workbook.GetSheetAt(0);
                }
                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(0);
                    int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数

                    if (isFirstRowColumn)
                    {
                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                        {
                            ICell cell = firstRow.GetCell(i);
                            if (cell != null)
                            {
                                string cellValue = cell.StringCellValue;
                                if (cellValue != null)
                                {
                                    DataColumn column = new DataColumn(cellValue);
                                    data.Columns.Add(column);
                                }
                            }
                        }
                        startRow = sheet.FirstRowNum + 1;
                    }
                    else
                    {
                        startRow = sheet.FirstRowNum;
                    }

                    //最后一列的标号
                    int rowCount = sheet.LastRowNum;
                    for (int i = startRow; i <= rowCount; ++i)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue; //没有数据的行默认是null　　　　　　　
                        
                        DataRow dataRow = data.NewRow();
                        for (int j = row.FirstCellNum; j < cellCount; ++j)
                        {
                            if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                                dataRow[j] = row.GetCell(j).ToString();
                        }
                        data.Rows.Add(dataRow);
                    }
                }

                return data;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return null;
            }
        }

        public bool ExcelUpdate(Stream stream , LoadCollection load, BarChooseInfo barChooseInfo, List<ChargeLoadInfo> lstLoadInfo)
        {
            workbook = new HSSFWorkbook(stream);
            ISheet sheet = workbook.GetSheetAt(0);
            IRow row11 = sheet.GetRow(0);
            row11.GetCell(3).SetCellValue(load.StdVol);
            row11.GetCell(6).SetCellValue(load.BarNum);

            IRow rowTemplate = sheet.GetRow(5);
            IRow rowBold = sheet.GetRow(6);
            for (int i = 0; i < lstLoadInfo.Count; i++)
            {                
                if (lstLoadInfo[i].ParentID == 0)
                {
                    rowBold.CopyRowTo(i + 7);
                }
                else
                {
                    rowTemplate.CopyRowTo(i + 7);
                }
                IRow row = sheet.GetRow(i + 7);
                row.GetCell(0).SetCellValue(lstLoadInfo[i].Num);
                row.GetCell(1).SetCellValue(lstLoadInfo[i].LoadName);
                row.GetCell(2).SetCellValue(lstLoadInfo[i].LoadDesc);
                row.GetCell(3).SetCellValue(lstLoadInfo[i].EqCap);
                row.GetCell(4).SetCellValue(lstLoadInfo[i].LoadRate);
                row.GetCell(5).SetCellValue(lstLoadInfo[i].ContinueTime);
                row.GetCell(6).SetCellValue(lstLoadInfo[i].CalCap);
                row.GetCell(7).SetCellValue(lstLoadInfo[i].CalCur);
                row.GetCell(8).SetCellValue(lstLoadInfo[i].Ijc);
                row.GetCell(9).SetCellValue(lstLoadInfo[i].I1);
                row.GetCell(10).SetCellValue(lstLoadInfo[i].I2);
                row.GetCell(11).SetCellValue(lstLoadInfo[i].I3);
                row.GetCell(12).SetCellValue(lstLoadInfo[i].I4);
                row.GetCell(13).SetCellValue(lstLoadInfo[i].I5);
                row.GetCell(14).SetCellValue(lstLoadInfo[i].I6);
                row.GetCell(15).SetCellValue(lstLoadInfo[i].Ichm);
            }

            sheet.ShiftRows(7, sheet.LastRowNum , -2);

            ISheet sheet2 = workbook.GetSheetAt(1);

            sheet2.GetRow(2).GetCell(1).SetCellValue(barChooseInfo.StdVol);
            sheet2.GetRow(2).GetCell(3).SetCellValue(barChooseInfo.SingelVol);
            sheet2.GetRow(3).GetCell(1).SetCellValue(barChooseInfo.CalNum);
            sheet2.GetRow(3).GetCell(3).SetCellValue(barChooseInfo.ChooseNum);

            sheet2.GetRow(5).GetCell(2).SetCellValue(barChooseInfo.CtrlMax);
            sheet2.GetRow(6).GetCell(2).SetCellValue(barChooseInfo.PowerMax);
            sheet2.GetRow(7).GetCell(2).SetCellValue(barChooseInfo.AllMax);
            sheet2.GetRow(7).GetCell(4).SetCellValue(barChooseInfo.Ctrl);
            sheet2.GetRow(6).GetCell(4).SetCellValue(barChooseInfo.Power);
            sheet2.GetRow(7).GetCell(4).SetCellValue(barChooseInfo.All);

            sheet2.GetRow(9).GetCell(2).SetCellValue(barChooseInfo.CtrlFinishMin);
            sheet2.GetRow(10).GetCell(2).SetCellValue(barChooseInfo.PowerFinishMin);
            sheet2.GetRow(11).GetCell(2).SetCellValue(barChooseInfo.AllFinishMin);
            sheet2.GetRow(9).GetCell(4).SetCellValue(barChooseInfo.CtrlFinish);
            sheet2.GetRow(10).GetCell(4).SetCellValue(barChooseInfo.PowerFinish);
            sheet2.GetRow(11).GetCell(4).SetCellValue(barChooseInfo.AllFinish);

            sheet2.GetRow(14).GetCell(1).SetCellValue(barChooseInfo.KK1);
            sheet2.GetRow(14).GetCell(3).SetCellValue(barChooseInfo.KK2);
            sheet2.GetRow(14).GetCell(6).SetCellValue(barChooseInfo.KK3);
            sheet2.GetRow(14).GetCell(10).SetCellValue(barChooseInfo.KK4);
            sheet2.GetRow(16).GetCell(1).SetCellValue(barChooseInfo.CapRate1);
            sheet2.GetRow(16).GetCell(3).SetCellValue(barChooseInfo.CapRate21);
            sheet2.GetRow(16).GetCell(4).SetCellValue(barChooseInfo.CapRate22);
            sheet2.GetRow(16).GetCell(6).SetCellValue(barChooseInfo.CapRate31);
            sheet2.GetRow(16).GetCell(7).SetCellValue(barChooseInfo.CapRate32);
            sheet2.GetRow(16).GetCell(8).SetCellValue(barChooseInfo.CapRate33);
            sheet2.GetRow(16).GetCell(10).SetCellValue(barChooseInfo.CapRate4);

            sheet2.GetRow(18).GetCell(1).SetCellValue(barChooseInfo.I1);
            sheet2.GetRow(18).GetCell(3).SetCellValue(barChooseInfo.I21);
            sheet2.GetRow(18).GetCell(4).SetCellValue(barChooseInfo.I22);
            sheet2.GetRow(18).GetCell(6).SetCellValue(barChooseInfo.I31);
            sheet2.GetRow(18).GetCell(7).SetCellValue(barChooseInfo.I32);
            sheet2.GetRow(18).GetCell(8).SetCellValue(barChooseInfo.I33);
            sheet2.GetRow(18).GetCell(10).SetCellValue(barChooseInfo.I4);

            sheet2.GetRow(19).GetCell(1).SetCellValue(barChooseInfo.CalCap1);
            sheet2.GetRow(19).GetCell(3).SetCellValue(barChooseInfo.CalCap2);
            sheet2.GetRow(19).GetCell(6).SetCellValue(barChooseInfo.CalCap3);
            sheet2.GetRow(19).GetCell(10).SetCellValue(barChooseInfo.CalCap4);
            sheet2.GetRow(20).GetCell(1).SetCellValue(barChooseInfo.MaxCap);
            sheet2.GetRow(21).GetCell(1).SetCellValue(barChooseInfo.ChooseCap);

            try
            {
                fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                workbook.Write(fs); //写入到excel
                return true;
            }
            catch
            {
                return false;            
            }
            
        }
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                if (disposing)
                {
                    if (fs != null)
                        fs.Close();
                }

                fs = null;
                disposed = true;
            }
        }
    }
}
