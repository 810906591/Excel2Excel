using Excel2Excel.Model;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Excel2Excel
{
    public partial class Form1 : Form
    {
        List<DerivedStatistics> returnVisitList = null;
        List<DerivedStatisticsVice> returnVisitListVice = null;
        private string newName;
        public Form1()
        {
            InitializeComponent();
        }

        private List<FileMapUtil> GetDerivedStatisticsLord()
        {
            List<FileMapUtil> pojoList = new List<FileMapUtil>();
            pojoList.Add(new FileMapUtil("流程编号", "ProcessNumber"));
            pojoList.Add(new FileMapUtil("创建日期", "CreationDate"));
            pojoList.Add(new FileMapUtil("申请人", "Claimant"));
            pojoList.Add(new FileMapUtil("类型", "Types"));
            pojoList.Add(new FileMapUtil("衍生料号", "DerivedPartNumber"));
            pojoList.Add(new FileMapUtil("产品经理", "ProductManager"));
            return pojoList;
        }
        private List<FileMapUtil> GetDerivedStatisticsVice()
        {
            List<FileMapUtil> pojoList = new List<FileMapUtil>();
            pojoList.Add(new FileMapUtil("物料编码", "DerivedPartNumber"));
            //pojoList.Add(new FileMapUtil("物料名称", "MaterialName"));
            //pojoList.Add(new FileMapUtil("规格型号", "Specifications"));
            //pojoList.Add(new FileMapUtil("存货类别", "Types"));
            //pojoList.Add(new FileMapUtil("单位", "Units"));
            pojoList.Add(new FileMapUtil("数量", "Number"));
            pojoList.Add(new FileMapUtil("金额", "Amounts"));
            return pojoList;
        }

        private void buttonLord_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBoxLordSheet.Text.Replace(" ", "")))
            {
                MessageBox.Show("主表工作簿不能为空");
                return;
            }
            using (OpenFileDialog open = new OpenFileDialog())
            {
                open.Filter = "Excel|*.xls;*.xlsx|All Files|*.*";
                if (open.ShowDialog() == DialogResult.OK)
                {

                    try
                    {
                        //returnVisitList = FileReadUtil.GM<DerivedStatistics>(GetDerivedStatisticsLord(), "衍生统计", "ProcessNumber", open.FileName);
                        returnVisitList = Function.Excel2Object.ExcelHelper.ExcelToObject<DerivedStatistics>(open.FileName, textBoxLordSheet.Text.Replace(" ", "")).ToList();
                        newName = Path.GetFileName(open.FileName);
                        textBoxLord.Text = open.FileName;
                    }
                    catch (Exception ex)
                    {
                        if (ex.Message.Contains("正由另一进程使用"))
                            MessageBox.Show("当前文件【" + Path.GetFileNameWithoutExtension(open.FileName) + "】已经被office打开");
                        else MessageBox.Show("错误信息：" + ex.Message);
                    }
                }
            }
        }

        private void buttonVice_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBoxViceSheet.Text.Replace(" ", "")))
            {
                MessageBox.Show("副表工作簿不能为空");
                return;
            }
            using (OpenFileDialog open = new OpenFileDialog())
            {
                open.Filter = "Excel|*.xls;*.xlsx|All Files|*.*";
                if (open.ShowDialog() == DialogResult.OK)
                {

                    try
                    {
                        //returnVisitListVice = FileReadUtil.GM<DerivedStatisticsVice>(GetDerivedStatisticsVice(), "Sheet1", "DerivedPartNumber", open.FileName);
                        returnVisitListVice = Function.Excel2Object.ExcelHelper.ExcelToObject<DerivedStatisticsVice>(open.FileName, textBoxViceSheet.Text.Replace(" ", "")).ToList();
                        textBoxVice.Text = open.FileName;
                        // var sameAges = returnVisitListVice.GroupBy(g => g.DerivedPartNumber).Where(s => s.Count() > 1).ToList();
                    }
                    catch (Exception ex)
                    {
                        if (ex.Message.Contains("正由另一进程使用"))
                            MessageBox.Show("当前文件【" + Path.GetFileNameWithoutExtension(open.FileName) + "】已经被office打开");
                        else MessageBox.Show("错误信息：" + ex.Message);
                    }
                }
            }
        }

        private void buttonNew_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog open = new FolderBrowserDialog())
            {

                if (open.ShowDialog() == DialogResult.OK)
                {
                    textBoxNew.Text = open.SelectedPath;
                }
            }
        }

        private void buttonConfirm_Click(object sender, EventArgs e)
        {
            if (returnVisitList.Count <= 0)
            {
                MessageBox.Show("请选择主表");
                return;
            }
            if (returnVisitListVice.Count <= 0)
            {
                MessageBox.Show("请选择主表");
                return;
            }
            string localFilePath = textBoxNew.Text;
            localFilePath = string.IsNullOrEmpty(localFilePath) ? Environment.GetFolderPath(Environment.SpecialFolder.Desktop) : localFilePath;
            if (!Directory.Exists(localFilePath))
                Directory.CreateDirectory(localFilePath);
            localFilePath += "\\new_" + newName;
            if (!File.Exists(localFilePath))
                File.Create(localFilePath).Close();
            labelFilePath.Text = "导出文件路径 " + localFilePath;
            var newExcel = from lord in returnVisitList
                           join vice in returnVisitListVice on lord.DerivedPartNumber equals vice.DerivedPartNumber into temp
                           from t in temp.DefaultIfEmpty()
                           select new DerivedStatisticsRep
                           {
                               ProcessNumber = lord.ProcessNumber,
                               CreationDate = lord.CreationDate,
                               Claimant = lord.Claimant,
                               Types = lord.Types,
                               DerivedPartNumber = lord.DerivedPartNumber,
                               ProductManager = lord.ProductManager,
                               Number = (t == null ? "0" : t.Number),
                               Amounts = (t == null ? "0" : t.Amounts)
                           };

            #region  Model2Excel-one
            //var TableName = new ModelHandler<DerivedStatisticsRep>().FillDataTable(newExcel.ToList()); //newExcel.ToList();
            ////数据初始化
            //int TotalCount;     //总行数
            //int RowRead = 0;    //已读行数
            //int Percent = 0;    //百分比
            //TotalCount = TableName.Rows.Count;
            ////NPOI
            //IWorkbook workbook;
            //string FileExt = Path.GetExtension(localFilePath).ToLower();
            //if (FileExt == ".xlsx")
            //    workbook = new XSSFWorkbook();
            //else if (FileExt == ".xls")
            //    workbook = new HSSFWorkbook();

            //else
            //    workbook = null;

            //if (workbook == null)
            //{
            //    MessageBox.Show("新的文件创建失败");
            //    return;
            //}
            //ISheet sheet = workbook.CreateSheet("Sheet1"); //string.IsNullOrEmpty(FileName) ? workbook.CreateSheet("Sheet1") : workbook.CreateSheet(FileName);
            ////秒钟
            //Stopwatch timer = new Stopwatch();
            //timer.Start();
            //try
            //{
            //    //读取标题  
            //    IRow rowHeader = sheet.CreateRow(0);
            //    for (int i = 0; i < TableName.Columns.Count; i++)
            //    {
            //        ICell cell = rowHeader.CreateCell(i);
            //        cell.SetCellValue(getalias(TableName.Columns[i].ColumnName));
            //    }

            //    //读取数据  
            //    for (int i = 0; i < TableName.Rows.Count; i++)
            //    {
            //        IRow rowData = sheet.CreateRow(i + 1);
            //        for (int j = 0; j < TableName.Columns.Count; j++)
            //        {
            //            ICell cell = rowData.CreateCell(j);
            //            cell.SetCellValue(TableName.Rows[i][j].ToString());
            //        }
            //        //状态栏显示
            //        RowRead++;
            //        Percent = (int)(100 * RowRead / TotalCount);
            //        //barStatus.Maximum = TotalCount;
            //        //barStatus.Value = RowRead;
            //        lblStatus.Text = "共有" + TotalCount + "条数据，已读取" + Percent.ToString() + "%的数据。";
            //        Application.DoEvents();
            //    }

            //    //状态栏更改
            //    lblStatus.Text = "正在生成Excel...";
            //    Application.DoEvents();

            //    //转为字节数组  
            //    MemoryStream stream = new MemoryStream();
            //    workbook.Write(stream);
            //    var buf = stream.ToArray();

            //    //保存为Excel文件  
            //    using (FileStream fs = new FileStream(localFilePath, FileMode.Create, FileAccess.Write))
            //    {
            //        fs.Write(buf, 0, buf.Length);
            //        fs.Flush();
            //        fs.Close();
            //    }

            //    //状态栏更改
            //    lblStatus.Text = "生成Excel成功，共耗时" + timer.ElapsedMilliseconds + "毫秒。";
            //    Application.DoEvents();

            //    //关闭秒钟
            //    timer.Reset();
            //    timer.Stop();

            //    //成功提示
            //    //if (MessageBox.Show("导出成功，是否立即打开？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
            //    //{
            //    //    System.Diagnostics.Process.Start(localFilePath);
            //    //}
            //    MessageBox.Show("导出成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
            //finally
            //{
            //    //关闭秒钟
            //    timer.Reset();
            //    timer.Stop();
            //}
            #endregion

            #region Model2Excel-two
            //秒钟
            Stopwatch timer = new Stopwatch();
            timer.Start();
            try
            {
                Function.Excel2Object.ExcelHelper.ObjectToExcel(newExcel.ToList(), localFilePath);
                lblStatus.Text = "生成Excel成功，共耗时" + timer.ElapsedMilliseconds + "毫秒。";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                //关闭秒钟
                timer.Reset();
                timer.Stop();
            }
            #endregion
        }

        private string getalias(string name)
        {
            string newname = name;
            switch (name)
            {
                case "ProcessNumber": newname = "流程编号"; break;
                case "CreationDate": newname = "创建日期"; break;
                case "Claimant": newname = "申请人"; break;
                case "Types": newname = "类型"; break;
                case "DerivedPartNumber": newname = "衍生料号"; break;
                case "ProductManager": newname = "产品经理"; break;
                case "Number": newname = "数量"; break;
                case "Amounts": newname = "金额"; break;
            }
            return newname;
        }
    }
}
