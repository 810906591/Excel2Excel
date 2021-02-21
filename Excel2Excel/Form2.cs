using Excel2Excel.Model;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Excel2Excel
{
    public partial class Form2 : Form
    {
        List<NewProducts> returnVisitList = null;
        List<DerivedStatisticsVice> returnVisitListVice = null;
        private string newName;
        public Form2()
        {
            InitializeComponent();
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
                        returnVisitList = Function.Excel2Object.ExcelHelper.ExcelToObject<NewProducts>(open.FileName, textBoxLordSheet.Text.Replace(" ", "")).ToList();
                        newName = Path.GetFileNameWithoutExtension(open.FileName);
                        textBoxLord.Text = open.FileName;
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("当前文件【" + Path.GetFileNameWithoutExtension(open.FileName) + "】已经被office打开");
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
                        returnVisitListVice = Function.Excel2Object.ExcelHelper.ExcelToObject<DerivedStatisticsVice>(open.FileName, textBoxViceSheet.Text.Replace(" ", "")).ToList();
                        textBoxVice.Text = open.FileName;
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("当前文件【" + Path.GetFileNameWithoutExtension(open.FileName) + "】已经被office打开");
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
            if (returnVisitList.Count == 0 || (returnVisitList.Count > 0 && returnVisitList.Count(o => !string.IsNullOrEmpty(o.ModelNum)) <= 0))
            {
                MessageBox.Show(this, "型号表数据为空");
                return;
            }
            if (returnVisitListVice.Count == 0)
            {
                MessageBox.Show(this, "料号表数据为空");
                return;
            }
            string strNewPro = string.Join("、", returnVisitList.Where(o => !string.IsNullOrEmpty(o.ModelNum)).Select(o => o.ModelNum.TrimStart('、').TrimEnd('、').Trim()).ToList());
            var listNewPro = strNewPro.Split('、', '\n').ToList().Distinct().OrderByDescending(o => o.ToString().Length).ToList();

            foreach (var item in returnVisitListVice.Where(o => string.IsNullOrEmpty(o.TagType)).ToList())
            {
                foreach (var itemnp in listNewPro)
                {
                    string regexStr = string.Format("(?<!([0-9a-zA-Z-])){0}(?!([0-9a-zA-Z]))", itemnp);
                    if (string.IsNullOrEmpty(item.TagType) && item.MaterialName.Contains(itemnp) && Regex.IsMatch(item.MaterialName, regexStr))
                    {
                        item.TagType = itemnp;
                        break;
                    }
                }
            }

            returnVisitListVice.RemoveAll(o => string.IsNullOrEmpty(o.TagType));

            #region Model2Excel-two
            string localFilePath = textBoxNew.Text;
            localFilePath = string.IsNullOrEmpty(localFilePath) ? Environment.GetFolderPath(Environment.SpecialFolder.Desktop) : localFilePath;
            if (!Directory.Exists(localFilePath))
                Directory.CreateDirectory(localFilePath);
            textBoxNew.Text = localFilePath;
            localFilePath += "\\new_" + newName + ".xls";
            if (!File.Exists(localFilePath))
                File.Create(localFilePath).Close();
            labelFilePath.Text = "导出文件路径 " + localFilePath;
            //秒钟
            Stopwatch timer = new Stopwatch();
            timer.Start();
            try
            {
                Function.Excel2Object.ExcelHelper.ObjectToExcel(returnVisitListVice.OrderBy(o => o.TagType).ThenBy(o=>o.TagType.Length).ToList(), localFilePath);
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




    }
}
