using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SHScoreForCollege103.DAO;
using Aspose.Cells;
using FISCA.UDT;

namespace SHScoreForCollege103
{
    public partial class ScoreForm : FISCA.Presentation.Controls.BaseForm
    {
        // 讀取資料
        BackgroundWorker _bgLoadMapping;
        // 產生資料
        BackgroundWorker _bgExporData;

        // 存在 UDT Mapping 資料
        List<FieldConfig> _FieldConfigList;

        // 儲存 UDT Mapping 資料用
        List<FieldConfig> _SaveFieldConfigList;


        public ScoreForm()
        {
            InitializeComponent();
            _FieldConfigList = new List<FieldConfig> ();
            _SaveFieldConfigList = new List<FieldConfig>();

            _bgLoadMapping = new BackgroundWorker();
            _bgLoadMapping.DoWork += _bgLoadMapping_DoWork;
            _bgLoadMapping.RunWorkerCompleted += _bgLoadMapping_RunWorkerCompleted;

            _bgExporData = new BackgroundWorker();
            _bgExporData.DoWork += _bgExporData_DoWork;
            _bgExporData.RunWorkerCompleted += _bgExporData_RunWorkerCompleted;

            
        }

        void _bgLoadMapping_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // 清除畫面並將資料填入
            dgData.Rows.Clear();

            // 檢查是否用預設欄位
            if (_FieldConfigList.Count > 3)
            {
                foreach (FieldConfig fc in _FieldConfigList)
                {
                    int rowIdx = dgData.Rows.Add();

                    dgData.Rows[rowIdx].Tag = fc;
                    dgData.Rows[rowIdx].Cells[colFieldName.Index].Value = fc.FieldName;
                    dgData.Rows[rowIdx].Cells[colFieldMapping.Index].Value = fc.FieldMapping;
                }
            }
            else
            { 
                // 使用預設
                List<string> filedList = GetDefaultFieldName();
                foreach (string str in filedList)
                {
                    int rowIdx = dgData.Rows.Add();
                    FieldConfig fc = new FieldConfig();
                    fc.FieldName = str;
                    dgData.Rows[rowIdx].Tag = fc;
                    dgData.Rows[rowIdx].Cells[colFieldName.Index].Value = fc.FieldName;
                    dgData.Rows[rowIdx].Cells[colFieldMapping.Index].Value = fc.FieldMapping;                    
                }

            }
            btnEnable(true);
        }

        void _bgLoadMapping_DoWork(object sender, DoWorkEventArgs e)
        {
            // 讀取 UDT 資料
            AccessHelper ah = new AccessHelper();
            List<FieldConfig> FieldConfigList = ah.Select<FieldConfig>();

            // 排序
            _FieldConfigList = (from data in FieldConfigList orderby data.FieldOrder ascending select data).ToList();
     
        }

        void _bgExporData_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            throw new NotImplementedException();
        }

        void _bgExporData_DoWork(object sender, DoWorkEventArgs e)
        {
            if (_SaveFieldConfigList.Count > 0)
            {

                Dictionary<string, int> FieldNameDict = new Dictionary<string, int>();
                Dictionary<string, int> FieldValueDict = new Dictionary<string, int>();

                // 讀取對


                // 取得所選學生ID
                List<string> StudentIDList = K12.Presentation.NLDPanels.Student.SelectedSource;

                // 取得學生成績資料
                
            
            }
        }        
    
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ScoreForm_Load(object sender, EventArgs e)
        {
            btnEnable(false);
            _bgLoadMapping.RunWorkerAsync();
        }

        private void btnEnable(bool bo)
        {
            btnExportMaping.Enabled = bo;
            btnImportMapping.Enabled = bo;
            btnExportCSV.Enabled = bo;
        }

        private void btnExportMaping_Click(object sender, EventArgs e)
        {
            // 匯出對照表
            Workbook wb = new Workbook();
            wb.Worksheets[0].Cells[0, 0].PutValue(dgData.Columns[colFieldName.Index].HeaderText);
            wb.Worksheets[0].Cells[0, 1].PutValue(dgData.Columns[colFieldMapping.Index].HeaderText);

            int rowIdx = 1;
            foreach (DataGridViewRow dgvr in dgData.Rows)
            {
                if (dgvr.IsNewRow)
                    continue;

                int colIdx = 0;
                foreach (DataGridViewCell dgvc in dgvr.Cells)
                {
                    if (dgvc.Value != null)
                        wb.Worksheets[0].Cells[rowIdx, colIdx].PutValue(dgvc.Value.ToString());

                    colIdx++;
                }
                rowIdx++;
            }
            Utility.CompletedXls("甄選對應表", wb);
        
        }

        private void btnImportMapping_Click(object sender, EventArgs e)
        {
            // 匯入對照表
         
            
        }

        /// <summary>
        /// 預設資料欄位
        /// </summary>
        /// <returns></returns>
        private List<string> GetDefaultFieldName()
        {
            List<string> retVal = new List<string>();
            retVal.Add("學號");
            retVal.Add("姓名");
            retVal.Add("身分證號碼");
            retVal.Add("學業總平均(高一上)");
            retVal.Add("國文(高一上)");
            retVal.Add("英文(高一上)");
            retVal.Add("數學(高一上)");
            retVal.Add("物理(高一上)");
            retVal.Add("化學(高一上)");
            retVal.Add("生物(高一上)");
            retVal.Add("地球科學(高一上)");
            retVal.Add("歷史(高一上)");
            retVal.Add("地理(高一上)");
            retVal.Add("公民與社會(高一上)");
            retVal.Add("音樂(高一上)");
            retVal.Add("美術(高一上)");
            retVal.Add("舞蹈(高一上)");
            retVal.Add("體育(高一上)");
            retVal.Add("藝術生活(高一上)");
            retVal.Add("生活科技(高一上)");
            retVal.Add("家政(高一上)");
            retVal.Add("資訊科技概論(高一上)");
            retVal.Add("健康與護理(高一上)");
            retVal.Add("全民國防教育(高一上)");
            retVal.Add("學業總平均(高一下)");
            retVal.Add("國文(高一下)");
            retVal.Add("英文(高一下)");
            retVal.Add("數學(高一下)");
            retVal.Add("物理(高一下)");
            retVal.Add("化學(高一下)");
            retVal.Add("生物(高一下)");
            retVal.Add("地球科學(高一下)");
            retVal.Add("歷史(高一下)");
            retVal.Add("地理(高一下)");
            retVal.Add("公民與社會(高一下)");
            retVal.Add("音樂(高一下)");
            retVal.Add("美術(高一下)");
            retVal.Add("舞蹈(高一下)");
            retVal.Add("體育(高一下)");
            retVal.Add("藝術生活(高一下)");
            retVal.Add("生活科技(高一下)");
            retVal.Add("家政(高一下)");
            retVal.Add("資訊科技概論(高一下)");
            retVal.Add("健康與護理(高一下)");
            retVal.Add("全民國防教育(高一下)");
            retVal.Add("就讀科、學程、班別");

            return retVal;
        
        }

        private void btnExportCSV_Click(object sender, EventArgs e)
        {
            if (SaveConfig())
            {
                // 產生資料
                _bgExporData.RunWorkerAsync();
            
            }
        }

        /// <summary>
        /// 儲存設定值
        /// </summary>
        private bool SaveConfig()
        {
            bool pass = true;
            // 檢查畫面資料是否重複
            List<string> chkStr = new List<string>();
            bool hasSame = false;
            foreach (DataGridViewRow dgvr in dgData.Rows)
            {
                if (dgvr.IsNewRow)
                    continue;

                foreach (DataGridViewCell dgvc in dgvr.Cells)
                {
                    if (dgvc.Value != null)
                        if (dgvc.ColumnIndex == colFieldName.Index)
                        {
                            string key = dgvc.Value.ToString();
                            if (!chkStr.Contains(key))
                                chkStr.Add(key);
                            else
                            {
                                hasSame = true;
                                break;
                            }
                        }
                }

                if (hasSame)
                    break;
            }

            if (hasSame)
            {
                FISCA.Presentation.Controls.MsgBox.Show("有相同欄位名稱無法產生，請檢查!");
                pass = false;
            }
            else
            { 
                // 儲存畫面值到 UDT
                
                List<string> fiedNameList = new List<string>();
                _SaveFieldConfigList.Clear();
                int fieldOrder = 0;
                foreach (DataGridViewRow dgvr in dgData.Rows)
                {
                    if (dgvr.IsNewRow)
                        continue;

                    FieldConfig fc = dgvr.Tag as FieldConfig;
                    if (fc == null)
                        fc = new FieldConfig();

                    if (dgvr.Cells[colFieldName.Index].Value != null)
                        fc.FieldName = dgvr.Cells[colFieldName.Index].Value.ToString();

                    if (dgvr.Cells[colFieldMapping.Index].Value != null)
                        fc.FieldMapping = dgvr.Cells[colFieldMapping.Index].Value.ToString();

                    // 欄位順序重設
                    fc.FieldOrder = fieldOrder;
                    fieldOrder++;

                    fiedNameList.Add(fc.FieldName);
                    _SaveFieldConfigList.Add(fc);
                }
            
                // 檢查不要需要刪除
                List<FieldConfig> delDataList = new List<FieldConfig>();

                foreach (FieldConfig fc in _FieldConfigList)
                {
                    if (!fiedNameList.Contains(fc.FieldName))
                        delDataList.Add(fc);
                }

                // 刪除多餘資料
                if (delDataList.Count > 0)
                {
                    foreach (FieldConfig fc in delDataList)
                        fc.Deleted = true;
                    delDataList.SaveAll();                
                }

                // 儲存資料
                _SaveFieldConfigList.SaveAll();
            }
            return pass;
        }



    }
}
