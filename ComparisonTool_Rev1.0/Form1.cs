
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Net.Sockets;
using static OfficeOpenXml.ExcelErrorValue;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using System.Runtime.InteropServices.ComTypes;
using System.Diagnostics;

namespace ComparisonTool_Rev1._0
{
    public partial class Form1 : Form
    {
        private string? TemplatePath = null;
        private List<string>? MultipleFilesPaths = new List<string>();   // 建立上傳檔案清單

        // 建立範本的數據集合
        private int BlackBGCol = 0;                         // RTS1.0_BCT RTS1.1_CT位置(背景色唯一黑色)
        private int AllStartCol = 0;                        // 起始讀取欄位置
        private int AllStartRow = 0;                        // 起始讀取列位置
        private int AllEndCol = 0;                          // 結束讀取欄位置
        private int AllEndRow = 0;                          // 結束讀取列位置
        private int ComStartRow = 0;                        // 起始Combine讀取列位置
        private int ComEndRow = 0;                          // 結束Combine讀取列位置
        private int SURow = 0;                              // SU讀取列位置
        private int TempColCount = 0;                       // 範本欄總數量
        private string? TemplateRepeatData = null;          // 範本重複的列
        private string? CompareRepeatData = null;           // 比對檔案重複的列

        private List<int> RYGList = new List<int> { };          // 取得欄 RYG 位置
        private List<int> StartList = new List<int> { };        // 取得Fuction Team 讀取起始位置
        private List<int> EndList = new List<int> { };          // 取得Fuction Team 讀取結束位置
        private List<int> OdmStartColList = new List<int> { };  // 取得 ODM 起始欄 位置清單
        private List<int> OdmEndColList = new List<int> { };    // 取得 ODM 結束欄 位置清單
        private List<int> OdmStartRowList = new List<int> { };  // 取得 ODM 起始列 位置清單
        private List<int> OdmEndRowList = new List<int> { };    // 取得 ODM 結束列 位置清單
        private List<string> GroupList = new List<string>       // 預設 Title Group
        {
            "Category", "DPN", "Description", "Manufacturers", "Remark"
        };

        private Dictionary<string, int> TemplateDataDict = new Dictionary<string, int>();   // 取得範本 Group 值
        private Dictionary<int, string> TempAllData = new Dictionary<int, string>();        // 取得範本 Group 後面各欄數值
        private Dictionary<int, string> TempAllColData = new Dictionary<int, string>();     // 取得範本 Group 後面各欄位置

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // 預設選中首次使用
            FirstUse.Checked = true;
            UploadFirstTemplateBtn.Visible = true;
            NonFileFirst.Visible = true;
            UploadLastTemplateBtn.Visible = false;
            NonFileLast.Visible = false;
        }

        // 首次上傳範本的Radio Button觸發事件
        private void FirstUseChecked(object sender, EventArgs e)
        {
            if (FirstUse.Checked)
            {
                // 首次使用範本
                UploadFirstTemplateBtn.Visible = true;
                NonFileFirst.Visible = true;
                UploadLastTemplateBtn.Visible = false;
                NonFileLast.Visible = false;
                TemplatePath = null; // 清空預選範本
                NonFileFirst.Text = "(尚未上傳檔案！)";
            }
        }

        // 前次使用範本的Radio Button觸發事件
        private void LastUseChecked(object sender, EventArgs e)
        {
            if (LastUse.Checked)
            {
                // 前次使用範本
                UploadFirstTemplateBtn.Visible = false;
                NonFileFirst.Visible = false;
                UploadLastTemplateBtn.Visible = true;
                NonFileLast.Visible = true;
                TemplatePath = null; // 清空預選範本
                NonFileLast.Text = "(尚未上傳檔案！)";
            }
        }

        // 首次使用範本上傳按鈕
        private void UploadFirstTemplate(object sender, EventArgs e)
        {
            // 設定 EPPlus 授權模式，非商業用途
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx";
                openFileDialog.Title = "請選擇空白範本檔案!";
                openFileDialog.Multiselect = false;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    TemplatePath = openFileDialog.FileName;
                    MessageBox.Show("範本檔案已選擇: " + TemplatePath);
                    NonFileFirst.Text = "(檔案已上傳！)";
                }
                else
                {
                    MessageBox.Show("請選擇空白範本檔案。");
                    return;
                }
            }
        }

        // 前次使用非空白範本上傳按鈕
        private void UploadLastTemplate(object sender, EventArgs e)
        {
            // 設定 EPPlus 授權模式，非商業用途
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx";
                openFileDialog.Title = "請選擇非空白範本檔案!";
                openFileDialog.Multiselect = false;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    TemplatePath = openFileDialog.FileName;
                    MessageBox.Show("範本檔案已選擇: " + TemplatePath);
                    NonFileLast.Text = "(檔案已上傳！)";
                }

                if (TemplatePath == null)
                {
                    MessageBox.Show("請選擇非空白範本檔案。");
                    return;
                }
            }
        }

        // 多筆檔案上傳按鈕
        private void UploadMultipleFile(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx";
                openFileDialog.Title = "請選擇多筆檔案!";
                openFileDialog.Multiselect = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    MultipleFilesPaths = new List<string>(openFileDialog.FileNames);
                    MessageBox.Show($"已選擇 {MultipleFilesPaths.Count} 筆檔案。");
                    NonFIleMuti.Text = "(檔案已上傳！)";
                }

                if (MultipleFilesPaths == null || MultipleFilesPaths.Count == 0)
                {
                    MessageBox.Show("請選擇至少一個檔案。");
                }
            }
        }


        // 送出按鈕
        private void SummitBtn(object sender, EventArgs e)
        {
            if (TemplatePath == null)
            {
                MessageBox.Show("請先上傳範本檔案。");
                return;
            }

            if (MultipleFilesPaths == null || MultipleFilesPaths.Count == 0)
            {
                MessageBox.Show("請選擇多筆 Excel 檔案進行比對。");
                return;
            }

            // 開始處理 Excel 比對
            CombineExcelFiles(TemplatePath, MultipleFilesPaths);

            // 清空暫存資料
            ClearCache();
        }

        private void CombineExcelFiles(string TempPath, List<string> filePaths)
        {
            try
            {
                // 以範本檔案為基礎
                using (ExcelPackage TemplatePackage = new ExcelPackage(new FileInfo(TempPath)))
                {
                    ExcelWorksheet TemplateSheet = TemplatePackage.Workbook.Worksheets[0];
                    string TempName = TemplatePackage.File.Name;

                    using (ExcelPackage newPackage = new ExcelPackage())
                    {
                        // 複製範本工作表到新檔案
                        ExcelWorksheet Polling1 = newPackage.Workbook.Worksheets.Add("Polling1", TemplateSheet);

                        try
                        {
                            GetTemplateData(Polling1, TemplateSheet);  //取得範本資料 Category到 Remark Group 
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"讀取範本資料時發生錯誤，請檢查範本檔案：{TempName}\n錯誤訊息：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        if (!string.IsNullOrEmpty(TemplateRepeatData))
                        {
                            TemplateRepeatData += "為重複資料，請檢查範本檔案：" + TempName;
                            MessageBox.Show(TemplateRepeatData);
                            return;
                        }

                        
                        // 從 C 欄開始記錄差異比對
                        int colIndex = 3;
                        foreach (string filePath in filePaths)
                        {
                            try
                            {
                                using (ExcelPackage comparePackage = new ExcelPackage(new FileInfo(filePath)))
                                {
                                    ExcelWorksheet CompareSheet = comparePackage.Workbook.Worksheets[0];
                                    string FileName = comparePackage.File.Name;
                                    try
                                    {
                                        CompareSheets(Polling1, CompareSheet, TemplateSheet, FileName, colIndex);
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show($"讀取比對檔案時發生錯誤，請檢查檔案：{FileName}\n錯誤訊息：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        return;
                                    }

                                    if (!string.IsNullOrEmpty(CompareRepeatData))
                                    {
                                        CompareRepeatData += "為重複資料，請檢查比對檔案：" + FileName;
                                        MessageBox.Show(CompareRepeatData);
                                        return;
                                    }
                                    colIndex++;
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"讀取比對檔案時發生錯誤，檔案路徑為：{filePath}\n錯誤訊息：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                        }
                        
                        try
                        {
                            // 複製 Polling1 內容到 Polling2
                            ExcelWorksheet Polling2 = newPackage.Workbook.Worksheets.Add("Polling2", Polling1);
                            CombinePolling(Polling2);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"合併資料時發生錯誤：\n錯誤訊息：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        try
                        {
                            // 儲存新的 Excel 檔案
                            SaveAsNewExcel(newPackage);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"另存新檔時發生錯誤：\n錯誤訊息：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"讀取範本檔案發生錯誤，請檢查範本檔案!\n檔案路徑：{TempPath}\n錯誤訊息：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void GetTemplateData(ExcelWorksheet Polling1, ExcelWorksheet TemplateSheet)
        {
            int TemplateMaxRow = TemplateSheet.Dimension.End.Row;
            int TemplateMaxCol = TemplateSheet.Dimension.End.Column;
            int StartTitleCol = 0;      // 取得欄 Category 位置
            int StartTitleRow = 0;      // 取得列 Category 位置
            int EndTitleCol = 0;        // 取得欄 Remark 位置
            int UpdateStartCol = 0;     // 取得欄 Revision 位置
            int StartRow = 0;           // 取得列 起始 位置
            int RemoveCol = 0;          // 取得欄 Remove 位置
            int RemarkIndex = 0;        // RemarkIndex
            int RevisionIndex = 0;      // RevisionIndex
            int SUIndex = 0;            // SUIndex
            int EndIndex = 0;           // EndIndex

            List<int> TitleList = new List<int> { };  // 取得欄 Group 位置

            // 取消複製範本篩選功能
            Polling1.Cells.AutoFilter = false;

            for (int row = 1; row <= TemplateMaxRow; row++)
            {
                // 取消複製範本群組(列)
                Polling1.Row(row).OutlineLevel = 0;

                // 取消複製範本隱藏列
                if (Polling1.Row(row).Hidden)
                {
                    Polling1.Row(row).Hidden = false;
                }

                string RowData = "";
                List<string> TempData = new List<string>();
                List<int> TempColData = new List<int>();
                int RemoveSign = 0;  // Remove欄位標記

                for (int col = 1; col <= TemplateMaxCol; col++)
                {
                    // 取消複製範本群組(欄)
                    Polling1.Column(col).OutlineLevel = 0;

                    // 取消複製範本隱藏欄
                    if (Polling1.Column(col).Hidden)
                    {
                        Polling1.Column(col).Hidden = false;
                    }

                    // 取得每個欄位數值
                    var CellValue = TemplateSheet.Cells[row, col].Text;

                    if (CellValue == "Start")
                    {
                        AllStartCol = col + 1;      // 標題欄 Start 位置
                        AllStartRow = row + 3;      // 標題欄 Start 位置
                        ComEndRow = row;            // Polling2 結束標題列位置
                        StartRow = row + 1;         // 起始讀取列位置
                    }
                    else if (CellValue == "Remove" && StartTitleCol == 0)
                    {
                        RemoveCol = col;            // 標題欄 Remove 位置
                    }
                    else if (CellValue == "Category")
                    {
                        StartTitleCol = col;        // 標題欄 Category 位置
                        StartTitleRow = row;
                        ComStartRow = row;          // Polling2 起始標題列位置
                    }
                    else if (CellValue == "Remark" && row == StartTitleRow && RemarkIndex == 0)
                    {
                        EndTitleCol = col;          // 標題欄 Remark 位置
                        RemarkIndex++;
                    }
                    else if (CellValue == "Revision")
                    {
                        if (RevisionIndex == 0)
                        {
                            UpdateStartCol = col;   // Revision為起始讀取欄位
                            RevisionIndex++;
                        }
                        StartList.Add(col + 1);
                    }
                    else if (CellValue.StartsWith("1S1U") && SUIndex == 0)
                    {
                        SURow = row;                // SU讀取列位置
                        SUIndex++;
                    }
                    else if (CellValue.Replace(" ", "").Replace("\r", "").Replace("\n", "") == "1SODMTotal")
                    {
                        EndList.Add(col - 1);       // 新增Function Team 讀取結束欄位位置
                        OdmStartColList.Add(col);   // 新增 ODM Total 讀取起始欄位位置
                        OdmStartRowList.Add(row);   // 新增 ODM Total 讀取起始列位置
                    }
                    else if (CellValue.Replace(" ", "").Replace("\r", "").Replace("\n", "") == "1S+2SODMTotal")
                    {
                        OdmEndColList.Add(col);
                        OdmEndRowList.Add(row);
                    }
                    else if (CellValue.Replace(" ", "").Replace("\r", "").Replace("\n", "") == "RTS1.0_BCTRTS1.1_CT")
                    {
                        BlackBGCol = col;           // 標題欄 0_BCTRTS1 位置
                    }
                    else if (CellValue == "End")
                    {
                        if (EndIndex == 0)
                        {
                            AllEndCol = col;        // 結束讀取欄位置
                            TempColCount = (AllEndCol - 1) - AllStartCol + 1; // 範本欄總數量
                            EndIndex++;
                        }
                        else if (EndIndex != 0 && AllEndCol != 0 && col < AllEndCol)
                        {
                            AllEndRow = row;        // 結束讀取列位置
                        }
                    }
                    else if (CellValue == "RYG" && StartRow == 0 && AllEndCol == 0)
                    {
                        RYGList.Add(col);           // 標題欄 RYG 位置
                    }

                    // 將Title的欄加入TitleList
                    if (!TitleList.Contains(col) && GroupList.Contains(CellValue) && row == StartTitleRow)
                    {
                        TitleList.Add(col);
                    }

                    if (StartTitleCol != 0 && EndTitleCol != 0 && StartRow != 0 && col >= StartTitleCol && col <= EndTitleCol && row >= StartRow && AllEndRow == 0 && TitleList.Contains(col))
                    {
                        RowData += (col == StartTitleCol) ? TemplateSheet.Cells[row, col].Text.Replace(" ", "") + "|" : TemplateSheet.Cells[row, col].Text.Trim() + "|";
                    }

                    if (UpdateStartCol != 0 && StartRow != 0 && col >= UpdateStartCol && row >= StartRow && AllEndCol != 0 && col < AllEndCol && AllEndRow == 0)
                    {
                        string? TempValue = null;
                        if (!string.IsNullOrEmpty(TemplateSheet.Cells[row, col].Formula))
                        {
                            TempValue = TemplateSheet.Cells[row, col].Formula;
                        }
                        else
                        {
                            TempValue = (GetCellValue(TemplateSheet.Cells[row, col]) != null) ? GetCellValue(TemplateSheet.Cells[row, col]).ToString() : null;
                        }
                        TempData.Add(TempValue);
                        TempColData.Add(col);
                    }

                    // Remove欄位有值時，標記RemoveSign
                    if (RemoveCol != 0 && col == RemoveCol && !string.IsNullOrEmpty(CellValue))
                    {
                        RemoveSign = 1;
                    }

                    
                    if (StartRow != 0 && row >= StartRow && AllStartCol != 0 && AllEndCol != 0 && col >= AllStartCol && col < AllEndCol && AllEndRow == 0)
                    {
                        ResetBackgroundColor(Polling1, TemplateSheet, row, col, RemoveSign);
                    }
                }

                // 只儲存資料內容，不包括列號
                if (!string.IsNullOrEmpty(RowData))
                {
                    //判斷是否有重複的Key
                    if (!TemplateDataDict.ContainsKey(RowData))
                    {
                        TemplateDataDict.Add(RowData, row); // 將列號當作Key
                    }
                    else
                    {
                        TemplateRepeatData += "第" + row + "列, ";
                    }
                }

                if (StartRow != 0 && row >= StartRow && UpdateStartCol != 0 && AllEndRow == 0)
                {
                    // 合併儲存格
                    string AllRowData = string.Join("|", TempData) + "|";
                    string AllColData = string.Join("|", TempColData) + "|";
                    TempAllData[row] = AllRowData;
                    TempAllColData[row] = AllColData;
                }

                // 如果AllEndRow有值，則跳出迴圈
                if (AllEndRow != 0)
                {
                    break;
                }
            }
        }

        private void ResetBackgroundColor(ExcelWorksheet Polling1, ExcelWorksheet TemplateSheet, int row, int col, int RemoveSign)
        {
            // 先取出原欄位數值
            object OriginalValue = GetCellValue(TemplateSheet.Cells[row, col]);

            // 判斷是否為RYG欄
            if (!RYGList.Contains(col))
            {
                // Start~End所有儲存格重置背景色
                Polling1.Cells[row, col].Clear();
                Polling1.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.None;

                // 如果Remove欄有值則設置灰色
                if (RemoveSign == 1)
                {
                    Polling1.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Polling1.Cells[row, col].Style.Fill.BackgroundColor.SetColor(Color.DimGray);

                    // 新增刪除線
                    Polling1.Cells[row, col].Style.Font.Strike = true;
                }
            }

            // 回填欄位值
            if (!string.IsNullOrEmpty(TemplateSheet.Cells[row, col].Formula))
            {
                // 使用正則表達式替換列號
                Polling1.Cells[row, col].Formula = TemplateSheet.Cells[row, col].Formula;
            }
            else
            {
                Polling1.Cells[row, col].Value = OriginalValue;
            }

            // 設置文字為黑色
            Polling1.Cells[row, col].Style.Font.Color.SetColor(Color.Black);

            // 設置文字置中
            Polling1.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            // 設置數字格式
            Polling1.Cells[row, col].Style.Numberformat.Format = TemplateSheet.Cells[row, col].Style.Numberformat.Format;

            // 設置框線
            SetBorder(Polling1, row, col);
        }

        private void CompareSheets(ExcelWorksheet Polling1, ExcelWorksheet CompareSheet, ExcelWorksheet TemplateSheet, string FileName, int colIndex)
        {
            // 取消篩選功能
            CompareSheet.Cells.AutoFilter = false;

            int TemplateMaxRow = TemplateSheet.Dimension.End.Row;
            int CompareMaxRow = CompareSheet.Dimension.End.Row;
            int maxCol = Math.Max(TemplateSheet.Dimension.End.Column, CompareSheet.Dimension.End.Column);

            int RemoveCol = 0;          // 取得欄 Remove 位置
            int StartCol = 0;           // 取得欄 Start 位置
            int StartRow = 0;           // 取得列 起始 位置
            int StartTitleCol = 0;      // 取得欄 Category 位置
            int StartTitleRow = 0;      // 取得列 Category 位置
            int EndTitleCol = 0;        // 取得欄 Remark 位置
            int UpdateStartCol = 0;     // 取得欄 Revision 位置
            int UpdateEndCol = 0;       // 取得欄 1S ODM Total 位置
            int EndCol = 0;             // 取得欄 End 位置
            int EndRow = 0;             // 取得列 End 位置
            int CompareColCount = 0;    // 比對檔案欄總數量
            int RemarkIndex = 0;        // RemarkIndex
            int RevisionIndex = 0;      // RevisionIndex
            int EndIndex = 0;           // EndIndex
            Dictionary<string, int> CompareDataDict = new Dictionary<string, int>();    // 取得比對檔案Group值
            List<int> TitleList = new List<int> { };  // 取得欄 Group 位置
            int SeparatedDataIndex = 0;

            // 比對每一列
            for (int row = 1; row <= CompareMaxRow; row++)
            {
                // 設置Remove Index為0
                int RemoveIndex = 0;

                // 設置 Category到 Remark 集合
                string CompareData = "";

                // 以'|'分割後的陣列宣告
                string[] SeparatedData = Array.Empty<string>();
                int SeparatedColIndex = 0;
                int IndexStatus = 0;

                // 比對每一欄
                for (int col = 1; col <= maxCol; col++)
                {
                    string cellAddress = TemplateSheet.Cells[row, col].Address;
                    var CellCombineTitle = CompareSheet.Cells[row, col].Text;
                    var CellTempTitle = TemplateSheet.Cells[row, col].Text;

                    if (CellCombineTitle == "Start")
                    {
                        StartCol = col + 1;         // 標題欄 Start 位置
                        StartRow = row + 1;         // 起始讀取列位置
                    }
                    else if (CellCombineTitle == "Remove" && StartTitleCol == 0)
                    {
                        RemoveCol = col;            // 標題欄 Remove 位置
                    }
                    else if (CellCombineTitle == "Category")
                    {
                        StartTitleCol = col;        // 標題欄 Category 位置
                        StartTitleRow = row;
                    }
                    else if (CellCombineTitle == "Remark" && row == StartTitleRow && RemarkIndex == 0)
                    {
                        EndTitleCol = col;           // 標題欄 Remark 位置
                        RemarkIndex++;
                    }
                    else if (CellCombineTitle == "Revision" && RevisionIndex == 0)
                    {
                        UpdateStartCol = col;   // Revision為起始讀取欄位
                        RevisionIndex++;
                    }
                    else if (CellCombineTitle.Replace(" ", "").Replace("\r", "").Replace("\n", "") == "1SODMTotal")
                    {
                        UpdateEndCol = col;         // Function Team填寫結束的後一個欄位
                    }
                    else if (CellCombineTitle == "End")
                    {
                        if (EndIndex == 0)
                        {
                            EndCol = col;           // 結束讀取欄位置
                            CompareColCount = (EndCol - 1) - StartCol + 1; // 比對檔案欄總數量
                            EndIndex++;
                        }
                        else if (EndIndex != 0 && EndCol != 0 && col < EndCol)
                        {
                            EndRow = row;           // 結束讀取列位置
                        }
                    }

                    if (CompareColCount != 0 && CompareColCount != TempColCount)
                    {
                        MessageBox.Show($"比對檔案與範本檔案欄位數量不符，請檢查比對檔案：{FileName}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // 將Title的欄加入TitleList
                    if (!TitleList.Contains(col) && GroupList.Contains(CellCombineTitle) && row == StartTitleRow)
                    {
                        TitleList.Add(col);
                    }

                    // 儲存需比對欄位的數據集合
                    if (StartTitleCol != 0 && EndTitleCol != 0 && StartRow != 0 && col >= StartTitleCol && col <= EndTitleCol && row >= StartRow && EndRow == 0 && TitleList.Contains(col))
                    {
                        CompareData += (col == StartTitleCol) ? CompareSheet.Cells[row, col].Text.Replace(" ", "") + "|" : CompareSheet.Cells[row, col].Text.Trim() + "|";
                        if (col == EndTitleCol)
                        {
                            //判斷是否有重複的Key
                            if (!CompareDataDict.ContainsKey(CompareData))
                            {
                                CompareDataDict.Add(CompareData, row);
                            }
                            else
                            {
                                CompareRepeatData += "第" + row + "列, ";
                            }
                        }
                    }

                    if (StartRow != 0 && row >= StartRow && EndRow == 0)
                    {
                        // 以'|'分割後的陣列宣告
                        string[] SeparatedColData = Array.Empty<string>();

                        // 判定Remove欄位是否有值
                        if (RemoveIndex == 0 && RemoveCol != 0 && col == RemoveCol && !string.IsNullOrEmpty(CellCombineTitle))
                        {
                            RemoveIndex = 1;
                        }

                        // 判定比對檔案此列資料是否存在範本資料中
                        if (!string.IsNullOrEmpty(CompareData) && TemplateDataDict.ContainsKey(CompareData) && UpdateStartCol != 0 && col >= UpdateStartCol && EndCol != 0 && col < EndCol)
                        {
                            int RowKey = row;
                            row = TemplateDataDict[CompareData];

                            string? compareValue = null;
                            if (!string.IsNullOrEmpty(CompareSheet.Cells[RowKey, col].Formula))
                            {
                                compareValue = CompareSheet.Cells[RowKey, col].Formula;
                            }
                            else
                            {
                                compareValue = (GetCellValue(CompareSheet.Cells[RowKey, col]) != null) ? GetCellValue(CompareSheet.Cells[RowKey, col]).ToString() : null;
                            }

                            // TempAllData 包含指定的 Key
                            TempAllData.TryGetValue(row, out string? TempRowData);
                            TempAllColData.TryGetValue(row, out string? TempColData);

                            // 以'|'分割字串儲存陣列
                            SeparatedData = TempRowData.Split('|');
                            SeparatedColData = TempColData.Split('|');


                            // 移除最後的空白元素（如果字串結尾有多餘的 '|'）
                            if (SeparatedData.Length > 0 && string.IsNullOrEmpty(SeparatedData[SeparatedData.Length - 1]))
                            {
                                Array.Resize(ref SeparatedData, SeparatedData.Length - 1);
                            }
                            if (SeparatedColData.Length > 0 && string.IsNullOrEmpty(SeparatedColData[SeparatedColData.Length - 1]))
                            {
                                Array.Resize(ref SeparatedColData, SeparatedColData.Length - 1);
                            }

                            // 取得分割後欄位數值
                            string SeparatedValue = SeparatedData[SeparatedColIndex];
                            int SeparatedCol = int.Parse(SeparatedColData[SeparatedColIndex]);

                            if (!Equals(SeparatedValue, compareValue) && ((string.IsNullOrEmpty(SeparatedValue) && !string.IsNullOrEmpty(compareValue)) || (!string.IsNullOrEmpty(SeparatedValue) && !string.IsNullOrEmpty(compareValue))))
                            {
                                if (string.IsNullOrEmpty(CompareSheet.Cells[RowKey, col].Formula))
                                {
                                    Polling1.Cells[row, SeparatedCol].Value = GetCellValue(CompareSheet.Cells[RowKey, col]);
                                    

                                    // 設置數字格式
                                    Polling1.Cells[row, SeparatedCol].Style.Numberformat.Format = CompareSheet.Cells[RowKey, col].Style.Numberformat.Format;

                                    // 設置新增、修改欄位背景色及數值
                                    Polling1.Cells[row, SeparatedCol].Style.Fill.PatternType = ExcelFillStyle.None;
                                    if (string.IsNullOrEmpty(SeparatedValue) && !string.IsNullOrEmpty(compareValue))
                                    {
                                        Polling1.Cells[row, SeparatedCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        Polling1.Cells[row, SeparatedCol].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(204, 153, 255));
                                    }
                                    else if (!string.IsNullOrEmpty(SeparatedValue) && !string.IsNullOrEmpty(compareValue))
                                    {
                                        Polling1.Cells[row, SeparatedCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        Polling1.Cells[row, SeparatedCol].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                        Polling1.Cells[row, SeparatedCol].Style.Font.Color.SetColor(Color.White);
                                    }
                                }
                            }

                            SeparatedColIndex++;
                            row = RowKey;
                            IndexStatus = 1;
                        }
                    }

                    if (EndRow != 0)
                    {
                        break;
                    }
                }
                if (IndexStatus == 1)
                {
                    SeparatedDataIndex++;
                }
            }

            var addrows = new HashSet<string>(CompareDataDict.Keys.Except(TemplateDataDict.Keys));
            var delrows = new HashSet<string>(TemplateDataDict.Keys.Except(CompareDataDict.Keys));
            // 判斷是否有新增列
            if (addrows.Count > 0)
            {
                foreach (var addrowskey in addrows)
                {
                    // 取得總表最後一列
                    int LastRow = Polling1.Dimension.End.Row;

                    if (CompareDataDict.TryGetValue(addrowskey, out int addrowsvalue))
                    {
                        // 刪除最後一列 (總表若延伸至最大列數1,048,576)
                        int PollingStartCol = AllStartCol;
                        if (LastRow >= 1048576)
                        {
                            Polling1.DeleteRow(LastRow);
                        }

                        // 新增列放置最後一列
                        Polling1.InsertRow(AllEndRow, 1);

                        // 設置每一列的高度
                        //Polling1.Row(AllEndRow).Height = 20;

                        // 將資料插入新的一列
                        for (int AddCol = 1; AddCol <= CompareSheet.Dimension.End.Column; AddCol++)
                        {
                            if (AddCol >= StartCol && AddCol < EndCol && PollingStartCol >= AllStartCol && PollingStartCol < AllEndCol)
                            {
                                if (!string.IsNullOrEmpty(CompareSheet.Cells[addrowsvalue, AddCol].Formula))
                                {
                                    string formula = TemplateSheet.Cells[AllStartRow, PollingStartCol].Formula;

                                    // 使用正則表達式替換列號
                                    Polling1.Cells[AllEndRow, PollingStartCol].Formula = Regex.Replace(formula, $@"(\D+){AllStartRow}\b", match => $"{match.Groups[1].Value}{AllEndRow}");
                                }
                                else
                                {
                                    Polling1.Cells[AllEndRow, PollingStartCol].Value = GetCellValue(CompareSheet.Cells[addrowsvalue, AddCol]);
                                }

                                // 設置文字置中
                                if (StartRow != 0 && addrowsvalue >= StartRow && UpdateStartCol != 0 && AddCol >= UpdateStartCol)
                                {
                                    Polling1.Cells[AllEndRow, PollingStartCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                }

                                // 設置數字格式
                                Polling1.Cells[AllEndRow, PollingStartCol].Style.Numberformat.Format = CompareSheet.Cells[addrowsvalue, AddCol].Style.Numberformat.Format;

                                // 設置框線
                                SetBorder(Polling1, AllEndRow, PollingStartCol);

                                // 設置背景色
                                Polling1.Cells[AllEndRow, PollingStartCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                Polling1.Cells[AllEndRow, PollingStartCol].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(204, 153, 255));

                                PollingStartCol++;
                            }
                        }
                    }
                }
            }

            // 判斷是否有刪除列
            if (delrows.Count > 0)
            {
                foreach (var delrowskey in delrows)
                {
                    if (TemplateDataDict.TryGetValue(delrowskey, out int delrowsvalue))
                    {
                        for (int DelCol = 1; DelCol <= Polling1.Dimension.End.Column; DelCol++)
                        {
                            if (DelCol >= AllStartCol && DelCol < AllEndCol)
                            {
                                // 設置背景色
                                Polling1.Cells[delrowsvalue, DelCol].Style.Fill.PatternType = ExcelFillStyle.None;
                                Polling1.Cells[delrowsvalue, DelCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                Polling1.Cells[delrowsvalue, DelCol].Style.Fill.BackgroundColor.SetColor(Color.DimGray);

                                // 新增刪除線
                                Polling1.Cells[delrowsvalue, DelCol].Style.Font.Strike = true;

                                // 設置框線
                                SetBorder(Polling1, delrowsvalue, DelCol);
                            }
                        }
                    }
                }
            }
        }

        private void CombinePolling(ExcelWorksheet Polling2)
        {
            // 計算完所有公式後取值
            Polling2.Workbook.Calculate();

            int PollingMaxCol = Polling2.Dimension.End.Column;
            int PollingMaxRow = Polling2.Dimension.End.Row;

            int StartCol = 0;                               // 取得欄 Start 位置
            int EndCol = 0;                                 // 取得欄 End 位置
            int StartRow = 0;                               // 取得列 起始讀取 位置
            int EndRow = 0;                                 // 預設列 結束讀取 位置
            int EndIndex = 0;                               // EndIndex
            List<int> RemoveColList = new List<int>();      // 存放要移除的欄(1S1U、2S1U)
            List<int> SocketColList = new List<int>();      // 存放要合併的欄(1S2U、2S2U)

            for (int row = 1; row <= PollingMaxRow; row++)
            {
                int FtIndex = 0;                            // 設置 Function Team 欄位引數
                int ODMIndex = 0;                           // 設置 ODMTotal 欄位引述
                int SocketIndex = 0;                        // 設置 合併Socket 欄位引數

                for (int col = 1; col <= PollingMaxCol; col++)
                {
                    var CellValue = Polling2.Cells[row, col].Text;

                    if (CellValue == "Start")
                    {
                        StartCol = col + 1;                 // 錨點 Start 下一個欄位置
                        StartRow = row + 1;                 // 錨點 Start 下一個列位置
                    }
                    else if (CellValue == "End")
                    {
                        if (EndIndex == 0)
                        {
                            EndCol = col;                   // 錨點 End 欄位置
                            EndIndex++;
                        }
                        else if (EndIndex != 0 && EndCol != 0 && col < EndCol)
                        {
                            EndRow = row;                   // 錨點 End 列位置
                        }
                    }

                    if (row >= ComStartRow && EndRow == 0 && col >= AllStartCol && col < AllEndCol)
                    {
                        // 取得 Function Team & ODM Total & 合併Socket 的欄
                        int FtStartCol = (StartList.Count > 0) ? StartList[FtIndex] : 0;
                        int FtEndCol = (EndList.Count > 0) ? EndList[FtIndex] : 0;
                        int OdmStartCol = (OdmStartColList.Count > 0) ? OdmStartColList[ODMIndex] : 0;
                        int OdmEndCol = (OdmEndColList.Count > 0) ? OdmEndColList[ODMIndex] : 0;
                        int OdmStartRow = (OdmStartRowList.Count > 0) ? OdmStartRowList[ODMIndex] : 0;
                        int OdmEndRow = (OdmEndRowList.Count > 0) ? OdmEndRowList[ODMIndex] : 0;
                        int SocketCol = (SocketColList.Count > 0) ? SocketColList[SocketIndex] : 0;

                        if (row > ComEndRow && SocketColList.Count > 0 && col == SocketCol)
                        {
                            // 取出欄位數值(含#、公式計算完的值)
                            string? SocketValue1 = Polling2.Cells[row, col].Value?.ToString(); ;
                            string? SocketValue2 = Polling2.Cells[row, col + 1].Value?.ToString();

                            // 取出並判斷欄位的背景色
                            var SocketColor1 = Polling2.Cells[row, col].Style.Fill.BackgroundColor;
                            var SocketColor2 = Polling2.Cells[row, col + 1].Style.Fill.BackgroundColor;

                            if ((SocketColor1.Rgb != Color.FromArgb(204, 153, 255).ToArgb().ToString("X8") || SocketColor2.Rgb != Color.FromArgb(204, 153, 255).ToArgb().ToString("X8")) && (SocketColor1.Rgb != Color.DimGray.ToArgb().ToString("X8") || SocketColor2.Rgb != Color.DimGray.ToArgb().ToString("X8")))
                            {
                                Polling2.Cells[row, col].Clear();

                                // 設置文字水平和垂直居中
                                Polling2.Cells[row, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                Polling2.Cells[row, col].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                                // 設置框線樣式
                                Polling2.Cells[row, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;      // 上框線
                                Polling2.Cells[row, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;   // 下框線
                                Polling2.Cells[row, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;     // 左框線
                                Polling2.Cells[row, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;    // 右框線

                                // 設置框線顏色為黑色
                                Polling2.Cells[row, col].Style.Border.Top.Color.SetColor(Color.Black);
                                Polling2.Cells[row, col].Style.Border.Bottom.Color.SetColor(Color.Black);
                                Polling2.Cells[row, col].Style.Border.Left.Color.SetColor(Color.Black);
                                Polling2.Cells[row, col].Style.Border.Right.Color.SetColor(Color.Black);

                                // 設置文字水平和垂直居中
                                Polling2.Cells[row, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                Polling2.Cells[row, col].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                                // 設置背景色為透明色
                                Polling2.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.None;
                            }


                            // 判定並合併欄位數值
                            if (int.TryParse(SocketValue1, out int SocketNum1) && int.TryParse(SocketValue2, out int SocketNum2))
                            {
                                // 若兩個欄位為數值，則進行加總
                                Polling2.Cells[row, col].Value = SocketNum1 + SocketNum2;
                            }
                            else if (int.TryParse(SocketValue1, out SocketNum1) && string.IsNullOrEmpty(SocketValue2))
                            {
                                // 若其中一個欄位為數值，則回填數值
                                Polling2.Cells[row, col].Value = SocketNum1;
                            }
                            else if (string.IsNullOrEmpty(SocketValue1) && int.TryParse(SocketValue2, out SocketNum2))
                            {
                                // 若其中一個欄位為數值，則回填數值
                                Polling2.Cells[row, col].Value = SocketNum2;
                            }
                            else if (SocketValue1 != null && SocketValue1.StartsWith('#') && string.IsNullOrEmpty(SocketValue2))
                             {
                                // 若第一欄位為有 # 符號、第二欄位為空，則回填第一欄位文字
                                Polling2.Cells[row, col].Value = SocketValue1;
                            }
                            else if (string.IsNullOrEmpty(SocketValue1) && SocketValue2 != null && SocketValue2.StartsWith("#"))
                            {
                                // 若第二欄位為有 # 符號、第一欄位為空，則回填第二欄位文字
                                Polling2.Cells[row, col].Value = SocketValue2;
                            }
                            else if (int.TryParse(SocketValue1, out SocketNum1) && SocketValue2 != null && SocketValue2.StartsWith("#"))
                            {
                                // 若第一欄位為數值、第二欄位為有 # 符號，則忽略第二欄位而回填第一欄位數值
                                Polling2.Cells[row, col].Value = SocketNum1;
                            }
                            else if (SocketValue1 != null && SocketValue1.StartsWith("#") && int.TryParse(SocketValue2, out SocketNum2))
                            {
                                // 若第一欄位為數值、第二欄位為有 # 符號，則忽略第二欄位而回填第一欄位數值
                                Polling2.Cells[row, col].Value = SocketNum2;
                            }

                            if (SocketIndex < SocketColList.Count - 1)
                            {
                                SocketIndex++;
                            }
                        }

                        // 標題欄位設置背景色
                        if (row <= ComEndRow)
                        {
                            if (row >= SURow && StartList.Count > 0 && EndList.Count > 0 && FtStartCol != 0 && col >= FtStartCol && (FtEndCol == 0 || (FtEndCol != 0 && col <= FtEndCol)))
                            {
                                // 檢查並替換 "1S1U" 為 "1S"、"2S1U" 為 "2S"
                                if (CellValue.StartsWith("1S1U") || (CellValue.StartsWith("2S1U")))
                                {
                                    if (CellValue.StartsWith("1S1U"))
                                    {
                                        Polling2.Cells[row, col].Value = "1S";
                                    }
                                    else if (CellValue.StartsWith("2S1U"))
                                    {
                                        Polling2.Cells[row, col].Value = "2S";
                                    }
                                    // 儲存需合併的欄
                                    SocketColList.Add(col);
                                }
                                else if (!RemoveColList.Contains(col) && (CellValue.StartsWith("1S2U") || CellValue.StartsWith("2S2U")))
                                {
                                    // 儲存需移除的欄
                                    RemoveColList.Add(col);
                                }

                                // 先取出原欄位數值
                                var OriginalValue = Polling2.Cells[row, col].Value;

                                Polling2.Cells[row, col].Clear();
                                // 設置框線樣式
                                Polling2.Cells[row, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;      // 上框線
                                Polling2.Cells[row, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;   // 下框線
                                Polling2.Cells[row, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;     // 左框線
                                Polling2.Cells[row, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;    // 右框線

                                // 設置框線顏色為黑色
                                Polling2.Cells[row, col].Style.Border.Top.Color.SetColor(Color.Black);
                                Polling2.Cells[row, col].Style.Border.Bottom.Color.SetColor(Color.Black);
                                Polling2.Cells[row, col].Style.Border.Left.Color.SetColor(Color.Black);
                                Polling2.Cells[row, col].Style.Border.Right.Color.SetColor(Color.Black);

                                // 設置文字顏色為白色
                                Polling2.Cells[row, col].Style.Font.Color.SetColor(Color.White);

                                // 設置文字為粗體
                                Polling2.Cells[row, col].Style.Font.Bold = true;

                                // 設置文字水平和垂直居中
                                Polling2.Cells[row, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                Polling2.Cells[row, col].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                                Polling2.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.None;
                                Polling2.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                Polling2.Cells[row, col].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(204, 153, 255));

                                // 回填欄位值
                                Polling2.Cells[row, col].Value = OriginalValue;

                                if (FtEndCol != 0 && FtEndCol == col && FtIndex < StartList.Count - 1)
                                {
                                    FtIndex++;
                                }
                            }

                            if (StartList.Count > 0 && OdmStartColList.Count > 0 && OdmEndColList.Count > 0 && OdmStartRowList.Count > 0)
                            {
                                if (OdmStartRow != 0 && FtStartCol != 0 && OdmStartCol != 0 && OdmEndCol != 0 && row >= OdmStartRow && col >= FtStartCol && col >= OdmStartCol && col <= OdmEndCol)
                                {
                                    Polling2.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.None;
                                    Polling2.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    Polling2.Cells[row, col].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(204, 153, 255));

                                    if (OdmEndCol != 0 && OdmEndCol == col && ODMIndex < OdmStartColList.Count - 1)
                                    {
                                        ODMIndex++;
                                    }
                                }
                            }
                        }
                    }
                }
                // 如果EndRow有值，則跳出迴圈
                if (EndRow != 0)
                {
                    break;
                }
            }
            // 移除儲存的欄
            int RemoveIndex = 0;
            foreach (int RemoveCol in RemoveColList)
            {
                Polling2.DeleteColumn(RemoveCol - RemoveIndex);
                RemoveIndex++;
            }
        }

        private object GetCellValue(ExcelRange cell)
        {
            // 如果欄位存在公式，則以公式計算結果，否則以欄位值為結果
            if (!string.IsNullOrEmpty(cell.Formula))
            {
                if (double.TryParse(cell.GetValue<string>(), out double result))
                {
                    return result;
                }
                else
                {
                    // 無法轉換數值時則使用文字
                    return cell.Text;
                }
            }
            else
            {
                return cell.Value;
            }
        }

        private void SetBorder(ExcelWorksheet Polling1, int row, int col)
        {
            // 設置框線樣式
            Polling1.Cells[row, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            Polling1.Cells[row, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            Polling1.Cells[row, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            Polling1.Cells[row, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            // 設置框線顏色為黑色
            Polling1.Cells[row, col].Style.Border.Top.Color.SetColor(Color.Black);
            Polling1.Cells[row, col].Style.Border.Bottom.Color.SetColor(Color.Black);
            Polling1.Cells[row, col].Style.Border.Left.Color.SetColor(Color.Black);
            Polling1.Cells[row, col].Style.Border.Right.Color.SetColor(Color.Black);
        }

        private string ChangeSymbol(string formula)
        {
            return formula.Replace("+", ",");
        }

        // 將處理後的 Excel 檔案另存為新檔
        private void SaveAsNewExcel(ExcelPackage package)
        {
            // 設定 EPPlus 授權模式，非商業用途
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel Files|*.xlsx";
                saveFileDialog.Title = "另存為新的 Excel 文件";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    FileInfo fileInfo = new FileInfo(saveFileDialog.FileName);
                    package.SaveAs(fileInfo);
                    MessageBox.Show("檔案已保存至: " + saveFileDialog.FileName);
                }
            }
        }

        // 清除暫存變數
        private void ClearCache()
        {
            // 清除 TemplatePath 的暫存
            if (!string.IsNullOrEmpty(TemplatePath))
            {
                TemplatePath = null;
                NonFileFirst.Text = "(尚未上傳檔案！)";
                NonFileLast.Text = "(尚未上傳檔案！)";
            }

            // 清除 MultipleFilesPaths 的暫存
            if (MultipleFilesPaths != null && MultipleFilesPaths.Count > 0)
            {
                MultipleFilesPaths?.Clear();
                NonFIleMuti.Text = "(尚未上傳檔案！)";
            }

            BlackBGCol = 0;
            AllStartCol = 0;
            AllStartRow = 0;
            AllEndCol = 0;
            AllEndRow = 0;
            ComStartRow = 0;
            ComEndRow = 0;
            SURow = 0;
            TempColCount = 0;
            TemplateRepeatData = null;
            CompareRepeatData = null;

            RYGList?.Clear();
            StartList?.Clear();
            EndList?.Clear();
            OdmStartColList?.Clear();
            OdmEndColList?.Clear();
            OdmStartRowList?.Clear();
            OdmEndRowList?.Clear();
            TemplateDataDict?.Clear();
            TempAllData?.Clear();
            TempAllColData?.Clear();

            // 手動回收記憶體
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
    }
}
