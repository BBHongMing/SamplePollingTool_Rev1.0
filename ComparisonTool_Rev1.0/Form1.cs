
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
        private List<string>? MultipleFilesPaths = new List<string>();   // �إߤW���ɮײM��

        // �إ߽d�����ƾڶ��X
        private int BlackBGCol = 0;                         // RTS1.0_BCT RTS1.1_CT��m(�I����ߤ@�¦�)
        private int AllStartCol = 0;                        // �_�lŪ�����m
        private int AllStartRow = 0;                        // �_�lŪ���C��m
        private int AllEndCol = 0;                          // ����Ū�����m
        private int AllEndRow = 0;                          // ����Ū���C��m
        private int ComStartRow = 0;                        // �_�lCombineŪ���C��m
        private int ComEndRow = 0;                          // ����CombineŪ���C��m
        private int SURow = 0;                              // SUŪ���C��m
        private int TempColCount = 0;                       // �d�����`�ƶq
        private string? TemplateRepeatData = null;          // �d�����ƪ��C
        private string? CompareRepeatData = null;           // ����ɮ׭��ƪ��C

        private List<int> RYGList = new List<int> { };          // ���o�� RYG ��m
        private List<int> StartList = new List<int> { };        // ���oFuction Team Ū���_�l��m
        private List<int> EndList = new List<int> { };          // ���oFuction Team Ū��������m
        private List<int> OdmStartColList = new List<int> { };  // ���o ODM �_�l�� ��m�M��
        private List<int> OdmEndColList = new List<int> { };    // ���o ODM ������ ��m�M��
        private List<int> OdmStartRowList = new List<int> { };  // ���o ODM �_�l�C ��m�M��
        private List<int> OdmEndRowList = new List<int> { };    // ���o ODM �����C ��m�M��
        private List<string> GroupList = new List<string>       // �w�] Title Group
        {
            "Category", "DPN", "Description", "Manufacturers", "Remark"
        };

        private Dictionary<string, int> TemplateDataDict = new Dictionary<string, int>();   // ���o�d�� Group ��
        private Dictionary<int, string> TempAllData = new Dictionary<int, string>();        // ���o�d�� Group �᭱�U��ƭ�
        private Dictionary<int, string> TempAllColData = new Dictionary<int, string>();     // ���o�d�� Group �᭱�U���m

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // �w�]�襤�����ϥ�
            FirstUse.Checked = true;
            UploadFirstTemplateBtn.Visible = true;
            NonFileFirst.Visible = true;
            UploadLastTemplateBtn.Visible = false;
            NonFileLast.Visible = false;
        }

        // �����W�ǽd����Radio ButtonĲ�o�ƥ�
        private void FirstUseChecked(object sender, EventArgs e)
        {
            if (FirstUse.Checked)
            {
                // �����ϥνd��
                UploadFirstTemplateBtn.Visible = true;
                NonFileFirst.Visible = true;
                UploadLastTemplateBtn.Visible = false;
                NonFileLast.Visible = false;
                TemplatePath = null; // �M�Źw��d��
                NonFileFirst.Text = "(�|���W���ɮסI)";
            }
        }

        // �e���ϥνd����Radio ButtonĲ�o�ƥ�
        private void LastUseChecked(object sender, EventArgs e)
        {
            if (LastUse.Checked)
            {
                // �e���ϥνd��
                UploadFirstTemplateBtn.Visible = false;
                NonFileFirst.Visible = false;
                UploadLastTemplateBtn.Visible = true;
                NonFileLast.Visible = true;
                TemplatePath = null; // �M�Źw��d��
                NonFileLast.Text = "(�|���W���ɮסI)";
            }
        }

        // �����ϥνd���W�ǫ��s
        private void UploadFirstTemplate(object sender, EventArgs e)
        {
            // �]�w EPPlus ���v�Ҧ��A�D�ӷ~�γ~
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx";
                openFileDialog.Title = "�п�ܪťսd���ɮ�!";
                openFileDialog.Multiselect = false;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    TemplatePath = openFileDialog.FileName;
                    MessageBox.Show("�d���ɮפw���: " + TemplatePath);
                    NonFileFirst.Text = "(�ɮפw�W�ǡI)";
                }
                else
                {
                    MessageBox.Show("�п�ܪťսd���ɮסC");
                    return;
                }
            }
        }

        // �e���ϥΫD�ťսd���W�ǫ��s
        private void UploadLastTemplate(object sender, EventArgs e)
        {
            // �]�w EPPlus ���v�Ҧ��A�D�ӷ~�γ~
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx";
                openFileDialog.Title = "�п�ܫD�ťսd���ɮ�!";
                openFileDialog.Multiselect = false;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    TemplatePath = openFileDialog.FileName;
                    MessageBox.Show("�d���ɮפw���: " + TemplatePath);
                    NonFileLast.Text = "(�ɮפw�W�ǡI)";
                }

                if (TemplatePath == null)
                {
                    MessageBox.Show("�п�ܫD�ťսd���ɮסC");
                    return;
                }
            }
        }

        // �h���ɮפW�ǫ��s
        private void UploadMultipleFile(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx";
                openFileDialog.Title = "�п�ܦh���ɮ�!";
                openFileDialog.Multiselect = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    MultipleFilesPaths = new List<string>(openFileDialog.FileNames);
                    MessageBox.Show($"�w��� {MultipleFilesPaths.Count} ���ɮסC");
                    NonFIleMuti.Text = "(�ɮפw�W�ǡI)";
                }

                if (MultipleFilesPaths == null || MultipleFilesPaths.Count == 0)
                {
                    MessageBox.Show("�п�ܦܤ֤@���ɮסC");
                }
            }
        }


        // �e�X���s
        private void SummitBtn(object sender, EventArgs e)
        {
            if (TemplatePath == null)
            {
                MessageBox.Show("�Х��W�ǽd���ɮסC");
                return;
            }

            if (MultipleFilesPaths == null || MultipleFilesPaths.Count == 0)
            {
                MessageBox.Show("�п�ܦh�� Excel �ɮ׶i����C");
                return;
            }

            // �}�l�B�z Excel ���
            CombineExcelFiles(TemplatePath, MultipleFilesPaths);

            // �M�żȦs���
            ClearCache();
        }

        private void CombineExcelFiles(string TempPath, List<string> filePaths)
        {
            try
            {
                // �H�d���ɮ׬���¦
                using (ExcelPackage TemplatePackage = new ExcelPackage(new FileInfo(TempPath)))
                {
                    ExcelWorksheet TemplateSheet = TemplatePackage.Workbook.Worksheets[0];
                    string TempName = TemplatePackage.File.Name;

                    using (ExcelPackage newPackage = new ExcelPackage())
                    {
                        // �ƻs�d���u�@���s�ɮ�
                        ExcelWorksheet Polling1 = newPackage.Workbook.Worksheets.Add("Polling1", TemplateSheet);

                        try
                        {
                            GetTemplateData(Polling1, TemplateSheet);  //���o�d����� Category�� Remark Group 
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Ū���d����Ʈɵo�Ϳ��~�A���ˬd�d���ɮסG{TempName}\n���~�T���G{ex.Message}", "���~", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        if (!string.IsNullOrEmpty(TemplateRepeatData))
                        {
                            TemplateRepeatData += "�����Ƹ�ơA���ˬd�d���ɮסG" + TempName;
                            MessageBox.Show(TemplateRepeatData);
                            return;
                        }

                        
                        // �q C ��}�l�O���t�����
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
                                        MessageBox.Show($"Ū������ɮ׮ɵo�Ϳ��~�A���ˬd�ɮסG{FileName}\n���~�T���G{ex.Message}", "���~", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        return;
                                    }

                                    if (!string.IsNullOrEmpty(CompareRepeatData))
                                    {
                                        CompareRepeatData += "�����Ƹ�ơA���ˬd����ɮסG" + FileName;
                                        MessageBox.Show(CompareRepeatData);
                                        return;
                                    }
                                    colIndex++;
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Ū������ɮ׮ɵo�Ϳ��~�A�ɮ׸��|���G{filePath}\n���~�T���G{ex.Message}", "���~", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                        }
                        
                        try
                        {
                            // �ƻs Polling1 ���e�� Polling2
                            ExcelWorksheet Polling2 = newPackage.Workbook.Worksheets.Add("Polling2", Polling1);
                            CombinePolling(Polling2);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"�X�ָ�Ʈɵo�Ϳ��~�G\n���~�T���G{ex.Message}", "���~", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        try
                        {
                            // �x�s�s�� Excel �ɮ�
                            SaveAsNewExcel(newPackage);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"�t�s�s�ɮɵo�Ϳ��~�G\n���~�T���G{ex.Message}", "���~", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ū���d���ɮ׵o�Ϳ��~�A���ˬd�d���ɮ�!\n�ɮ׸��|�G{TempPath}\n���~�T���G{ex.Message}", "���~", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void GetTemplateData(ExcelWorksheet Polling1, ExcelWorksheet TemplateSheet)
        {
            int TemplateMaxRow = TemplateSheet.Dimension.End.Row;
            int TemplateMaxCol = TemplateSheet.Dimension.End.Column;
            int StartTitleCol = 0;      // ���o�� Category ��m
            int StartTitleRow = 0;      // ���o�C Category ��m
            int EndTitleCol = 0;        // ���o�� Remark ��m
            int UpdateStartCol = 0;     // ���o�� Revision ��m
            int StartRow = 0;           // ���o�C �_�l ��m
            int RemoveCol = 0;          // ���o�� Remove ��m
            int RemarkIndex = 0;        // RemarkIndex
            int RevisionIndex = 0;      // RevisionIndex
            int SUIndex = 0;            // SUIndex
            int EndIndex = 0;           // EndIndex

            List<int> TitleList = new List<int> { };  // ���o�� Group ��m

            // �����ƻs�d���z��\��
            Polling1.Cells.AutoFilter = false;

            for (int row = 1; row <= TemplateMaxRow; row++)
            {
                // �����ƻs�d���s��(�C)
                Polling1.Row(row).OutlineLevel = 0;

                // �����ƻs�d�����æC
                if (Polling1.Row(row).Hidden)
                {
                    Polling1.Row(row).Hidden = false;
                }

                string RowData = "";
                List<string> TempData = new List<string>();
                List<int> TempColData = new List<int>();
                int RemoveSign = 0;  // Remove���аO

                for (int col = 1; col <= TemplateMaxCol; col++)
                {
                    // �����ƻs�d���s��(��)
                    Polling1.Column(col).OutlineLevel = 0;

                    // �����ƻs�d��������
                    if (Polling1.Column(col).Hidden)
                    {
                        Polling1.Column(col).Hidden = false;
                    }

                    // ���o�C�����ƭ�
                    var CellValue = TemplateSheet.Cells[row, col].Text;

                    if (CellValue == "Start")
                    {
                        AllStartCol = col + 1;      // ���D�� Start ��m
                        AllStartRow = row + 3;      // ���D�� Start ��m
                        ComEndRow = row;            // Polling2 �������D�C��m
                        StartRow = row + 1;         // �_�lŪ���C��m
                    }
                    else if (CellValue == "Remove" && StartTitleCol == 0)
                    {
                        RemoveCol = col;            // ���D�� Remove ��m
                    }
                    else if (CellValue == "Category")
                    {
                        StartTitleCol = col;        // ���D�� Category ��m
                        StartTitleRow = row;
                        ComStartRow = row;          // Polling2 �_�l���D�C��m
                    }
                    else if (CellValue == "Remark" && row == StartTitleRow && RemarkIndex == 0)
                    {
                        EndTitleCol = col;          // ���D�� Remark ��m
                        RemarkIndex++;
                    }
                    else if (CellValue == "Revision")
                    {
                        if (RevisionIndex == 0)
                        {
                            UpdateStartCol = col;   // Revision���_�lŪ�����
                            RevisionIndex++;
                        }
                        StartList.Add(col + 1);
                    }
                    else if (CellValue.StartsWith("1S1U") && SUIndex == 0)
                    {
                        SURow = row;                // SUŪ���C��m
                        SUIndex++;
                    }
                    else if (CellValue.Replace(" ", "").Replace("\r", "").Replace("\n", "") == "1SODMTotal")
                    {
                        EndList.Add(col - 1);       // �s�WFunction Team Ū����������m
                        OdmStartColList.Add(col);   // �s�W ODM Total Ū���_�l����m
                        OdmStartRowList.Add(row);   // �s�W ODM Total Ū���_�l�C��m
                    }
                    else if (CellValue.Replace(" ", "").Replace("\r", "").Replace("\n", "") == "1S+2SODMTotal")
                    {
                        OdmEndColList.Add(col);
                        OdmEndRowList.Add(row);
                    }
                    else if (CellValue.Replace(" ", "").Replace("\r", "").Replace("\n", "") == "RTS1.0_BCTRTS1.1_CT")
                    {
                        BlackBGCol = col;           // ���D�� 0_BCTRTS1 ��m
                    }
                    else if (CellValue == "End")
                    {
                        if (EndIndex == 0)
                        {
                            AllEndCol = col;        // ����Ū�����m
                            TempColCount = (AllEndCol - 1) - AllStartCol + 1; // �d�����`�ƶq
                            EndIndex++;
                        }
                        else if (EndIndex != 0 && AllEndCol != 0 && col < AllEndCol)
                        {
                            AllEndRow = row;        // ����Ū���C��m
                        }
                    }
                    else if (CellValue == "RYG" && StartRow == 0 && AllEndCol == 0)
                    {
                        RYGList.Add(col);           // ���D�� RYG ��m
                    }

                    // �NTitle����[�JTitleList
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

                    // Remove��즳�ȮɡA�аORemoveSign
                    if (RemoveCol != 0 && col == RemoveCol && !string.IsNullOrEmpty(CellValue))
                    {
                        RemoveSign = 1;
                    }

                    
                    if (StartRow != 0 && row >= StartRow && AllStartCol != 0 && AllEndCol != 0 && col >= AllStartCol && col < AllEndCol && AllEndRow == 0)
                    {
                        ResetBackgroundColor(Polling1, TemplateSheet, row, col, RemoveSign);
                    }
                }

                // �u�x�s��Ƥ��e�A���]�A�C��
                if (!string.IsNullOrEmpty(RowData))
                {
                    //�P�_�O�_�����ƪ�Key
                    if (!TemplateDataDict.ContainsKey(RowData))
                    {
                        TemplateDataDict.Add(RowData, row); // �N�C����@Key
                    }
                    else
                    {
                        TemplateRepeatData += "��" + row + "�C, ";
                    }
                }

                if (StartRow != 0 && row >= StartRow && UpdateStartCol != 0 && AllEndRow == 0)
                {
                    // �X���x�s��
                    string AllRowData = string.Join("|", TempData) + "|";
                    string AllColData = string.Join("|", TempColData) + "|";
                    TempAllData[row] = AllRowData;
                    TempAllColData[row] = AllColData;
                }

                // �p�GAllEndRow���ȡA�h���X�j��
                if (AllEndRow != 0)
                {
                    break;
                }
            }
        }

        private void ResetBackgroundColor(ExcelWorksheet Polling1, ExcelWorksheet TemplateSheet, int row, int col, int RemoveSign)
        {
            // �����X�����ƭ�
            object OriginalValue = GetCellValue(TemplateSheet.Cells[row, col]);

            // �P�_�O�_��RYG��
            if (!RYGList.Contains(col))
            {
                // Start~End�Ҧ��x�s�歫�m�I����
                Polling1.Cells[row, col].Clear();
                Polling1.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.None;

                // �p�GRemove�榳�ȫh�]�m�Ǧ�
                if (RemoveSign == 1)
                {
                    Polling1.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Polling1.Cells[row, col].Style.Fill.BackgroundColor.SetColor(Color.DimGray);

                    // �s�W�R���u
                    Polling1.Cells[row, col].Style.Font.Strike = true;
                }
            }

            // �^������
            if (!string.IsNullOrEmpty(TemplateSheet.Cells[row, col].Formula))
            {
                // �ϥΥ��h��F�������C��
                Polling1.Cells[row, col].Formula = TemplateSheet.Cells[row, col].Formula;
            }
            else
            {
                Polling1.Cells[row, col].Value = OriginalValue;
            }

            // �]�m��r���¦�
            Polling1.Cells[row, col].Style.Font.Color.SetColor(Color.Black);

            // �]�m��r�m��
            Polling1.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            // �]�m�Ʀr�榡
            Polling1.Cells[row, col].Style.Numberformat.Format = TemplateSheet.Cells[row, col].Style.Numberformat.Format;

            // �]�m�ؽu
            SetBorder(Polling1, row, col);
        }

        private void CompareSheets(ExcelWorksheet Polling1, ExcelWorksheet CompareSheet, ExcelWorksheet TemplateSheet, string FileName, int colIndex)
        {
            // �����z��\��
            CompareSheet.Cells.AutoFilter = false;

            int TemplateMaxRow = TemplateSheet.Dimension.End.Row;
            int CompareMaxRow = CompareSheet.Dimension.End.Row;
            int maxCol = Math.Max(TemplateSheet.Dimension.End.Column, CompareSheet.Dimension.End.Column);

            int RemoveCol = 0;          // ���o�� Remove ��m
            int StartCol = 0;           // ���o�� Start ��m
            int StartRow = 0;           // ���o�C �_�l ��m
            int StartTitleCol = 0;      // ���o�� Category ��m
            int StartTitleRow = 0;      // ���o�C Category ��m
            int EndTitleCol = 0;        // ���o�� Remark ��m
            int UpdateStartCol = 0;     // ���o�� Revision ��m
            int UpdateEndCol = 0;       // ���o�� 1S ODM Total ��m
            int EndCol = 0;             // ���o�� End ��m
            int EndRow = 0;             // ���o�C End ��m
            int CompareColCount = 0;    // ����ɮ����`�ƶq
            int RemarkIndex = 0;        // RemarkIndex
            int RevisionIndex = 0;      // RevisionIndex
            int EndIndex = 0;           // EndIndex
            Dictionary<string, int> CompareDataDict = new Dictionary<string, int>();    // ���o����ɮ�Group��
            List<int> TitleList = new List<int> { };  // ���o�� Group ��m
            int SeparatedDataIndex = 0;

            // ���C�@�C
            for (int row = 1; row <= CompareMaxRow; row++)
            {
                // �]�mRemove Index��0
                int RemoveIndex = 0;

                // �]�m Category�� Remark ���X
                string CompareData = "";

                // �H'|'���Ϋ᪺�}�C�ŧi
                string[] SeparatedData = Array.Empty<string>();
                int SeparatedColIndex = 0;
                int IndexStatus = 0;

                // ���C�@��
                for (int col = 1; col <= maxCol; col++)
                {
                    string cellAddress = TemplateSheet.Cells[row, col].Address;
                    var CellCombineTitle = CompareSheet.Cells[row, col].Text;
                    var CellTempTitle = TemplateSheet.Cells[row, col].Text;

                    if (CellCombineTitle == "Start")
                    {
                        StartCol = col + 1;         // ���D�� Start ��m
                        StartRow = row + 1;         // �_�lŪ���C��m
                    }
                    else if (CellCombineTitle == "Remove" && StartTitleCol == 0)
                    {
                        RemoveCol = col;            // ���D�� Remove ��m
                    }
                    else if (CellCombineTitle == "Category")
                    {
                        StartTitleCol = col;        // ���D�� Category ��m
                        StartTitleRow = row;
                    }
                    else if (CellCombineTitle == "Remark" && row == StartTitleRow && RemarkIndex == 0)
                    {
                        EndTitleCol = col;           // ���D�� Remark ��m
                        RemarkIndex++;
                    }
                    else if (CellCombineTitle == "Revision" && RevisionIndex == 0)
                    {
                        UpdateStartCol = col;   // Revision���_�lŪ�����
                        RevisionIndex++;
                    }
                    else if (CellCombineTitle.Replace(" ", "").Replace("\r", "").Replace("\n", "") == "1SODMTotal")
                    {
                        UpdateEndCol = col;         // Function Team��g��������@�����
                    }
                    else if (CellCombineTitle == "End")
                    {
                        if (EndIndex == 0)
                        {
                            EndCol = col;           // ����Ū�����m
                            CompareColCount = (EndCol - 1) - StartCol + 1; // ����ɮ����`�ƶq
                            EndIndex++;
                        }
                        else if (EndIndex != 0 && EndCol != 0 && col < EndCol)
                        {
                            EndRow = row;           // ����Ū���C��m
                        }
                    }

                    if (CompareColCount != 0 && CompareColCount != TempColCount)
                    {
                        MessageBox.Show($"����ɮ׻P�d���ɮ����ƶq���šA���ˬd����ɮסG{FileName}", "���~", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // �NTitle����[�JTitleList
                    if (!TitleList.Contains(col) && GroupList.Contains(CellCombineTitle) && row == StartTitleRow)
                    {
                        TitleList.Add(col);
                    }

                    // �x�s�ݤ����쪺�ƾڶ��X
                    if (StartTitleCol != 0 && EndTitleCol != 0 && StartRow != 0 && col >= StartTitleCol && col <= EndTitleCol && row >= StartRow && EndRow == 0 && TitleList.Contains(col))
                    {
                        CompareData += (col == StartTitleCol) ? CompareSheet.Cells[row, col].Text.Replace(" ", "") + "|" : CompareSheet.Cells[row, col].Text.Trim() + "|";
                        if (col == EndTitleCol)
                        {
                            //�P�_�O�_�����ƪ�Key
                            if (!CompareDataDict.ContainsKey(CompareData))
                            {
                                CompareDataDict.Add(CompareData, row);
                            }
                            else
                            {
                                CompareRepeatData += "��" + row + "�C, ";
                            }
                        }
                    }

                    if (StartRow != 0 && row >= StartRow && EndRow == 0)
                    {
                        // �H'|'���Ϋ᪺�}�C�ŧi
                        string[] SeparatedColData = Array.Empty<string>();

                        // �P�wRemove���O�_����
                        if (RemoveIndex == 0 && RemoveCol != 0 && col == RemoveCol && !string.IsNullOrEmpty(CellCombineTitle))
                        {
                            RemoveIndex = 1;
                        }

                        // �P�w����ɮצ��C��ƬO�_�s�b�d����Ƥ�
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

                            // TempAllData �]�t���w�� Key
                            TempAllData.TryGetValue(row, out string? TempRowData);
                            TempAllColData.TryGetValue(row, out string? TempColData);

                            // �H'|'���Φr���x�s�}�C
                            SeparatedData = TempRowData.Split('|');
                            SeparatedColData = TempColData.Split('|');


                            // �����̫᪺�ťդ����]�p�G�r�굲�����h�l�� '|'�^
                            if (SeparatedData.Length > 0 && string.IsNullOrEmpty(SeparatedData[SeparatedData.Length - 1]))
                            {
                                Array.Resize(ref SeparatedData, SeparatedData.Length - 1);
                            }
                            if (SeparatedColData.Length > 0 && string.IsNullOrEmpty(SeparatedColData[SeparatedColData.Length - 1]))
                            {
                                Array.Resize(ref SeparatedColData, SeparatedColData.Length - 1);
                            }

                            // ���o���Ϋ����ƭ�
                            string SeparatedValue = SeparatedData[SeparatedColIndex];
                            int SeparatedCol = int.Parse(SeparatedColData[SeparatedColIndex]);

                            if (!Equals(SeparatedValue, compareValue) && ((string.IsNullOrEmpty(SeparatedValue) && !string.IsNullOrEmpty(compareValue)) || (!string.IsNullOrEmpty(SeparatedValue) && !string.IsNullOrEmpty(compareValue))))
                            {
                                if (string.IsNullOrEmpty(CompareSheet.Cells[RowKey, col].Formula))
                                {
                                    Polling1.Cells[row, SeparatedCol].Value = GetCellValue(CompareSheet.Cells[RowKey, col]);
                                    

                                    // �]�m�Ʀr�榡
                                    Polling1.Cells[row, SeparatedCol].Style.Numberformat.Format = CompareSheet.Cells[RowKey, col].Style.Numberformat.Format;

                                    // �]�m�s�W�B�ק����I����μƭ�
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
            // �P�_�O�_���s�W�C
            if (addrows.Count > 0)
            {
                foreach (var addrowskey in addrows)
                {
                    // ���o�`��̫�@�C
                    int LastRow = Polling1.Dimension.End.Row;

                    if (CompareDataDict.TryGetValue(addrowskey, out int addrowsvalue))
                    {
                        // �R���̫�@�C (�`��Y�����̤ܳj�C��1,048,576)
                        int PollingStartCol = AllStartCol;
                        if (LastRow >= 1048576)
                        {
                            Polling1.DeleteRow(LastRow);
                        }

                        // �s�W�C��m�̫�@�C
                        Polling1.InsertRow(AllEndRow, 1);

                        // �]�m�C�@�C������
                        //Polling1.Row(AllEndRow).Height = 20;

                        // �N��ƴ��J�s���@�C
                        for (int AddCol = 1; AddCol <= CompareSheet.Dimension.End.Column; AddCol++)
                        {
                            if (AddCol >= StartCol && AddCol < EndCol && PollingStartCol >= AllStartCol && PollingStartCol < AllEndCol)
                            {
                                if (!string.IsNullOrEmpty(CompareSheet.Cells[addrowsvalue, AddCol].Formula))
                                {
                                    string formula = TemplateSheet.Cells[AllStartRow, PollingStartCol].Formula;

                                    // �ϥΥ��h��F�������C��
                                    Polling1.Cells[AllEndRow, PollingStartCol].Formula = Regex.Replace(formula, $@"(\D+){AllStartRow}\b", match => $"{match.Groups[1].Value}{AllEndRow}");
                                }
                                else
                                {
                                    Polling1.Cells[AllEndRow, PollingStartCol].Value = GetCellValue(CompareSheet.Cells[addrowsvalue, AddCol]);
                                }

                                // �]�m��r�m��
                                if (StartRow != 0 && addrowsvalue >= StartRow && UpdateStartCol != 0 && AddCol >= UpdateStartCol)
                                {
                                    Polling1.Cells[AllEndRow, PollingStartCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                }

                                // �]�m�Ʀr�榡
                                Polling1.Cells[AllEndRow, PollingStartCol].Style.Numberformat.Format = CompareSheet.Cells[addrowsvalue, AddCol].Style.Numberformat.Format;

                                // �]�m�ؽu
                                SetBorder(Polling1, AllEndRow, PollingStartCol);

                                // �]�m�I����
                                Polling1.Cells[AllEndRow, PollingStartCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                Polling1.Cells[AllEndRow, PollingStartCol].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(204, 153, 255));

                                PollingStartCol++;
                            }
                        }
                    }
                }
            }

            // �P�_�O�_���R���C
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
                                // �]�m�I����
                                Polling1.Cells[delrowsvalue, DelCol].Style.Fill.PatternType = ExcelFillStyle.None;
                                Polling1.Cells[delrowsvalue, DelCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                Polling1.Cells[delrowsvalue, DelCol].Style.Fill.BackgroundColor.SetColor(Color.DimGray);

                                // �s�W�R���u
                                Polling1.Cells[delrowsvalue, DelCol].Style.Font.Strike = true;

                                // �]�m�ؽu
                                SetBorder(Polling1, delrowsvalue, DelCol);
                            }
                        }
                    }
                }
            }
        }

        private void CombinePolling(ExcelWorksheet Polling2)
        {
            // �p�⧹�Ҧ����������
            Polling2.Workbook.Calculate();

            int PollingMaxCol = Polling2.Dimension.End.Column;
            int PollingMaxRow = Polling2.Dimension.End.Row;

            int StartCol = 0;                               // ���o�� Start ��m
            int EndCol = 0;                                 // ���o�� End ��m
            int StartRow = 0;                               // ���o�C �_�lŪ�� ��m
            int EndRow = 0;                                 // �w�]�C ����Ū�� ��m
            int EndIndex = 0;                               // EndIndex
            List<int> RemoveColList = new List<int>();      // �s��n��������(1S1U�B2S1U)
            List<int> SocketColList = new List<int>();      // �s��n�X�֪���(1S2U�B2S2U)

            for (int row = 1; row <= PollingMaxRow; row++)
            {
                int FtIndex = 0;                            // �]�m Function Team ���޼�
                int ODMIndex = 0;                           // �]�m ODMTotal ���ޭz
                int SocketIndex = 0;                        // �]�m �X��Socket ���޼�

                for (int col = 1; col <= PollingMaxCol; col++)
                {
                    var CellValue = Polling2.Cells[row, col].Text;

                    if (CellValue == "Start")
                    {
                        StartCol = col + 1;                 // ���I Start �U�@�����m
                        StartRow = row + 1;                 // ���I Start �U�@�ӦC��m
                    }
                    else if (CellValue == "End")
                    {
                        if (EndIndex == 0)
                        {
                            EndCol = col;                   // ���I End ���m
                            EndIndex++;
                        }
                        else if (EndIndex != 0 && EndCol != 0 && col < EndCol)
                        {
                            EndRow = row;                   // ���I End �C��m
                        }
                    }

                    if (row >= ComStartRow && EndRow == 0 && col >= AllStartCol && col < AllEndCol)
                    {
                        // ���o Function Team & ODM Total & �X��Socket ����
                        int FtStartCol = (StartList.Count > 0) ? StartList[FtIndex] : 0;
                        int FtEndCol = (EndList.Count > 0) ? EndList[FtIndex] : 0;
                        int OdmStartCol = (OdmStartColList.Count > 0) ? OdmStartColList[ODMIndex] : 0;
                        int OdmEndCol = (OdmEndColList.Count > 0) ? OdmEndColList[ODMIndex] : 0;
                        int OdmStartRow = (OdmStartRowList.Count > 0) ? OdmStartRowList[ODMIndex] : 0;
                        int OdmEndRow = (OdmEndRowList.Count > 0) ? OdmEndRowList[ODMIndex] : 0;
                        int SocketCol = (SocketColList.Count > 0) ? SocketColList[SocketIndex] : 0;

                        if (row > ComEndRow && SocketColList.Count > 0 && col == SocketCol)
                        {
                            // ���X���ƭ�(�t#�B�����p�⧹����)
                            string? SocketValue1 = Polling2.Cells[row, col].Value?.ToString(); ;
                            string? SocketValue2 = Polling2.Cells[row, col + 1].Value?.ToString();

                            // ���X�çP�_��쪺�I����
                            var SocketColor1 = Polling2.Cells[row, col].Style.Fill.BackgroundColor;
                            var SocketColor2 = Polling2.Cells[row, col + 1].Style.Fill.BackgroundColor;

                            if ((SocketColor1.Rgb != Color.FromArgb(204, 153, 255).ToArgb().ToString("X8") || SocketColor2.Rgb != Color.FromArgb(204, 153, 255).ToArgb().ToString("X8")) && (SocketColor1.Rgb != Color.DimGray.ToArgb().ToString("X8") || SocketColor2.Rgb != Color.DimGray.ToArgb().ToString("X8")))
                            {
                                Polling2.Cells[row, col].Clear();

                                // �]�m��r�����M�����~��
                                Polling2.Cells[row, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                Polling2.Cells[row, col].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                                // �]�m�ؽu�˦�
                                Polling2.Cells[row, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;      // �W�ؽu
                                Polling2.Cells[row, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;   // �U�ؽu
                                Polling2.Cells[row, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;     // ���ؽu
                                Polling2.Cells[row, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;    // �k�ؽu

                                // �]�m�ؽu�C�⬰�¦�
                                Polling2.Cells[row, col].Style.Border.Top.Color.SetColor(Color.Black);
                                Polling2.Cells[row, col].Style.Border.Bottom.Color.SetColor(Color.Black);
                                Polling2.Cells[row, col].Style.Border.Left.Color.SetColor(Color.Black);
                                Polling2.Cells[row, col].Style.Border.Right.Color.SetColor(Color.Black);

                                // �]�m��r�����M�����~��
                                Polling2.Cells[row, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                Polling2.Cells[row, col].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                                // �]�m�I���⬰�z����
                                Polling2.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.None;
                            }


                            // �P�w�æX�����ƭ�
                            if (int.TryParse(SocketValue1, out int SocketNum1) && int.TryParse(SocketValue2, out int SocketNum2))
                            {
                                // �Y�����쬰�ƭȡA�h�i��[�`
                                Polling2.Cells[row, col].Value = SocketNum1 + SocketNum2;
                            }
                            else if (int.TryParse(SocketValue1, out SocketNum1) && string.IsNullOrEmpty(SocketValue2))
                            {
                                // �Y�䤤�@����쬰�ƭȡA�h�^��ƭ�
                                Polling2.Cells[row, col].Value = SocketNum1;
                            }
                            else if (string.IsNullOrEmpty(SocketValue1) && int.TryParse(SocketValue2, out SocketNum2))
                            {
                                // �Y�䤤�@����쬰�ƭȡA�h�^��ƭ�
                                Polling2.Cells[row, col].Value = SocketNum2;
                            }
                            else if (SocketValue1 != null && SocketValue1.StartsWith('#') && string.IsNullOrEmpty(SocketValue2))
                             {
                                // �Y�Ĥ@��쬰�� # �Ÿ��B�ĤG��쬰�šA�h�^��Ĥ@����r
                                Polling2.Cells[row, col].Value = SocketValue1;
                            }
                            else if (string.IsNullOrEmpty(SocketValue1) && SocketValue2 != null && SocketValue2.StartsWith("#"))
                            {
                                // �Y�ĤG��쬰�� # �Ÿ��B�Ĥ@��쬰�šA�h�^��ĤG����r
                                Polling2.Cells[row, col].Value = SocketValue2;
                            }
                            else if (int.TryParse(SocketValue1, out SocketNum1) && SocketValue2 != null && SocketValue2.StartsWith("#"))
                            {
                                // �Y�Ĥ@��쬰�ƭȡB�ĤG��쬰�� # �Ÿ��A�h�����ĤG���Ӧ^��Ĥ@���ƭ�
                                Polling2.Cells[row, col].Value = SocketNum1;
                            }
                            else if (SocketValue1 != null && SocketValue1.StartsWith("#") && int.TryParse(SocketValue2, out SocketNum2))
                            {
                                // �Y�Ĥ@��쬰�ƭȡB�ĤG��쬰�� # �Ÿ��A�h�����ĤG���Ӧ^��Ĥ@���ƭ�
                                Polling2.Cells[row, col].Value = SocketNum2;
                            }

                            if (SocketIndex < SocketColList.Count - 1)
                            {
                                SocketIndex++;
                            }
                        }

                        // ���D���]�m�I����
                        if (row <= ComEndRow)
                        {
                            if (row >= SURow && StartList.Count > 0 && EndList.Count > 0 && FtStartCol != 0 && col >= FtStartCol && (FtEndCol == 0 || (FtEndCol != 0 && col <= FtEndCol)))
                            {
                                // �ˬd�ô��� "1S1U" �� "1S"�B"2S1U" �� "2S"
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
                                    // �x�s�ݦX�֪���
                                    SocketColList.Add(col);
                                }
                                else if (!RemoveColList.Contains(col) && (CellValue.StartsWith("1S2U") || CellValue.StartsWith("2S2U")))
                                {
                                    // �x�s�ݲ�������
                                    RemoveColList.Add(col);
                                }

                                // �����X�����ƭ�
                                var OriginalValue = Polling2.Cells[row, col].Value;

                                Polling2.Cells[row, col].Clear();
                                // �]�m�ؽu�˦�
                                Polling2.Cells[row, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;      // �W�ؽu
                                Polling2.Cells[row, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;   // �U�ؽu
                                Polling2.Cells[row, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;     // ���ؽu
                                Polling2.Cells[row, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;    // �k�ؽu

                                // �]�m�ؽu�C�⬰�¦�
                                Polling2.Cells[row, col].Style.Border.Top.Color.SetColor(Color.Black);
                                Polling2.Cells[row, col].Style.Border.Bottom.Color.SetColor(Color.Black);
                                Polling2.Cells[row, col].Style.Border.Left.Color.SetColor(Color.Black);
                                Polling2.Cells[row, col].Style.Border.Right.Color.SetColor(Color.Black);

                                // �]�m��r�C�⬰�զ�
                                Polling2.Cells[row, col].Style.Font.Color.SetColor(Color.White);

                                // �]�m��r������
                                Polling2.Cells[row, col].Style.Font.Bold = true;

                                // �]�m��r�����M�����~��
                                Polling2.Cells[row, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                Polling2.Cells[row, col].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                                Polling2.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.None;
                                Polling2.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                Polling2.Cells[row, col].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(204, 153, 255));

                                // �^������
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
                // �p�GEndRow���ȡA�h���X�j��
                if (EndRow != 0)
                {
                    break;
                }
            }
            // �����x�s����
            int RemoveIndex = 0;
            foreach (int RemoveCol in RemoveColList)
            {
                Polling2.DeleteColumn(RemoveCol - RemoveIndex);
                RemoveIndex++;
            }
        }

        private object GetCellValue(ExcelRange cell)
        {
            // �p�G���s�b�����A�h�H�����p�⵲�G�A�_�h�H���Ȭ����G
            if (!string.IsNullOrEmpty(cell.Formula))
            {
                if (double.TryParse(cell.GetValue<string>(), out double result))
                {
                    return result;
                }
                else
                {
                    // �L�k�ഫ�ƭȮɫh�ϥΤ�r
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
            // �]�m�ؽu�˦�
            Polling1.Cells[row, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            Polling1.Cells[row, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            Polling1.Cells[row, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            Polling1.Cells[row, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            // �]�m�ؽu�C�⬰�¦�
            Polling1.Cells[row, col].Style.Border.Top.Color.SetColor(Color.Black);
            Polling1.Cells[row, col].Style.Border.Bottom.Color.SetColor(Color.Black);
            Polling1.Cells[row, col].Style.Border.Left.Color.SetColor(Color.Black);
            Polling1.Cells[row, col].Style.Border.Right.Color.SetColor(Color.Black);
        }

        private string ChangeSymbol(string formula)
        {
            return formula.Replace("+", ",");
        }

        // �N�B�z�᪺ Excel �ɮץt�s���s��
        private void SaveAsNewExcel(ExcelPackage package)
        {
            // �]�w EPPlus ���v�Ҧ��A�D�ӷ~�γ~
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel Files|*.xlsx";
                saveFileDialog.Title = "�t�s���s�� Excel ���";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    FileInfo fileInfo = new FileInfo(saveFileDialog.FileName);
                    package.SaveAs(fileInfo);
                    MessageBox.Show("�ɮפw�O�s��: " + saveFileDialog.FileName);
                }
            }
        }

        // �M���Ȧs�ܼ�
        private void ClearCache()
        {
            // �M�� TemplatePath ���Ȧs
            if (!string.IsNullOrEmpty(TemplatePath))
            {
                TemplatePath = null;
                NonFileFirst.Text = "(�|���W���ɮסI)";
                NonFileLast.Text = "(�|���W���ɮסI)";
            }

            // �M�� MultipleFilesPaths ���Ȧs
            if (MultipleFilesPaths != null && MultipleFilesPaths.Count > 0)
            {
                MultipleFilesPaths?.Clear();
                NonFIleMuti.Text = "(�|���W���ɮסI)";
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

            // ��ʦ^���O����
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
    }
}
