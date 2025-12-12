using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;
using ClosedXML.Excel;

namespace RandomScoreAllocatorWPF
{
    public partial class MainWindow : Window
    {
        private string _loadedFilePath = "";
        private DataTable _previewTable = null;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Header_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) { if (e.ButtonState == MouseButtonState.Pressed) this.DragMove(); }
        private void BtnMinimize_Click(object sender, RoutedEventArgs e) => this.WindowState = WindowState.Minimized;
        private void BtnClose_Click(object sender, RoutedEventArgs e) => this.Close();
        private void BtnHelp_Click(object sender, RoutedEventArgs e) => MessageBox.Show("1. 选择Excel文件。\n2. 点击生成预览。\n3. 程序将严格控制单科不超满分，并优先填补靠后的科目。", "使用说明");

        private void BtnBrowse_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls",
                Title = "选择 Excel 文件"
            };

            if (dialog.ShowDialog() == true)
            {
                _loadedFilePath = dialog.FileName;
                TxtFile.Text = _loadedFilePath;
                MessageBox.Show("文件已选中。", "准备就绪");
            }
        }

        private void BtnPreview_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_loadedFilePath))
            {
                MessageBox.Show("请先选择 Excel 文件。", "提示");
                return;
            }

            try
            {
                using var wb = new XLWorkbook(_loadedFilePath);
                var ws = wb.Worksheet(1);
                var headerRow = ws.Row(1);
                int lastCol = headerRow.LastCellUsed().Address.ColumnNumber;

                // --- 1. 识别列 ---
                int nameColIndex = -1;
                int totalColIndex = -1;
                var subjectMaxScores = new Dictionary<string, int>();

                var regexSubject = new Regex(@"([\s\S]+?)[（\(]\s*(\d+)\s*(?:分|points)?\s*[）\)]");

                for (int col = 1; col <= lastCol; col++)
                {
                    string rawTitle = headerRow.Cell(col).GetString().Trim();
                    string cleanTitle = rawTitle.Replace("\n", "").Replace("\r", "").Trim();

                    if (cleanTitle == "姓名" || cleanTitle == "Name" || cleanTitle == "学生姓名")
                    {
                        nameColIndex = col;
                        continue;
                    }
                    if (cleanTitle.Contains("合计") || cleanTitle.Contains("总分") || cleanTitle == "Total")
                    {
                        totalColIndex = col;
                        continue;
                    }

                    var m = regexSubject.Match(rawTitle);
                    if (m.Success)
                    {
                        int maxScore = int.Parse(m.Groups[2].Value);
                        if (!subjectMaxScores.ContainsKey(cleanTitle))
                        {
                            subjectMaxScores.Add(cleanTitle, maxScore);
                        }
                    }
                }

                if (nameColIndex == -1 || totalColIndex == -1 || subjectMaxScores.Count == 0)
                    throw new Exception($"未能识别列。找到的科目数：{subjectMaxScores.Count}。\n请检查表头是否包含“(20分)”字样。");

                // --- 2. 准备 DataTable (修改了列的添加顺序) ---
                _previewTable = new DataTable();
                _previewTable.Columns.Add("姓名");

                // 【修改点】先添加所有科目列
                foreach (var subjectName in subjectMaxScores.Keys)
                {
                    _previewTable.Columns.Add(subjectName, typeof(int));
                }

                // 【修改点】最后再添加总分列
                _previewTable.Columns.Add("原总分", typeof(int));
                _previewTable.Columns.Add("计算和", typeof(int));

                // --- 3. 逐行计算 ---
                var rows = ws.RowsUsed().Skip(1);
                int rowIndex = 0;
                int minEach = 0;
                if (int.TryParse(TxtMinEach.Text, out int minVal)) minEach = minVal;

                bool useFixedSeed = ChkFixedSeed.IsChecked == true;
                var subjectList = subjectMaxScores.Keys.ToList();
                var maxScoreList = subjectMaxScores.Values.ToList();

                int recognizedPaperMax = maxScoreList.Sum();

                foreach (var row in rows)
                {
                    rowIndex++;
                    string name = row.Cell(nameColIndex).GetString();
                    if (string.IsNullOrWhiteSpace(name)) continue;

                    if (!row.Cell(totalColIndex).TryGetValue(out int targetTotal)) targetTotal = 0;

                    Dictionary<string, int> scores;
                    if (targetTotal <= 0)
                    {
                        scores = subjectList.ToDictionary(s => s, s => 0);
                    }
                    else
                    {
                        int? seed = useFixedSeed ? (rowIndex * 999 + targetTotal) : (int?)null;

                        scores = ScoreAllocator.Allocate(
                            targetTotal,
                            subjectList,
                            maxScoreList,
                            minEach: minEach,
                            seed: seed
                        );
                    }

                    DataRow dtRow = _previewTable.NewRow();
                    dtRow["姓名"] = name;

                    // 填充各科分数
                    foreach (var kvp in scores)
                    {
                        dtRow[kvp.Key] = kvp.Value;
                    }

                    // 填充总分（顺序无关紧要，因为是通过列名赋值的，但DataTable结构决定了显示顺序）
                    dtRow["原总分"] = targetTotal;
                    dtRow["计算和"] = scores.Values.Sum();

                    _previewTable.Rows.Add(dtRow);
                }

                GridPreview.ItemsSource = _previewTable.DefaultView;
                MessageBox.Show($"生成完成！\n程序识别到的卷面总满分是: {recognizedPaperMax} 分。", "完成");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"生成预览失败: {ex.Message}", "错误");
            }
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            if (_previewTable == null || _previewTable.Rows.Count == 0)
            {
                MessageBox.Show("无数据可导出。", "提示");
                return;
            }

            var saveDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"分配结果_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
            };

            if (saveDialog.ShowDialog() == true)
            {
                try
                {
                    using (var wb = new XLWorkbook())
                    {
                        var ws = wb.Worksheets.Add("分配详情");
                        ws.Cell(1, 1).InsertTable(_previewTable);
                        ws.Columns().AdjustToContents();
                        wb.SaveAs(saveDialog.FileName);
                    }
                    MessageBox.Show("导出成功。", "成功");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"导出失败: {ex.Message}", "错误");
                }
            }
        }
    }
}