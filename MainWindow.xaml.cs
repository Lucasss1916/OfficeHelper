using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;
using System.Windows.Controls;
using System.Windows.Threading;
using Microsoft.Win32;
using ClosedXML.Excel;

namespace RandomScoreAllocatorWPF
{
    public partial class MainWindow : Window
    {
        private string _loadedFilePath = "";
        private string _selectedSheetName = "";
        private DataTable? _previewTable = null;
        private List<string> _currentSubjectList = new();
        private List<int> _currentEffectiveMaxScoreList = new();
        private int _currentMinEach = 0;
        private bool _currentUseFixedSeed = true;
        private bool _isUpdatingPreviewGrid = false;

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
                try
                {
                    _loadedFilePath = dialog.FileName;
                    TxtFile.Text = _loadedFilePath;
                    LoadWorkbookSheets();
                    MessageBox.Show("文件已选中。", "准备就绪");
                }
                catch (Exception ex)
                {
                    _loadedFilePath = "";
                    _selectedSheetName = "";
                    TxtFile.Text = "请选择文件...";
                    CmbSheets.ItemsSource = null;
                    CmbSheets.SelectedItem = null;
                    MessageBox.Show($"读取工作表失败: {ex.Message}", "错误");
                }
            }
        }

        private void LoadWorkbookSheets()
        {
            CmbSheets.ItemsSource = null;
            CmbSheets.SelectedItem = null;
            _selectedSheetName = "";
            _previewTable = null;
            _currentSubjectList.Clear();
            _currentEffectiveMaxScoreList.Clear();
            GridPreview.ItemsSource = null;

            using var wb = new XLWorkbook(_loadedFilePath);
            var sheetNames = wb.Worksheets.Select(ws => ws.Name).ToList();

            CmbSheets.ItemsSource = sheetNames;
            if (sheetNames.Count > 0)
            {
                CmbSheets.SelectedIndex = 0;
            }
        }

        private void CmbSheets_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _selectedSheetName = CmbSheets.SelectedItem?.ToString() ?? "";
            _previewTable = null;
            _currentSubjectList.Clear();
            _currentEffectiveMaxScoreList.Clear();
            GridPreview.ItemsSource = null;
        }

        private void GridPreview_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            if (e.PropertyName == "计算和")
            {
                e.Column.IsReadOnly = true;
            }
        }

        private void GridPreview_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (_isUpdatingPreviewGrid || e.EditAction != DataGridEditAction.Commit)
                return;

            if (e.Row.Item is not DataRowView rowView)
                return;

            string columnName = e.Column.Header?.ToString() ?? string.Empty;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                HandlePreviewCellEdited(rowView.Row, columnName);
            }), DispatcherPriority.Background);
        }

        private void HandlePreviewCellEdited(DataRow row, string columnName)
        {
            if (_previewTable == null || _currentSubjectList.Count == 0 || _isUpdatingPreviewGrid)
                return;

            if (columnName == "姓名" || string.IsNullOrWhiteSpace(columnName))
                return;

            try
            {
                _isUpdatingPreviewGrid = true;

                if (columnName == "原总分")
                {
                    RecalculateRowFromTotal(row);
                }
                else if (_currentSubjectList.Contains(columnName))
                {
                    RecalculateRowFromEditedSubject(row, columnName);
                }
                else if (columnName == "计算和")
                {
                    row["计算和"] = CalculateRowSubjectSum(row);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"更新当前行失败: {ex.Message}", "错误");
            }
            finally
            {
                _isUpdatingPreviewGrid = false;
            }
        }

        private void BtnPreview_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_loadedFilePath))
            {
                MessageBox.Show("请先选择 Excel 文件。", "提示");
                return;
            }

            if (string.IsNullOrWhiteSpace(_selectedSheetName))
            {
                MessageBox.Show("请选择需要查看的工作表。", "提示");
                return;
            }

            try
            {
                using var wb = new XLWorkbook(_loadedFilePath);
                var ws = wb.Worksheet(_selectedSheetName);
                var headerRow = ws.Row(1);
                var lastHeaderCell = headerRow.LastCellUsed();

                if (lastHeaderCell == null)
                    throw new Exception("当前工作表没有可识别的表头。");

                int lastCol = lastHeaderCell.Address.ColumnNumber;

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
                double? maxEachLimitPercent = null;

                if (!string.IsNullOrWhiteSpace(TxtMaxEachLimit.Text))
                {
                    if (!double.TryParse(TxtMaxEachLimit.Text, out double maxLimitPercentValue) || maxLimitPercentValue < 0 || maxLimitPercentValue > 100)
                        throw new Exception("最高分百分比上限必须是 0 到 100 之间的数字。");

                    maxEachLimitPercent = maxLimitPercentValue;
                }

                bool useFixedSeed = ChkFixedSeed.IsChecked == true;
                var subjectList = subjectMaxScores.Keys.ToList();
                var maxScoreList = subjectMaxScores.Values.ToList();
                var effectiveMaxScoreList = maxScoreList
                    .Select(maxScore =>
                    {
                        if (!maxEachLimitPercent.HasValue)
                            return maxScore;

                        int percentLimitedMax = (int)Math.Floor(maxScore * (maxEachLimitPercent.Value / 100.0));
                        return Math.Min(maxScore, percentLimitedMax);
                    })
                    .ToList();

                if (effectiveMaxScoreList.Any(maxScore => maxScore < minEach))
                    throw new Exception("存在科目的最高分上限低于最低分，请调整参数。");

                _currentSubjectList = new List<string>(subjectList);
                _currentEffectiveMaxScoreList = new List<int>(effectiveMaxScoreList);
                _currentMinEach = minEach;
                _currentUseFixedSeed = useFixedSeed;

                int recognizedPaperMax = effectiveMaxScoreList.Sum();

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
                            effectiveMaxScoreList,
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
                MessageBox.Show($"生成完成！\n当前工作表: {_selectedSheetName}\n按限制后可分配总满分: {recognizedPaperMax} 分。", "完成");
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
                        var exportSheetName = string.IsNullOrWhiteSpace(_selectedSheetName) ? "分配详情" : $"{_selectedSheetName}_分配详情";
                        var ws = wb.Worksheets.Add(exportSheetName.Length > 31 ? exportSheetName.Substring(0, 31) : exportSheetName);
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

        private void RecalculateRowFromTotal(DataRow row)
        {
            int targetTotal = GetRowIntValue(row, "原总分");
            int rowIndex = GetRowIndex(row);

            Dictionary<string, int> scores;
            bool shouldUpdateTargetTotal = false;
            if (targetTotal <= 0)
            {
                scores = _currentSubjectList.ToDictionary(subject => subject, _ => 0);
                targetTotal = 0;
                shouldUpdateTargetTotal = true;
            }
            else
            {
                int? seed = _currentUseFixedSeed ? (rowIndex * 999 + targetTotal) : (int?)null;
                scores = ScoreAllocator.Allocate(
                    targetTotal,
                    _currentSubjectList,
                    _currentEffectiveMaxScoreList,
                    minEach: _currentMinEach,
                    seed: seed
                );
            }

            ApplyRowScores(row, scores, targetTotal, updateTargetTotal: shouldUpdateTargetTotal);
        }

        private void RecalculateRowFromEditedSubject(DataRow row, string editedSubject)
        {
            int editedIndex = _currentSubjectList.IndexOf(editedSubject);
            if (editedIndex < 0)
                return;

            int editedValue = GetRowIntValue(row, editedSubject);
            editedValue = Math.Max(_currentMinEach, Math.Min(_currentEffectiveMaxScoreList[editedIndex], editedValue));

            int originalTargetTotal = GetRowIntValue(row, "原总分");

            var remainingSubjects = new List<string>();
            var remainingMaxScores = new List<int>();

            for (int i = 0; i < _currentSubjectList.Count; i++)
            {
                if (i == editedIndex)
                    continue;

                remainingSubjects.Add(_currentSubjectList[i]);
                remainingMaxScores.Add(_currentEffectiveMaxScoreList[i]);
            }

            int remainingMinTotal = remainingSubjects.Count * _currentMinEach;
            int remainingMaxTotal = remainingMaxScores.Sum();
            int adjustedTargetTotal = Math.Max(originalTargetTotal, editedValue);
            adjustedTargetTotal = Math.Max(adjustedTargetTotal, editedValue + remainingMinTotal);
            adjustedTargetTotal = Math.Min(adjustedTargetTotal, editedValue + remainingMaxTotal);

            Dictionary<string, int> recalculatedScores;
            if (remainingSubjects.Count == 0)
            {
                recalculatedScores = new Dictionary<string, int>();
            }
            else
            {
                int remainingTarget = adjustedTargetTotal - editedValue;
                int rowIndex = GetRowIndex(row);
                int? seed = _currentUseFixedSeed ? (rowIndex * 999 + adjustedTargetTotal) : (int?)null;

                recalculatedScores = ScoreAllocator.Allocate(
                    remainingTarget,
                    remainingSubjects,
                    remainingMaxScores,
                    minEach: _currentMinEach,
                    seed: seed
                );
            }

            recalculatedScores[editedSubject] = editedValue;
            row["原总分"] = adjustedTargetTotal;
            ApplyRowScores(row, recalculatedScores, adjustedTargetTotal, updateTargetTotal: true);
        }

        private void ApplyRowScores(DataRow row, Dictionary<string, int> scores, int targetTotal, bool updateTargetTotal)
        {
            foreach (var subject in _currentSubjectList)
            {
                row[subject] = scores.TryGetValue(subject, out int score) ? score : 0;
            }

            if (updateTargetTotal)
            {
                row["原总分"] = targetTotal;
            }

            row["计算和"] = CalculateRowSubjectSum(row);
        }

        private int CalculateRowSubjectSum(DataRow row)
        {
            int sum = 0;
            foreach (var subject in _currentSubjectList)
            {
                sum += GetRowIntValue(row, subject);
            }

            return sum;
        }

        private int GetRowIntValue(DataRow row, string columnName)
        {
            if (!_previewTable!.Columns.Contains(columnName))
                return 0;

            string rawValue = row[columnName]?.ToString()?.Trim() ?? string.Empty;
            if (!int.TryParse(rawValue, out int value))
                return 0;

            return value;
        }

        private int GetRowIndex(DataRow row)
        {
            return _previewTable == null ? 1 : _previewTable.Rows.IndexOf(row) + 1;
        }
    }
}
