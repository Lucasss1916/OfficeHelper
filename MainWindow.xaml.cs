using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using ClosedXML.Excel;
using Microsoft.Win32;

namespace RandomScoreAllocatorWPF
{
    public partial class MainWindow : Window
    {
        private const string SourceRowColumnName = "__SourceRow";
        private static readonly Regex SubjectHeaderRegex = new(@"(?<name>.+?)[（\(]\s*(?<score>\d+)\s*(?:分|points)\s*[）\)]", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private string _loadedFilePath = "";
        private string _selectedSheetName = "";
        private DataTable? _previewTable;
        private SheetLayout? _currentSheetLayout;
        private List<string> _currentSubjectList = new();
        private List<int> _currentEffectiveMaxScoreList = new();
        private int _currentMinEach;
        private bool _currentUseFixedSeed = true;
        private bool _isUpdatingPreviewGrid;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Header_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ButtonState == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }

        private void BtnMinimize_Click(object sender, RoutedEventArgs e) => WindowState = WindowState.Minimized;
        private void BtnClose_Click(object sender, RoutedEventArgs e) => Close();

        private void BtnHelp_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(
                "1. 选择 Excel 文件并切换工作表。\n2. 点击生成预览，程序会自动识别复杂表头。\n3. 编辑预览中的总分或单科分数后，程序会按规则联动更新该行。\n4. 导出时会保留原工作表样式，只回填成绩区域。",
                "使用说明");
        }

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
                    ResetLoadedWorkbook();
                    MessageBox.Show($"读取工作表失败: {ex.Message}", "错误");
                }
            }
        }

        private void LoadWorkbookSheets()
        {
            ResetPreviewState();

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
            ResetPreviewState();
        }

        private void GridPreview_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            if (e.PropertyName == SourceRowColumnName)
            {
                e.Cancel = true;
                return;
            }

            if (e.PropertyName == "计算和" || e.PropertyName == "学号" || e.PropertyName == "姓名")
            {
                e.Column.IsReadOnly = true;
            }
        }

        private void GridPreview_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (_isUpdatingPreviewGrid || e.EditAction != DataGridEditAction.Commit)
            {
                return;
            }

            if (e.Row.Item is not DataRowView rowView)
            {
                return;
            }

            string columnName = e.Column.Header?.ToString() ?? string.Empty;
            Dispatcher.BeginInvoke(new Action(() => HandlePreviewCellEdited(rowView.Row, columnName)), DispatcherPriority.Background);
        }

        private void BtnPreview_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_loadedFilePath))
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
                var layout = DetectSheetLayout(ws);
                var generationOptions = ReadGenerationOptions();
                var effectiveMaxScoreList = layout.SubjectColumns
                    .Select(subject => ApplyMaxPercentLimit(subject.MaxScore, generationOptions.MaxEachLimitPercent))
                    .ToList();

                if (effectiveMaxScoreList.Any(maxScore => maxScore < generationOptions.MinEach))
                {
                    throw new Exception("存在科目的最高分上限低于最低分，请调整参数。");
                }

                _currentSheetLayout = layout;
                _currentSubjectList = layout.SubjectColumns.Select(subject => subject.DisplayName).ToList();
                _currentEffectiveMaxScoreList = effectiveMaxScoreList;
                _currentMinEach = generationOptions.MinEach;
                _currentUseFixedSeed = generationOptions.UseFixedSeed;

                _previewTable = BuildPreviewTable(layout, ws, generationOptions, effectiveMaxScoreList);
                GridPreview.ItemsSource = _previewTable.DefaultView;

                int recognizedPaperMax = effectiveMaxScoreList.Sum();
                MessageBox.Show(
                    $"生成完成！\n当前工作表: {_selectedSheetName}\n识别表头行: {layout.HeaderStartRow}-{layout.HeaderEndRow}\n按限制后可分配总满分: {recognizedPaperMax} 分。",
                    "完成");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"生成预览失败: {ex.Message}", "错误");
            }
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            if (_previewTable == null || _previewTable.Rows.Count == 0 || _currentSheetLayout == null)
            {
                MessageBox.Show("无数据可导出。", "提示");
                return;
            }

            var saveDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"分配结果_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
            };

            if (saveDialog.ShowDialog() != true)
            {
                return;
            }

            try
            {
                using var wb = new XLWorkbook(_loadedFilePath);
                var ws = wb.Worksheet(_selectedSheetName);

                foreach (DataRow previewRow in _previewTable.Rows)
                {
                    int sourceRow = GetSourceRowNumber(previewRow);
                    foreach (var subject in _currentSheetLayout.SubjectColumns)
                    {
                        ws.Cell(sourceRow, subject.ColumnIndex).Value = GetRowIntValue(previewRow, subject.DisplayName);
                    }

                    ws.Cell(sourceRow, _currentSheetLayout.TotalColumnIndex).Value = GetRowIntValue(previewRow, "原总分");
                }

                wb.SaveAs(saveDialog.FileName);
                MessageBox.Show("导出成功，已保留原工作表样式。", "成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出失败: {ex.Message}", "错误");
            }
        }

        private void HandlePreviewCellEdited(DataRow row, string columnName)
        {
            if (_previewTable == null || _currentSheetLayout == null || _currentSubjectList.Count == 0 || _isUpdatingPreviewGrid)
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(columnName) || columnName == "姓名" || columnName == "学号")
            {
                return;
            }

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

        private DataTable BuildPreviewTable(SheetLayout layout, IXLWorksheet ws, GenerationOptions options, List<int> effectiveMaxScoreList)
        {
            var previewTable = CreatePreviewTable(layout);
            int previewOrder = 0;

            for (int rowNumber = layout.DataStartRow; rowNumber <= layout.LastUsedRow; rowNumber++)
            {
                var row = ws.Row(rowNumber);
                if (!IsStudentDataRow(row, layout))
                {
                    continue;
                }

                previewOrder++;
                string name = GetCellDisplayText(ws, rowNumber, layout.NameColumnIndex);
                string studentId = layout.StudentIdColumnIndex.HasValue
                    ? GetCellDisplayText(ws, rowNumber, layout.StudentIdColumnIndex.Value)
                    : string.Empty;
                int targetTotal = ReadIntegerCell(row.Cell(layout.TotalColumnIndex));

                var scores = targetTotal <= 0
                    ? layout.SubjectColumns.ToDictionary(subject => subject.DisplayName, _ => 0)
                    : ScoreAllocator.Allocate(
                        targetTotal,
                        layout.SubjectColumns.Select(subject => subject.DisplayName).ToList(),
                        effectiveMaxScoreList,
                        minEach: options.MinEach,
                        seed: options.UseFixedSeed ? BuildSeed(rowNumber, targetTotal, previewOrder) : (int?)null);

                DataRow previewRow = previewTable.NewRow();
                previewRow[SourceRowColumnName] = rowNumber;
                if (previewTable.Columns.Contains("学号"))
                {
                    previewRow["学号"] = studentId;
                }

                previewRow["姓名"] = name;
                foreach (var score in scores)
                {
                    previewRow[score.Key] = score.Value;
                }

                previewRow["原总分"] = targetTotal;
                previewRow["计算和"] = scores.Values.Sum();
                previewTable.Rows.Add(previewRow);
            }

            return previewTable;
        }

        private DataTable CreatePreviewTable(SheetLayout layout)
        {
            var table = new DataTable();
            table.Columns.Add(SourceRowColumnName, typeof(int));

            if (layout.StudentIdColumnIndex.HasValue)
            {
                table.Columns.Add("学号");
            }

            table.Columns.Add("姓名");
            foreach (var subject in layout.SubjectColumns)
            {
                table.Columns.Add(subject.DisplayName, typeof(int));
            }

            table.Columns.Add("原总分", typeof(int));
            table.Columns.Add("计算和", typeof(int));
            return table;
        }

        private SheetLayout DetectSheetLayout(IXLWorksheet ws)
        {
            var usedRange = ws.RangeUsed();
            if (usedRange == null)
            {
                throw new Exception("当前工作表为空。");
            }

            int lastUsedRow = usedRange.LastRow().RowNumber();
            int lastUsedColumn = usedRange.LastColumn().ColumnNumber();
            int maxScanRow = Math.Min(lastUsedRow, 12);
            SheetLayout? bestLayout = null;
            int bestScore = -1;

            // 复杂模板通常会把标题、课程信息和真正表头分成多行。
            // 这里扫描前几行的连续区间，挑出最像“成绩表头”的那一段。
            for (int headerStart = 1; headerStart <= maxScanRow; headerStart++)
            {
                for (int headerEnd = headerStart; headerEnd <= Math.Min(maxScanRow, headerStart + 5); headerEnd++)
                {
                    var candidate = TryBuildSheetLayout(ws, headerStart, headerEnd, lastUsedRow, lastUsedColumn);
                    if (candidate == null)
                    {
                        continue;
                    }

                    int score = candidate.SubjectColumns.Count * 10 + (candidate.StudentIdColumnIndex.HasValue ? 2 : 0) - headerStart;
                    if (score > bestScore)
                    {
                        bestScore = score;
                        bestLayout = candidate;
                    }
                }
            }

            if (bestLayout == null)
            {
                throw new Exception("未能识别复杂表头。请确保表格中包含“姓名”、学科“(xx分)”列和“总分”列。");
            }

            return bestLayout;
        }

        private SheetLayout? TryBuildSheetLayout(IXLWorksheet ws, int headerStart, int headerEnd, int lastUsedRow, int lastUsedColumn)
        {
            int? nameColumnIndex = null;
            int? studentIdColumnIndex = null;
            int? totalColumnIndex = null;
            var subjects = new List<SubjectColumn>();

            for (int col = 1; col <= lastUsedColumn; col++)
            {
                // 将多行表头在同一列上的文本合并后识别。
                // 这样像“期末考试(40%)”+“填空题(20分)”这种结构也能解析到最终题目列。
                string combinedHeader = BuildCombinedHeader(ws, headerStart, headerEnd, col);
                if (string.IsNullOrWhiteSpace(combinedHeader))
                {
                    continue;
                }

                string normalizedHeader = NormalizeHeader(combinedHeader);

                if (nameColumnIndex == null && IsNameHeader(normalizedHeader))
                {
                    nameColumnIndex = col;
                    continue;
                }

                if (studentIdColumnIndex == null && IsStudentIdHeader(normalizedHeader))
                {
                    studentIdColumnIndex = col;
                    continue;
                }

                if (totalColumnIndex == null && IsTotalHeader(normalizedHeader))
                {
                    totalColumnIndex = col;
                    continue;
                }

                var subjectMatch = SubjectHeaderRegex.Match(normalizedHeader);
                if (!subjectMatch.Success)
                {
                    continue;
                }

                string displayName = $"{subjectMatch.Groups["name"].Value.Trim()}({subjectMatch.Groups["score"].Value}分)";
                if (subjects.Any(subject => subject.DisplayName == displayName))
                {
                    displayName = $"{displayName}_{col}";
                }

                subjects.Add(new SubjectColumn(col, displayName, int.Parse(subjectMatch.Groups["score"].Value)));
            }

            if (nameColumnIndex == null || totalColumnIndex == null || subjects.Count == 0)
            {
                return null;
            }

            return new SheetLayout(
                headerStart,
                headerEnd,
                headerEnd + 1,
                lastUsedRow,
                nameColumnIndex.Value,
                studentIdColumnIndex,
                totalColumnIndex.Value,
                subjects);
        }

        private GenerationOptions ReadGenerationOptions()
        {
            int minEach = 0;
            if (int.TryParse(TxtMinEach.Text, out int minVal))
            {
                minEach = minVal;
            }

            double? maxEachLimitPercent = null;
            if (!string.IsNullOrWhiteSpace(TxtMaxEachLimit.Text))
            {
                if (!double.TryParse(TxtMaxEachLimit.Text, out double maxPercent) || maxPercent < 0 || maxPercent > 100)
                {
                    throw new Exception("最高分百分比上限必须是 0 到 100 之间的数字。");
                }

                maxEachLimitPercent = maxPercent;
            }

            return new GenerationOptions(minEach, maxEachLimitPercent, ChkFixedSeed.IsChecked == true);
        }

        private int ApplyMaxPercentLimit(int originalMaxScore, double? maxEachLimitPercent)
        {
            if (!maxEachLimitPercent.HasValue)
            {
                return originalMaxScore;
            }

            int percentLimitedMax = (int)Math.Floor(originalMaxScore * (maxEachLimitPercent.Value / 100.0));
            return Math.Min(originalMaxScore, percentLimitedMax);
        }

        private bool IsStudentDataRow(IXLRow row, SheetLayout layout)
        {
            string name = GetCellDisplayText(row.Worksheet, row.RowNumber(), layout.NameColumnIndex);
            string studentId = layout.StudentIdColumnIndex.HasValue
                ? GetCellDisplayText(row.Worksheet, row.RowNumber(), layout.StudentIdColumnIndex.Value)
                : string.Empty;

            return !string.IsNullOrWhiteSpace(name) || !string.IsNullOrWhiteSpace(studentId);
        }

        private void RecalculateRowFromTotal(DataRow row)
        {
            int targetTotal = GetRowIntValue(row, "原总分");
            int seedKey = GetSeedKey(row);
            bool updateTargetTotal = false;
            Dictionary<string, int> scores;

            if (targetTotal <= 0)
            {
                targetTotal = 0;
                updateTargetTotal = true;
                scores = _currentSubjectList.ToDictionary(subject => subject, _ => 0);
            }
            else
            {
                scores = ScoreAllocator.Allocate(
                    targetTotal,
                    _currentSubjectList,
                    _currentEffectiveMaxScoreList,
                    minEach: _currentMinEach,
                    seed: _currentUseFixedSeed ? seedKey * 997 + targetTotal : (int?)null);
            }

            ApplyRowScores(row, scores, targetTotal, updateTargetTotal);
        }

        private void RecalculateRowFromEditedSubject(DataRow row, string editedSubject)
        {
            int editedIndex = _currentSubjectList.IndexOf(editedSubject);
            if (editedIndex < 0)
            {
                return;
            }

            int editedValue = GetRowIntValue(row, editedSubject);
            editedValue = Math.Max(_currentMinEach, Math.Min(_currentEffectiveMaxScoreList[editedIndex], editedValue));

            int originalTargetTotal = GetRowIntValue(row, "原总分");
            var remainingSubjects = new List<string>();
            var remainingMaxScores = new List<int>();

            for (int i = 0; i < _currentSubjectList.Count; i++)
            {
                if (i == editedIndex)
                {
                    continue;
                }

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
                int seedKey = GetSeedKey(row);
                recalculatedScores = ScoreAllocator.Allocate(
                    remainingTarget,
                    remainingSubjects,
                    remainingMaxScores,
                    minEach: _currentMinEach,
                    seed: _currentUseFixedSeed ? seedKey * 997 + adjustedTargetTotal : (int?)null);
            }

            recalculatedScores[editedSubject] = editedValue;
            ApplyRowScores(row, recalculatedScores, adjustedTargetTotal, updateTargetTotal: true);
        }

        private void ApplyRowScores(DataRow row, Dictionary<string, int> scores, int targetTotal, bool updateTargetTotal)
        {
            foreach (var subject in _currentSubjectList)
            {
                row[subject] = scores.TryGetValue(subject, out int value) ? value : 0;
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
            if (_previewTable == null || !_previewTable.Columns.Contains(columnName))
            {
                return 0;
            }

            string rawValue = row[columnName]?.ToString()?.Trim() ?? string.Empty;
            return int.TryParse(rawValue, out int value) ? value : 0;
        }

        private int GetSourceRowNumber(DataRow row)
        {
            return int.TryParse(row[SourceRowColumnName]?.ToString(), out int value) ? value : 0;
        }

        private int GetSeedKey(DataRow row)
        {
            int sourceRow = GetSourceRowNumber(row);
            return sourceRow > 0 ? sourceRow : (_previewTable?.Rows.IndexOf(row) ?? 0) + 1;
        }

        private int BuildSeed(int sourceRow, int targetTotal, int previewOrder)
        {
            int seedKey = sourceRow > 0 ? sourceRow : previewOrder;
            return seedKey * 997 + targetTotal;
        }

        private string BuildCombinedHeader(IXLWorksheet ws, int headerStart, int headerEnd, int columnIndex)
        {
            var parts = new List<string>();
            for (int row = headerStart; row <= headerEnd; row++)
            {
                string value = GetCellDisplayText(ws, row, columnIndex);
                if (string.IsNullOrWhiteSpace(value))
                {
                    continue;
                }

                if (!parts.Contains(value))
                {
                    parts.Add(value);
                }
            }

            return string.Join(" ", parts);
        }

        private string GetCellDisplayText(IXLWorksheet ws, int rowNumber, int columnIndex)
        {
            var cell = ws.Cell(rowNumber, columnIndex);
            if (cell.IsMerged())
            {
                var mergedRange = cell.MergedRange();

                // 对横向合并的大标题只在起始列读取一次，其他列直接忽略，
                // 避免“成绩形成明细表”“期末考试(40%)”这类横幅标题污染具体题目列名。
                if (mergedRange.RangeAddress.FirstAddress.ColumnNumber != columnIndex)
                {
                    return string.Empty;
                }

                if (mergedRange.ColumnCount() > 1)
                {
                    return string.Empty;
                }

                cell = mergedRange.FirstCell();
            }

            return NormalizeHeader(cell.GetString());
        }

        private int ReadIntegerCell(IXLCell cell)
        {
            if (cell.TryGetValue(out int intValue))
            {
                return intValue;
            }

            string raw = cell.GetFormattedString().Trim();
            return int.TryParse(raw, out int parsed) ? parsed : 0;
        }

        private string NormalizeHeader(string value)
        {
            return value
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Replace("  ", " ")
                .Trim();
        }

        private bool IsNameHeader(string header)
        {
            return header.Contains("姓名", StringComparison.OrdinalIgnoreCase)
                || header.Equals("Name", StringComparison.OrdinalIgnoreCase)
                || header.Contains("学生姓名", StringComparison.OrdinalIgnoreCase);
        }

        private bool IsStudentIdHeader(string header)
        {
            return header.Contains("学号", StringComparison.OrdinalIgnoreCase)
                || header.Contains("studentid", StringComparison.OrdinalIgnoreCase)
                || header.Contains("student id", StringComparison.OrdinalIgnoreCase);
        }

        private bool IsTotalHeader(string header)
        {
            return header.Contains("总分", StringComparison.OrdinalIgnoreCase)
                || header.Contains("合计", StringComparison.OrdinalIgnoreCase)
                || header.Equals("Total", StringComparison.OrdinalIgnoreCase)
                || header.Contains("总成绩", StringComparison.OrdinalIgnoreCase);
        }

        private void ResetLoadedWorkbook()
        {
            _loadedFilePath = string.Empty;
            _selectedSheetName = string.Empty;
            TxtFile.Text = "请选择文件...";
            CmbSheets.ItemsSource = null;
            CmbSheets.SelectedItem = null;
            ResetPreviewState();
        }

        private void ResetPreviewState()
        {
            _previewTable = null;
            _currentSheetLayout = null;
            _currentSubjectList.Clear();
            _currentEffectiveMaxScoreList.Clear();
            GridPreview.ItemsSource = null;
        }

        private sealed record GenerationOptions(int MinEach, double? MaxEachLimitPercent, bool UseFixedSeed);

        private sealed record SubjectColumn(int ColumnIndex, string DisplayName, int MaxScore);

        private sealed record SheetLayout(
            int HeaderStartRow,
            int HeaderEndRow,
            int DataStartRow,
            int LastUsedRow,
            int NameColumnIndex,
            int? StudentIdColumnIndex,
            int TotalColumnIndex,
            List<SubjectColumn> SubjectColumns);
    }
}
