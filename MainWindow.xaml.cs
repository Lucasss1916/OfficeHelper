using ClosedXML.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;

namespace RandomScoreAllocatorWPF
{
    public partial class MainWindow : Window
    {
        private string? _excelPath;
        private List<string> _subjects = new();
        private List<double> _weights = new();
        private List<string> _students = new();
        private DataTable? _previewTable;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void BtnHelp_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("表格格式示例：\n\n姓名 | 数学(50%) | 语文(30%) | 英语(20%)\n张三 | ...\n李四 | ...\n\n“生成预览”不会修改原表，只在界面展示；“导出结果”会在同目录生成一个新文件。", "使用说明");
        }

        private void BtnBrowse_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { Filter = "Excel 文件 (*.xlsx)|*.xlsx" };
            if (dlg.ShowDialog() == true)
            {
                _excelPath = dlg.FileName;
                TxtFile.Text = _excelPath;
                try
                {
                    LoadExcelHeader(_excelPath);
                    LoadStudents(_excelPath);
                    MessageBox.Show($"读取成功：{_students.Count} 个学生，{_subjects.Count} 个科目。", "成功");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("读取失败：" + ex.Message);
                }
            }
        }

        private void LoadExcelHeader(string path)
        {
            _subjects.Clear();
            _weights.Clear();

            using var wb = new XLWorkbook(path);
            var ws = wb.Worksheet(1);
            var header = ws.Row(1);
            int lastCol = header.LastCellUsed().Address.ColumnNumber;

            var regex = new Regex(@"(.+)\((\d+)%\)");
            for (int col = 2; col <= lastCol; col++)
            {
                string title = header.Cell(col).GetString().Trim();
                var m = regex.Match(title);
                if (m.Success)
                {
                    _subjects.Add(m.Groups[1].Value.Trim());
                    _weights.Add(double.Parse(m.Groups[2].Value));
                }
                else
                {
                    _subjects.Add(title);
                    _weights.Add(1.0);
                }
            }
        }

        private void LoadStudents(string path)
        {
            _students.Clear();
            using var wb = new XLWorkbook(path);
            var ws = wb.Worksheet(1);
            int lastRow = ws.LastRowUsed().RowNumber();
            for (int row = 2; row <= lastRow; row++)
            {
                string name = ws.Cell(row, 1).GetString().Trim();
                if (!string.IsNullOrEmpty(name))
                    _students.Add(name);
            }
        }

        private void BtnPreview_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_excelPath))
            {
                MessageBox.Show("请先选择 Excel 文件。");
                return;
            }
            if (_subjects.Count == 0 || _students.Count == 0)
            {
                MessageBox.Show("未读取到科目或学生，请检查 Excel。");
                return;
            }

            if (!ValidateInputs(out int minEach, out double maxFrac))
                return;

            // 读取每个学生的总分
            Dictionary<string, int> studentTotals = new Dictionary<string, int>();
            try
            {
                using var wb = new XLWorkbook(_excelPath);
                var ws = wb.Worksheet(1);
                var header = ws.Row(1);
                int lastCol = header.LastCellUsed().Address.ColumnNumber;
                
                // 查找总分列的索引
                int totalColIndex = -1;
                for (int col = 2; col <= lastCol; col++)
                {
                    string title = header.Cell(col).GetString().Trim();
                    if (title == "总分" || title == "合计")
                    {
                        totalColIndex = col;
                        break;
                    }
                }
                
                if (totalColIndex == -1)
                {
                    MessageBox.Show("在Excel中未找到\"总分\"或\"合计\"列，请确保存在此列。");
                    return;
                }
                
                // 读取每个学生的总分
                int lastRow = ws.LastRowUsed().RowNumber();
                for (int row = 2; row <= lastRow; row++)
                {
                    string name = ws.Cell(row, 1).GetString().Trim();
                    if (!string.IsNullOrEmpty(name))
                    {
                        if (!int.TryParse(ws.Cell(row, totalColIndex).GetString(), out int total) || total <= 0)
                        {
                            MessageBox.Show($"学生 {name} 的总分无效，请确保总分为正整数。");
                            return;
                        }
                        studentTotals[name] = total;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("读取总分失败：" + ex.Message);
                return;
            }

            var table = new DataTable();
            table.Columns.Add("姓名", typeof(string));
            foreach (var s in _subjects) table.Columns.Add(s, typeof(int));
            table.Columns.Add("合计", typeof(int));

            for (int i = 0; i < _students.Count; i++)
            {
                var name = _students[i];
                
                // 获取当前学生的总分
                if (!studentTotals.TryGetValue(name, out int total))
                {
                    MessageBox.Show($"未找到学生 {name} 的总分数据。");
                    return;
                }
                
                // 验证该学生的总分是否符合要求
                if (_subjects.Count * minEach > total)
                {
                    MessageBox.Show($"学生 {name} 的总分过小：至少需要 {_subjects.Count * minEach} 才能满足每科最少 {minEach}。");
                    return;
                }
                
                var weights = ChkUseWeights.IsChecked == true ? _weights : Enumerable.Repeat(1.0, _subjects.Count).ToList();
                var alloc = ScoreAllocator.Allocate(
                    totalScore: total,
                    subjects: _subjects,
                    weights: weights,
                    minEach: minEach,
                    maxEachFraction: maxFrac,
                    randomness: 0.25,
                    seed: (ChkFixedSeed.IsChecked == true ? i + 1 : (int?)null)
                );
                
                var row = table.NewRow();
                row["姓名"] = name;
                int sum = 0;
                foreach (var s in _subjects)
                {
                    int v = alloc[s];
                    row[s] = v;
                    sum += v;
                }
                row["合计"] = sum;
                table.Rows.Add(row);
            }

            _previewTable = table;
            GridPreview.ItemsSource = _previewTable.DefaultView;
        }

        private bool ValidateInputs(out int minEach, out double maxFrac)
        {
            minEach = 0; maxFrac = 0;
            if (!int.TryParse(TxtMinEach.Text, out minEach) || minEach < 0)
            {
                MessageBox.Show("最小单科分请输入 >=0 的整数。");
                return false;
            }
            if (!double.TryParse(TxtMaxFrac.Text, out maxFrac) || maxFrac <= 0 || maxFrac > 1)
            {
                MessageBox.Show("单科最大占比请输入 (0,1] 的小数。");
                return false;
            }
            return true;
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            if (_previewTable == null || _previewTable.Rows.Count == 0)
            {
                MessageBox.Show("请先生成预览。");
                return;
            }
            if (string.IsNullOrEmpty(_excelPath))
            {
                MessageBox.Show("未选择源 Excel。");
                return;
            }

            var dir = Path.GetDirectoryName(_excelPath)!;
            var name = Path.GetFileNameWithoutExtension(_excelPath);
            var outPath = Path.Combine(dir, $"{name}_随机分配结果.xlsx");

            try
            {
                using var wb = new XLWorkbook();
                var ws = wb.AddWorksheet("随机分配");
                // 写入表头
                for (int c = 0; c < _previewTable.Columns.Count; c++)
                    ws.Cell(1, c + 1).Value = _previewTable.Columns[c].ColumnName;
                // 写入行
                for (int r = 0; r < _previewTable.Rows.Count; r++)
                {
                    for (int c = 0; c < _previewTable.Columns.Count; c++)
                    {
                        ws.Cell(r + 2, c + 1).SetValue(_previewTable.Rows[r][c]?.ToString());

                    }
                }
                ws.Columns().AdjustToContents();
                wb.SaveAs(outPath);
                MessageBox.Show($"导出完成：\n{outPath}");
            }
            catch (Exception ex)
            {
                MessageBox.Show("导出失败：" + ex.Message);
            }
        }
    }
}
