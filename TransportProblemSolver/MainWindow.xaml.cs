using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using ClosedXML.Excel;
using System.Windows.Controls;

namespace TransportProblemSolver
{
    public partial class MainWindow : Window
    {
        private int[,] cost;
        private List<int> Osupply;
        private List<int> Odemand;
        private List<TransportResult> solution;
        private int rowCount = 4;
        private int columnCount = 4;

        public MainWindow()
        {
            InitializeComponent();
            RowSizeComboBox.SelectedIndex = 2;
            ColumnSizeComboBox.SelectedIndex = 2;
        }

        private void SizeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (RowSizeComboBox.SelectedItem != null)
            {
                rowCount = int.Parse(((ComboBoxItem)RowSizeComboBox.SelectedItem).Content.ToString());
            }

            if (ColumnSizeComboBox.SelectedItem != null)
            {
                columnCount = int.Parse(((ComboBoxItem)ColumnSizeComboBox.SelectedItem).Content.ToString());
            }
        }

        private void SolveButton_Click(object sender, RoutedEventArgs e)
        {
            if (cost == null || Osupply == null || Odemand == null)
            {
                MessageBox.Show("Пожалуйста, загрузите данные из файла перед решением задачи.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            List<int> _s = new(Osupply);
            List<int> _d = new(Odemand);
            solution = SolveTransportProblem(cost, _s, _d);

            ResultDataGrid.ItemsSource = solution.Select(r => new
            {
                From = r.From,
                To = r.To,
                Quantity = r.Quantity,
                Cost = r.Cost
            }).ToList();
        }

        private void UploadButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Файлы Excel|*.xlsx;*.xls"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    using (var workbook = new XLWorkbook(openFileDialog.FileName))
                    {
                        var worksheet = workbook.Worksheet(1);
                        cost = new int[rowCount, columnCount];
                        Osupply = new List<int>();
                        Odemand = new List<int>();

                        for (int i = 0; i < rowCount; i++)
                        {
                            for (int j = 0; j < columnCount; j++)
                            {
                                cost[i, j] = (int)worksheet.Cell(i + 2, j + 2).GetValue<double>();
                            }
                        }

                        for (int i = 0; i < rowCount; i++)
                        {
                            Osupply.Add((int)worksheet.Cell(i + 2, columnCount + 2).GetValue<double>());
                        }

                        for (int j = 0; j < columnCount; j++)
                        {
                            Odemand.Add((int)worksheet.Cell(rowCount + 2, j + 2).GetValue<double>());
                        }

                        MessageBox.Show("Данные успешно загружены.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при чтении файла: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (solution == null || !solution.Any())
            {
                MessageBox.Show("Сначала решите задачу, чтобы сохранить результат.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Файлы Excel|*.xlsx;*.xls"
            };
            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Solution");

                        worksheet.Cell(1, 1).Value = "Поставщики\\Потребители";
                        for (int j = 0; j < columnCount; j++)
                        {
                            worksheet.Cell(1, j + 2).Value = $"Потребитель {j + 1}";
                        }
                        worksheet.Cell(1, columnCount + 2).Value = "Использовано";

                        for (int i = 0; i < rowCount; i++)
                        {
                            worksheet.Cell(i + 2, 1).Value = $"Поставщик {i + 1}";
                            for (int j = 0; j < columnCount; j++)
                            {
                                var cell = worksheet.Cell(i + 2, j + 2);
                                var result = solution.FirstOrDefault(r => r.From == $"Запасы {i + 1}" && r.To == $"Потребность {j + 1}");
                                cell.Value = result != null ? result.Quantity : 0;
                            }
                        }

                        worksheet.Cell(rowCount + 2, 1).Value = "Объем доставки (шт)";
                        for (int j = 0; j < columnCount; j++)
                        {
                            worksheet.Cell(rowCount + 2, j + 2).FormulaA1 = $"SUM({worksheet.Cell(2, j + 2).Address}:{worksheet.Cell(rowCount + 1, j + 2).Address})";
                            worksheet.Cell(rowCount + 2, j + 2).Style.Fill.SetBackgroundColor(XLColor.LightGray);
                        }

                        for (int i = 0; i < rowCount; i++)
                        {
                            worksheet.Cell(i + 2, columnCount + 2).FormulaA1 = $"SUM({worksheet.Cell(i + 2, 2).Address}:{worksheet.Cell(i + 2, columnCount + 1).Address})";
                            worksheet.Cell(i + 2, columnCount + 2).Style.Fill.SetBackgroundColor(XLColor.LightGray);
                        }

                        worksheet.Cell(rowCount + 3, 1).Value = "Потребность";
                        for (int j = 0; j < columnCount; j++)
                        {
                            worksheet.Cell(rowCount + 3, j + 2).Value = Odemand[j];
                            worksheet.Cell(rowCount + 3, j + 2).Style.Fill.SetBackgroundColor(XLColor.Gray);
                        }

                        worksheet.Cell(1, columnCount + 3).Value = "Запасы";
                        for (int i = 0; i < rowCount; i++)
                        {
                            worksheet.Cell(i + 2, columnCount + 3).Value = Osupply[i];
                            worksheet.Cell(i + 2, columnCount + 3).Style.Fill.SetBackgroundColor(XLColor.Gray);
                        }

                        worksheet.Cell(rowCount + 5, 1).Value = "F(x)";
                        worksheet.Cell(rowCount + 5, 2).Value = solution.Sum(r => r.Cost);
                        worksheet.Cell(rowCount + 5, 2).Style.Fill.SetBackgroundColor(XLColor.LightGreen);
                        worksheet.Columns().AdjustToContents();
                        AddBordersToUsedRange(worksheet);

                        workbook.SaveAs(saveFileDialog.FileName);
                        MessageBox.Show("Результат успешно сохранен.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при сохранении файла: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private List<TransportResult> SolveTransportProblem(int[,] cost, List<int> supply, List<int> demand)
        {
            List<TransportResult> result = new List<TransportResult>();
            int m = cost.GetLength(0);
            int n = cost.GetLength(1);
            int[,] x = new int[m, n];
            bool[] u = new bool[m];
            bool[] v = new bool[n];
            int[] uValues = new int[m];
            int[] vValues = new int[n];

            int i = 0, j = 0;
            while (i < m && j < n)
            {
                int min = Math.Min(supply[i], demand[j]);
                x[i, j] = min;
                supply[i] -= min;
                demand[j] -= min;
                if (supply[i] == 0) i++;
                if (demand[j] == 0) j++;
            }

            u[0] = true;
            while (true)
            {
                bool updated = false;
                for (i = 0; i < m; i++)
                {
                    for (j = 0; j < n; j++)
                    {
                        if (x[i, j] > 0)
                        {
                            if (u[i] && !v[j])
                            {
                                v[j] = true;
                                vValues[j] = cost[i, j] - uValues[i];
                                updated = true;
                            }
                            else if (v[j] && !u[i])
                            {
                                u[i] = true;
                                uValues[i] = cost[i, j] - vValues[j];
                                updated = true;
                            }
                        }
                    }
                }
                if (!updated) break;
            }

            for (i = 0; i < m; i++)
            {
                for (j = 0; j < n; j++)
                {
                    if (x[i, j] > 0)
                    {
                        result.Add(new TransportResult
                        {
                            From = $"Запасы {i + 1}",
                            To = $"Потребность {j + 1}",
                            Quantity = x[i, j],
                            Cost = x[i, j] * cost[i, j]
                        });
                    }
                }
            }

            return result;
        }

        private void AddBordersToUsedRange(IXLWorksheet worksheet)
        {
            var range = worksheet.RangeUsed();

            if (range != null)
            {
                foreach (IXLCell cell in range.Cells())
                {
                    if (!cell.Value.IsBlank)
                    {
                        cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        cell.Style.Border.OutsideBorderColor = XLColor.Black;
                    }
                    if (cell.Value.IsText)
                    {
                        cell.Style.Font.SetBold();
                    }
                }
            }
        }
    }

    public class TransportResult
    {
        public string From { get; set; }
        public string To { get; set; }
        public int Quantity { get; set; }
        public int Cost { get; set; }
    }
}
