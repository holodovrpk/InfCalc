using ClosedXML.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace InfCalc
{
    
    public partial class Table2Window : Window
    {
        private readonly List<Table2Row> _rows;

        public Table2Window(List<Table2Row> rows)
        {
            InitializeComponent();

            _rows = rows ?? new List<Table2Row>();
            Table2DataGrid.ItemsSource = _rows;
        }

        private void ExportExcelButton_Click(object sender, RoutedEventArgs e)
        {
            if (_rows.Count == 0)
            {
                MessageBox.Show("Нет данных для экспорта.",
                                "Экспорт",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning);
                return;
            }

            var dialog = new SaveFileDialog
            {
                Title = "Сохранить Excel-файл",
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                FileName = "Таблица 2 - эффективность СОУЭ.xlsx"
            };

            if (dialog.ShowDialog() != true)
                return;

            try
            {
                ExportToExcel(dialog.FileName);

                MessageBox.Show("Файл Excel успешно сохранён.",
                                "Готово",
                                MessageBoxButton.OK,
                                MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Excel:\n{ex.Message}",
                                "Ошибка",
                                MessageBoxButton.OK,
                                MessageBoxImage.Error);
            }
        }

        private void ExportToExcel(string filePath)
        {
            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Таблица 2");

            ws.Cell(1, 1).Value = "Информация об эффективности функционирования систем оповещения и управления эвакуацией (СОУЭ) либо автономных средств оповещения";
            ws.Range(1, 1, 1, 5).Merge();

            var titleRange = ws.Range(1, 1, 1, 5);
            titleRange.Style.Font.Bold = true;
            titleRange.Style.Font.FontSize = 16;
            titleRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            titleRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            titleRange.Style.Alignment.WrapText = true;

            // Шапка
            ws.Cell(3, 1).Value = "Субъект Российской Федерации";
            ws.Cell(3, 2).Value = "Тип образовательной организации";
            ws.Cell(3, 3).Value = "Всего объектов (территорий)";
            ws.Cell(3, 4).Value = "Объекты (территории), оснащенные СОУЭ либо автономными средствами оповещения";
            ws.Cell(3, 5).Value = "Количество обоснованных включений СОУЭ либо автономных средств оповещения (за исключением тренировок и учений)";

            // Номера столбцов
            ws.Cell(4, 1).Value = "1";
            ws.Cell(4, 2).Value = "2";
            ws.Cell(4, 3).Value = "3";
            ws.Cell(4, 4).Value = "4";
            ws.Cell(4, 5).Value = "5";

            var headerRange1 = ws.Range(3, 1, 3, 5);
            headerRange1.Style.Font.Bold = true;
            headerRange1.Style.Alignment.WrapText = true;
            headerRange1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRange1.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            headerRange1.Style.Fill.BackgroundColor = XLColor.FromHtml("#E8EEF9");
            headerRange1.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            headerRange1.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            var headerRange2 = ws.Range(4, 1, 4, 5);
            headerRange2.Style.Font.Bold = true;
            headerRange2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRange2.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            headerRange2.Style.Fill.BackgroundColor = XLColor.FromHtml("#F8FAFC");
            headerRange2.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            headerRange2.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            int dataStartRow = 5;
            int currentRow = dataStartRow;

            foreach (var row in _rows)
            {
                ws.Cell(currentRow, 2).Value = row.OrganizationType;
                ws.Cell(currentRow, 3).Value = row.TotalObjects;
                ws.Cell(currentRow, 4).Value = row.EquippedWithSoue;
                ws.Cell(currentRow, 5).Value = row.ValidSoueActivations;

                var dataRange = ws.Range(currentRow, 1, currentRow, 5);
                dataRange.Style.Alignment.WrapText = true;
                dataRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(currentRow, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(currentRow, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(currentRow, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                if (string.Equals(row.OrganizationType, "ВСЕГО", StringComparison.OrdinalIgnoreCase))
                {
                    dataRange.Style.Font.Bold = true;
                }

                currentRow++;
            }

            // Объединённый первый столбец с субъектом РФ
            ws.Range(dataStartRow, 1, currentRow - 1, 1).Merge();
            ws.Cell(dataStartRow, 1).Value = "Воронежская область";
            ws.Cell(dataStartRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(dataStartRow, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Cell(dataStartRow, 1).Style.Alignment.WrapText = true;
            ws.Cell(dataStartRow, 1).Style.Font.Bold = true;

            ws.Column(1).Width = 24;
            ws.Column(2).Width = 40;
            ws.Column(3).Width = 20;
            ws.Column(4).Width = 32;
            ws.Column(5).Width = 42;

            ws.Row(1).Height = 42;
            ws.Row(3).Height = 85;
            ws.Row(4).Height = 22;

            for (int i = dataStartRow; i < currentRow; i++)
            {
                ws.Row(i).AdjustToContents();
            }

            var usedRange = ws.Range(3, 1, currentRow - 1, 5);
            usedRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            usedRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(filePath)!);
            workbook.SaveAs(filePath);
        }
    }
}