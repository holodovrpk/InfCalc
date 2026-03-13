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
    public partial class Table5Window : Window
    {
        private readonly List<Table5Row> _rows;

        public Table5Window(List<Table5Row> rows)
        {
            InitializeComponent();
            _rows = rows ?? new List<Table5Row>();
            Table5DataGrid.ItemsSource = _rows;
        }

        private void ExportExcelButton_Click(object sender, RoutedEventArgs e)
        {
            if (_rows.Count == 0)
            {
                MessageBox.Show("Нет данных для экспорта.", "Экспорт", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var dialog = new SaveFileDialog
            {
                Title = "Сохранить Excel-файл",
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                FileName = "Таблица 5 - проведение учений и тренировок.xlsx"
            };

            if (dialog.ShowDialog() != true)
                return;

            try
            {
                ExportToExcel(dialog.FileName);
                MessageBox.Show("Файл Excel успешно сохранён.", "Готово", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Excel:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportToExcel(string filePath)
        {
            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Таблица 5");

            ws.Cell(1, 1).Value = "Проведение учений и тренировок с целью отработки и закрепления навыков реагирования в случае совершения (угрозы совершения) вооруженного нападения";
            ws.Range(1, 1, 1, 13).Merge();

            var titleRange = ws.Range(1, 1, 1, 13);
            titleRange.Style.Font.Bold = true;
            titleRange.Style.Font.FontSize = 14;
            titleRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            titleRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            titleRange.Style.Alignment.WrapText = true;

            ws.Cell(3, 1).Value = "Субъект Российской Федерации";
            ws.Cell(3, 2).Value = "Тип образовательной организации";
            ws.Cell(3, 3).Value = "Всего объектов (территорий)";
            ws.Cell(3, 4).Value = "Кол-во объектов (территорий), где проведены тренировки";
            ws.Cell(3, 5).Value = "Общее количество проведенных тренировок";
            ws.Cell(3, 6).Value = "Руководители образовательных организаций и их заместители, чел.";
            ws.Cell(3, 7).Value = "% от общего кол-ва";
            ws.Cell(3, 8).Value = "Работники, чел.";
            ws.Cell(3, 9).Value = "% от общего кол-ва";
            ws.Cell(3, 10).Value = "Обучающиеся";
            ws.Cell(3, 11).Value = "% от общего кол-ва";
            ws.Cell(3, 12).Value = "Работники охраны, чел.";
            ws.Cell(3, 13).Value = "% от общего кол-ва";

            for (int i = 1; i <= 13; i++)
                ws.Cell(4, i).Value = i.ToString();

            var headerRange1 = ws.Range(3, 1, 3, 13);
            headerRange1.Style.Font.Bold = true;
            headerRange1.Style.Alignment.WrapText = true;
            headerRange1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRange1.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            headerRange1.Style.Fill.BackgroundColor = XLColor.FromHtml("#E8EEF9");
            headerRange1.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            headerRange1.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            var headerRange2 = ws.Range(4, 1, 4, 13);
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
                ws.Cell(currentRow, 4).Value = row.TrainingObjects;
                ws.Cell(currentRow, 5).Value = row.TotalTrainings;
                ws.Cell(currentRow, 6).Value = row.ManagersCount;
                ws.Cell(currentRow, 7).Value = row.ManagersPercent;
                ws.Cell(currentRow, 8).Value = row.WorkersCount;
                ws.Cell(currentRow, 9).Value = row.WorkersPercent;
                ws.Cell(currentRow, 10).Value = row.StudentsCount;
                ws.Cell(currentRow, 11).Value = row.StudentsPercent;
                ws.Cell(currentRow, 12).Value = row.SecurityCount;
                ws.Cell(currentRow, 13).Value = row.SecurityPercent;

                var dataRange = ws.Range(currentRow, 1, currentRow, 13);
                dataRange.Style.Alignment.WrapText = true;
                dataRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                for (int col = 2; col <= 13; col++)
                    ws.Cell(currentRow, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                ws.Cell(currentRow, 7).Style.NumberFormat.Format = "0.00";
                ws.Cell(currentRow, 9).Style.NumberFormat.Format = "0.00";
                ws.Cell(currentRow, 11).Style.NumberFormat.Format = "0.00";
                ws.Cell(currentRow, 13).Style.NumberFormat.Format = "0.00";

                if (string.Equals(row.OrganizationType, "ВСЕГО", StringComparison.OrdinalIgnoreCase))
                    dataRange.Style.Font.Bold = true;

                currentRow++;
            }

            ws.Range(dataStartRow, 1, currentRow - 1, 1).Merge();
            ws.Cell(dataStartRow, 1).Value = "Воронежская область";
            ws.Cell(dataStartRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(dataStartRow, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Cell(dataStartRow, 1).Style.Alignment.WrapText = true;
            ws.Cell(dataStartRow, 1).Style.Font.Bold = true;

            ws.Column(1).Width = 20;
            ws.Column(2).Width = 28;
            ws.Column(3).Width = 16;
            ws.Column(4).Width = 18;
            ws.Column(5).Width = 18;
            ws.Column(6).Width = 18;
            ws.Column(7).Width = 12;
            ws.Column(8).Width = 14;
            ws.Column(9).Width = 12;
            ws.Column(10).Width = 14;
            ws.Column(11).Width = 12;
            ws.Column(12).Width = 14;
            ws.Column(13).Width = 12;

            ws.Row(1).Height = 40;
            ws.Row(3).Height = 85;
            ws.Row(4).Height = 22;

            for (int i = dataStartRow; i < currentRow; i++)
                ws.Row(i).AdjustToContents();

            var usedRange = ws.Range(3, 1, currentRow - 1, 13);
            usedRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            usedRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(filePath)!);
            workbook.SaveAs(filePath);
        }
    }
}