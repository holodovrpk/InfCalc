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
    /// <summary>
    /// Логика взаимодействия для Table4Window.xaml
    /// </summary>
    public partial class Table4Window : Window
    {
        private readonly List<Table4Row> _rows;

        public Table4Window(List<Table4Row> rows)
        {
            InitializeComponent();
            _rows = rows ?? new List<Table4Row>();
            Table4DataGrid.ItemsSource = _rows;
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
                FileName = "Таблица 4 - обучение по профстандарту.xlsx"
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
            var ws = workbook.Worksheets.Add("Таблица 4");

            ws.Cell(1, 1).Value =
                "Информация об обучении должностных лиц органов местного самоуправления, осуществляющих управление в сфере образования (органы управления образованием), и образовательных организаций в соответствии с профессиональным стандартом «Специалист по обеспечению АТЗ объекта (территории)», утвержденным приказом Минтруда России от 27 апреля 2023 г. № 374н (Профстандарт)";
            ws.Range(1, 1, 1, 10).Merge();

            var titleRange = ws.Range(1, 1, 1, 10);
            titleRange.Style.Font.Bold = true;
            titleRange.Style.Font.FontSize = 13;
            titleRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            titleRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            titleRange.Style.Alignment.WrapText = true;

            ws.Cell(3, 1).Value = "Субъект Российской Федерации";
            ws.Cell(3, 2).Value = "Наименование муниципального образования";
            ws.Cell(3, 3).Value = "Количество должностных лиц органа управления образованием, обученных в соответствии с Профстандартом";
            ws.Cell(3, 4).Value = "Количество должностных лиц органа управления образованием, обучение которых запланировано в текущем году";
            ws.Cell(3, 5).Value = "Количество должностных лиц органа управления образованием, обучение которых запланировано в следующем году";
            ws.Cell(3, 6).Value = "Тип образовательной организации";
            ws.Cell(3, 7).Value = "Всего объектов (территорий)";
            ws.Cell(3, 8).Value = "Количество должностных лиц образовательных организаций, обученных в соответствии с Профстандартом";
            ws.Cell(3, 9).Value = "Количество должностных лиц образовательных организаций, обучение которых запланировано в текущем году";
            ws.Cell(3, 10).Value = "Количество должностных лиц образовательных организаций, обучение которых запланировано в следующем году";

            for (int i = 1; i <= 10; i++)
                ws.Cell(4, i).Value = i.ToString();

            var headerRange1 = ws.Range(3, 1, 3, 10);
            headerRange1.Style.Font.Bold = true;
            headerRange1.Style.Alignment.WrapText = true;
            headerRange1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRange1.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            headerRange1.Style.Fill.BackgroundColor = XLColor.FromHtml("#E8EEF9");
            headerRange1.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            headerRange1.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            var headerRange2 = ws.Range(4, 1, 4, 10);
            headerRange2.Style.Font.Bold = true;
            headerRange2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRange2.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            headerRange2.Style.Fill.BackgroundColor = XLColor.FromHtml("#F8FAFC");
            headerRange2.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            headerRange2.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            int dataStartRow = 5;
            int currentRow = dataStartRow;

            var municipalityGroups = _rows
                .Where(r => !r.IsGrandTotalRow)
                .GroupBy(r => r.Municipality)
                .ToList();

            foreach (var municipalityGroup in municipalityGroups)
            {
                var groupRows = municipalityGroup.ToList();
                int groupStartRow = currentRow;

                foreach (var row in groupRows)
                {
                    ws.Cell(currentRow, 6).Value = row.OrganizationType;
                    ws.Cell(currentRow, 7).Value = row.TotalObjects;
                    ws.Cell(currentRow, 8).Value = row.OrgTrained;
                    ws.Cell(currentRow, 9).Value = row.OrgPlannedCurrentYear;
                    ws.Cell(currentRow, 10).Value = row.OrgPlannedNextYear;

                    var dataRange = ws.Range(currentRow, 1, currentRow, 10);
                    dataRange.Style.Alignment.WrapText = true;
                    dataRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                    for (int col = 6; col <= 10; col++)
                        ws.Cell(currentRow, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    if (row.IsMunicipalityTotalRow)
                        dataRange.Style.Font.Bold = true;

                    currentRow++;
                }

                int groupEndRow = currentRow - 1;
                var firstRow = groupRows.First();

                ws.Range(groupStartRow, 2, groupEndRow, 2).Merge();
                ws.Range(groupStartRow, 3, groupEndRow, 3).Merge();
                ws.Range(groupStartRow, 4, groupEndRow, 4).Merge();
                ws.Range(groupStartRow, 5, groupEndRow, 5).Merge();

                ws.Cell(groupStartRow, 2).Value = firstRow.Municipality;
                ws.Cell(groupStartRow, 3).Value = firstRow.MunicipalTrained;
                ws.Cell(groupStartRow, 4).Value = firstRow.MunicipalPlannedCurrentYear;
                ws.Cell(groupStartRow, 5).Value = firstRow.MunicipalPlannedNextYear;

                for (int col = 2; col <= 5; col++)
                {
                    ws.Cell(groupStartRow, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    ws.Cell(groupStartRow, col).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    ws.Cell(groupStartRow, col).Style.Alignment.WrapText = true;
                }
            }

            var grandTotal = _rows.FirstOrDefault(r => r.IsGrandTotalRow);
            if (grandTotal != null)
            {
                ws.Cell(currentRow, 2).Value = "ИТОГО";
                ws.Cell(currentRow, 3).Value = grandTotal.MunicipalTrained;
                ws.Cell(currentRow, 4).Value = grandTotal.MunicipalPlannedCurrentYear;
                ws.Cell(currentRow, 5).Value = grandTotal.MunicipalPlannedNextYear;
                ws.Cell(currentRow, 6).Value = "ИТОГО";
                ws.Cell(currentRow, 7).Value = grandTotal.TotalObjects;
                ws.Cell(currentRow, 8).Value = grandTotal.OrgTrained;
                ws.Cell(currentRow, 9).Value = grandTotal.OrgPlannedCurrentYear;
                ws.Cell(currentRow, 10).Value = grandTotal.OrgPlannedNextYear;

                var totalRange = ws.Range(currentRow, 1, currentRow, 10);
                totalRange.Style.Font.Bold = true;
                totalRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#E0ECFF");
                totalRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                totalRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                totalRange.Style.Alignment.WrapText = true;
                totalRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                totalRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

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
            ws.Column(4).Width = 16;
            ws.Column(5).Width = 16;
            ws.Column(6).Width = 28;
            ws.Column(7).Width = 14;
            ws.Column(8).Width = 16;
            ws.Column(9).Width = 16;
            ws.Column(10).Width = 16;

            ws.Row(1).Height = 42;
            ws.Row(3).Height = 90;
            ws.Row(4).Height = 22;

            for (int i = dataStartRow; i < currentRow; i++)
                ws.Row(i).AdjustToContents();

            var usedRange = ws.Range(3, 1, currentRow - 1, 10);
            usedRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            usedRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(filePath)!);
            workbook.SaveAs(filePath);
        }
    }
}