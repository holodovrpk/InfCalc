using Microsoft.Win32;
using System.Globalization;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace InfCalc
{
    public partial class MainWindow : Window
    {
        private List<EducationRecord> _records = new();
        private string _currentFilePath = string.Empty;

        private const string PreschoolType = "Дошкольные образовательные организации";
        private const string GeneralType = "Общеобразовательные организации";
        private const string SpoType = "Образовательные организации СПО";

        public MainWindow()
        {
            InitializeComponent();
        }

        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Выберите JSON-файл",
                Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
            };

            if (dialog.ShowDialog() != true)
                return;

            try
            {
                _currentFilePath = dialog.FileName;
                FilePathTextBlock.Text = _currentFilePath;

                _records = JsonDataLoader.Load(_currentFilePath);

                UpdateStatistics();

                MessageBox.Show("Файл успешно импортирован.", "Готово", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке файла:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UpdateStatistics()
        {
            TotalRecordsTextBlock.Text = _records.Count.ToString();

            MunicipalitiesTextBlock.Text = _records
                .Select(r => r.Municipality?.Trim())
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Count()
                .ToString();

            PreschoolTextBlock.Text = _records.Count(r =>
                string.Equals(r.OrganizationType?.Trim(), PreschoolType, StringComparison.OrdinalIgnoreCase)).ToString();

            GeneralTextBlock.Text = _records.Count(r =>
                string.Equals(r.OrganizationType?.Trim(), GeneralType, StringComparison.OrdinalIgnoreCase)).ToString();

            SpoTextBlock.Text = _records.Count(r =>
                string.Equals(r.OrganizationType?.Trim(), SpoType, StringComparison.OrdinalIgnoreCase)).ToString();
        }

        private bool EnsureDataLoaded()
        {
            if (_records.Count > 0)
                return true;

            MessageBox.Show("Сначала импортируйте JSON-файл.", "Нет данных", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        private void Table1Button_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureDataLoaded())
                return;

            try
            {
                var rows = BuildTable1Rows();

                var window = new Table1Window(rows)
                {
                    Owner = this
                };

                window.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при формировании Таблицы 1:\n{ex.Message}",
                                "Ошибка",
                                MessageBoxButton.OK,
                                MessageBoxImage.Error);
            }
        }

        private void Table2Button_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureDataLoaded())
                return;

            try
            {
                var rows = BuildTable2Rows();

                var window = new Table2Window(rows)
                {
                    Owner = this
                };

                window.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при формировании Таблицы 2:\n{ex.Message}",
                                "Ошибка",
                                MessageBoxButton.OK,
                                MessageBoxImage.Error);
            }
        }

        private void Table3Button_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureDataLoaded())
                return;

            try
            {
                var rows = BuildTable3Rows();

                var window = new Table3Window(rows)
                {
                    Owner = this
                };

                window.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при формировании Таблицы 3:\n{ex.Message}",
                                "Ошибка",
                                MessageBoxButton.OK,
                                MessageBoxImage.Error);
            }
        }

        private void Table4Button_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureDataLoaded()) return;
            MessageBox.Show("Таблица 4 будет реализована следующим шагом.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void Table5Button_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureDataLoaded())
                return;

            try
            {
                var rows = BuildTable5Rows();

                var window = new Table5Window(rows)
                {
                    Owner = this
                };

                window.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при формировании Таблицы 5:\n{ex.Message}",
                                "Ошибка",
                                MessageBoxButton.OK,
                                MessageBoxImage.Error);
            }
        }


        private static int ParseInt(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return 0;

            value = value.Trim()
                         .Replace(" ", "")
                         .Replace("\u00A0", "");

            if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int result))
                return result;

            if (int.TryParse(value, NumberStyles.Integer, CultureInfo.GetCultureInfo("ru-RU"), out result))
                return result;

            return 0;
        }

        private List<Table1Row> BuildTable1Rows()
        {
            const string preschool = "Дошкольные образовательные организации";
            const string general = "Общеобразовательные организации";
            const string spo = "Образовательные организации СПО";

            var typesInOrder = new List<string>
    {
        preschool,
        general,
        spo
    };

            var rows = new List<Table1Row>();

            foreach (var type in typesInOrder)
            {
                var filtered = _records.Where(r =>
                    string.Equals(r.OrganizationType?.Trim(), type, StringComparison.OrdinalIgnoreCase));

                var row = new Table1Row
                {
                    OrganizationType = type,
                    TotalObjects = filtered.Sum(r => ParseInt(r.GetValue("Всего объектов (территорий)"))),
                    EquippedWithTs = filtered.Sum(r => ParseInt(r.GetValue("Объекты (территории), оснащенные системами передачи ТС"))),
                    ValidCalls = filtered.Sum(r => ParseInt(r.GetValue("Количество обоснованных вызовов оперативных служб и частных охранных организаций"))),
                    PreventedOffenses = filtered.Sum(r => ParseInt(r.GetValue("Количество правонарушений, предотвращенных либо пресеченных в результате передачи ТС"))),
                    DetainedPersons = filtered.Sum(r => ParseInt(r.GetValue("Количество лиц, задержанных в результате выезда оперслужб при получении ТС")))
                };

                rows.Add(row);
            }

            rows.Add(new Table1Row
            {
                OrganizationType = "ВСЕГО",
                TotalObjects = rows.Sum(r => r.TotalObjects),
                EquippedWithTs = rows.Sum(r => r.EquippedWithTs),
                ValidCalls = rows.Sum(r => r.ValidCalls),
                PreventedOffenses = rows.Sum(r => r.PreventedOffenses),
                DetainedPersons = rows.Sum(r => r.DetainedPersons)
            });

            return rows;
        }
        private List<Table2Row> BuildTable2Rows()
        {
            const string preschool = "Дошкольные образовательные организации";
            const string general = "Общеобразовательные организации";
            const string spo = "Образовательные организации СПО";

            var typesInOrder = new List<string>
    {
        preschool,
        general,
        spo
    };

            var rows = new List<Table2Row>();

            foreach (var type in typesInOrder)
            {
                var filtered = _records.Where(r =>
                    string.Equals(r.OrganizationType?.Trim(), type, StringComparison.OrdinalIgnoreCase));

                var row = new Table2Row
                {
                    OrganizationType = type,
                    TotalObjects = filtered.Sum(r => ParseInt(r.GetValue("Всего объектов (территорий)"))),
                    EquippedWithSoue = filtered.Sum(r => ParseInt(r.GetValue("Объекты (территории) оснащенные СОУЭ либо автономными средствами оповещения"))),
                    ValidSoueActivations = filtered.Sum(r => ParseInt(r.GetValue("Количество обоснованных включений СОУЭ либо автономных средств оповещения (за исключением тренировок и учений)")))
                };

                rows.Add(row);
            }

            rows.Add(new Table2Row
            {
                OrganizationType = "ВСЕГО",
                TotalObjects = rows.Sum(r => r.TotalObjects),
                EquippedWithSoue = rows.Sum(r => r.EquippedWithSoue),
                ValidSoueActivations = rows.Sum(r => r.ValidSoueActivations)
            });

            return rows;
        }


        private List<Table3Row> BuildTable3Rows()
        {
            const string preschool = "Дошкольные образовательные организации";
            const string general = "Общеобразовательные организации";
            const string spo = "Образовательные организации СПО";

            var typesInOrder = new List<string>
    {
        preschool,
        general,
        spo
    };

            var rows = new List<Table3Row>();

            foreach (var type in typesInOrder)
            {
                var filtered = _records.Where(r =>
                    string.Equals(r.OrganizationType?.Trim(), type, StringComparison.OrdinalIgnoreCase));

                var row = new Table3Row
                {
                    OrganizationType = type,
                    TotalObjects = filtered.Sum(r => ParseInt(r.GetValue("Всего объектов (территорий)"))),
                    HasAlgorithms = filtered.Sum(r => ParseInt(r.GetValue("Объекты (территории), где имеются в наличии Алгоритмы"))),
                    UpdatedAlgorithms = filtered.Sum(r => ParseInt(r.GetValue("Объекты (территории), где  Алгоритмы актуализированы с учетом характеристики зданий, места расположения, фактической оснащенности техническими средствами охраны и тому подобного")))
                };

                rows.Add(row);
            }

            rows.Add(new Table3Row
            {
                OrganizationType = "ВСЕГО",
                TotalObjects = rows.Sum(r => r.TotalObjects),
                HasAlgorithms = rows.Sum(r => r.HasAlgorithms),
                UpdatedAlgorithms = rows.Sum(r => r.UpdatedAlgorithms)
            });

            return rows;
        }

        private static double ParseDouble(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return 0;

            value = value.Trim()
                         .Replace("%", "")
                         .Replace(" ", "")
                         .Replace("\u00A0", "")
                         .Replace(",", ".");

            if (double.TryParse(value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double result))
                return result;

            return 0;
        }

        private static double AverageOrZero(IEnumerable<double> values)
        {
            var list = values.Where(v => v > 0).ToList();
            return list.Count == 0 ? 0 : list.Average();
        }

        private List<Table5Row> BuildTable5Rows()
        {
            const string preschool = "Дошкольные образовательные организации";
            const string general = "Общеобразовательные организации";
            const string spo = "Образовательные организации СПО";

            var typesInOrder = new List<string>
    {
        preschool,
        general,
        spo
    };

            var rows = new List<Table5Row>();

            foreach (var type in typesInOrder)
            {
                var filtered = _records.Where(r =>
                    string.Equals(r.OrganizationType?.Trim(), type, StringComparison.OrdinalIgnoreCase)).ToList();

                var row = new Table5Row
                {
                    OrganizationType = type,
                    TotalObjects = filtered.Sum(r => ParseInt(r.GetValue("Всего объектов (территорий)"))),
                    TrainingObjects = filtered.Sum(r => ParseInt(r.GetValue("Количество объектов (территорий), где  проведены тренировки"))),
                    TotalTrainings = filtered.Sum(r => ParseInt(r.GetValue("Общее количество проведенных тренировок"))),
                    ManagersCount = filtered.Sum(r => ParseInt(r.GetValue("руководители образовательных организаций и их заместители, чел."))),
                    ManagersPercent = AverageOrZero(filtered.Select(r => ParseDouble(r.GetValueByOccurrence("% от общего количества", 0)))),
                    WorkersCount = filtered.Sum(r => ParseInt(r.GetValue("работники, чел."))),
                    WorkersPercent = AverageOrZero(filtered.Select(r => ParseDouble(r.GetValueByOccurrence("% от общего количества", 1)))),
                    StudentsCount = filtered.Sum(r => ParseInt(r.GetValue("обучающиеся"))),
                    StudentsPercent = AverageOrZero(filtered.Select(r => ParseDouble(r.GetValueByOccurrence("% от общего количества", 2)))),
                    SecurityCount = filtered.Sum(r => ParseInt(r.GetValue("работники охраны, чел."))),
                    SecurityPercent = AverageOrZero(filtered.Select(r => ParseDouble(r.GetValueByOccurrence("% от общего количества", 3))))
                };

                rows.Add(row);
            }

            rows.Add(new Table5Row
            {
                OrganizationType = "ВСЕГО",
                TotalObjects = rows.Sum(r => r.TotalObjects),
                TrainingObjects = rows.Sum(r => r.TrainingObjects),
                TotalTrainings = rows.Sum(r => r.TotalTrainings),
                ManagersCount = rows.Sum(r => r.ManagersCount),
                ManagersPercent = AverageOrZero(rows.Select(r => r.ManagersPercent)),
                WorkersCount = rows.Sum(r => r.WorkersCount),
                WorkersPercent = AverageOrZero(rows.Select(r => r.WorkersPercent)),
                StudentsCount = rows.Sum(r => r.StudentsCount),
                StudentsPercent = AverageOrZero(rows.Select(r => r.StudentsPercent)),
                SecurityCount = rows.Sum(r => r.SecurityCount),
                SecurityPercent = AverageOrZero(rows.Select(r => r.SecurityPercent))
            });

            return rows;
        }

    }
}