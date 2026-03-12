using Microsoft.Win32;
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
            if (!EnsureDataLoaded()) return;
            MessageBox.Show("Таблица 1 будет реализована следующим шагом.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void Table2Button_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureDataLoaded()) return;
            MessageBox.Show("Таблица 2 будет реализована следующим шагом.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void Table3Button_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureDataLoaded()) return;
            MessageBox.Show("Таблица 3 будет реализована следующим шагом.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void Table4Button_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureDataLoaded()) return;
            MessageBox.Show("Таблица 4 будет реализована следующим шагом.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void Table5Button_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureDataLoaded()) return;
            MessageBox.Show("Таблица 5 будет реализована следующим шагом.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}