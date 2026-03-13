using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfCalc
{
    public class EducationRecord
    {
        public List<KeyValuePair<string, string>> Pairs { get; set; } = new();
        public Dictionary<string, string> Fields { get; set; } = new();

        public string Id => GetValue("ID");
        public string Municipality => GetValue("Муниципальное образование");
        public string OrganizationType => GetValue("Тип образовательной организации");
        public string OrganizationName => GetValue("Название образовательной организации");

        public string GetValue(string key)
        {
            return Fields.TryGetValue(key, out var value) ? value : string.Empty;
        }

        public List<string> GetValues(string key)
        {
            return Pairs
                .Where(p => p.Key == key)
                .Select(p => p.Value ?? string.Empty)
                .ToList();
        }

        public string GetValueByOccurrence(string key, int occurrenceIndex)
        {
            var values = GetValues(key);

            if (occurrenceIndex < 0 || occurrenceIndex >= values.Count)
                return string.Empty;

            return values[occurrenceIndex];
        }
    }
}
