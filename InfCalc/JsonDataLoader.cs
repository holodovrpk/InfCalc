using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace InfCalc
{
    public static class JsonDataLoader
    {
        public static List<EducationRecord> Load(string filePath)
        {
            var result = new List<EducationRecord>();

            using FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            using JsonDocument doc = JsonDocument.Parse(fs);

            if (doc.RootElement.ValueKind != JsonValueKind.Array)
                throw new Exception("Корневой элемент JSON должен быть массивом.");

            foreach (JsonElement recordElement in doc.RootElement.EnumerateArray())
            {
                if (recordElement.ValueKind != JsonValueKind.Array)
                    continue;

                var record = new EducationRecord();

                foreach (JsonElement pairElement in recordElement.EnumerateArray())
                {
                    if (pairElement.ValueKind != JsonValueKind.Array || pairElement.GetArrayLength() < 2)
                        continue;

                    string key = pairElement[0].GetString() ?? string.Empty;
                    string value = pairElement[1].GetString() ?? string.Empty;

                    if (!string.IsNullOrWhiteSpace(key))
                    {
                        record.Pairs.Add(new KeyValuePair<string, string>(key, value));
                        record.Fields[key] = value;
                    }
                }

                result.Add(record);
            }

            return result;
        }
    }
}
