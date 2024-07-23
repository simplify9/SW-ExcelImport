using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace SW.ExcelImport.Services
{
    public class SheetReader<T> where T : new()
    {
        readonly IExcelReader reader;
        readonly ExcelSheetOnTypeValidator sheetValidator;
        private bool loaded = false;
        private SheetValidationResult sheetValidationResult;
        private JsonNamingStrategy namingStrategy;
        private bool ignoreFirst;

        public int RowsCount { get; private set; }
        private IDictionary<string, string> _customMap;

        public SheetReader(IExcelReader reader, ExcelSheetOnTypeValidator sheetValidator)
        {
            this.reader = reader;
            this.sheetValidator = sheetValidator;
        }

        public async Task Load(string url, IDictionary<string, string> customMap, int sheetIndex = 0,
            JsonNamingStrategy jsonNamingStrategy = JsonNamingStrategy.None)
        {
            var validationResult = await Validate(url, customMap, sheetIndex, jsonNamingStrategy);
            if (validationResult.HasErrors)
                throw new InvalidOperationException("Sheet is invalid");
        }

        public void Reset()
        {
            reader.Reset();
        }

        public async Task<SheetValidationResult> Validate(string url, IDictionary<string, string> customMap,
            int sheetIndex = 0,
            JsonNamingStrategy jsonNamingStrategy = JsonNamingStrategy.None)
        {
            loaded = true;
            ignoreFirst = true;
            namingStrategy = jsonNamingStrategy;
            _customMap = customMap;
            var sheet = await reader.LoadSheet(url, sheetIndex, _customMap);

            if (sheet == null)
                return SheetValidationResult.SheetNotPresent();
            RowsCount = sheet.RowCount;
            var request = new SheetOnTypeParseRequest
            {
                MappingOptions = new SheetMappingOptions
                {
                    Ignore = ignoreFirst,
                    CustomMap = customMap,
                    IndexAsId = true
                },
                NamingStrategy = namingStrategy,
                RootType = typeof(T),
                Sheet = sheet
            };

            sheetValidationResult = sheetValidator.ValidateCustom<T>(request);
            return sheetValidationResult;
        }


        public async Task<(bool, RowParseResultTyped<T>)> Read()
        {
            if (!loaded)
                throw new InvalidOperationException("Sheet not loaded. Call the Load or Validate sheet first");

            if (sheetValidationResult.HasErrors)
                throw new InvalidOperationException("Sheet is invalid");


            var found = await reader.ReadRow();
            if (reader.Current.Index == 1 && ignoreFirst)
                found = await reader.ReadRow();

            if (!found)
                return (false, null);

            var result = new RowParseResultTyped<T>();

            var invalidCells = new List<int>();
            var values = new Dictionary<string, object>();
            var row = reader.Current;
            var parseOnType = typeof(T);

            var headerMap = row.Sheet.Header.Select(x => x.Value?.ToString()?.Trim()).ToArray();

            for (var i = 0; i < row.Cells.Length; i++)
            {
                if (!_customMap.ContainsKey(headerMap[i] ?? string.Empty))
                    continue;

                var mapName = _customMap[headerMap[i] ?? string.Empty];

                var value = row.Cells[i].Value;

                var propertyName = mapName.Transform(namingStrategy);

                var propertyPath = PropertyPath.TryParse(parseOnType, propertyName);

                var convertSucceeded =
                    Converter.TryCreate(value, propertyPath.PropertyType, out var castValue);

                if (convertSucceeded)
                {
                    if (castValue is string cv && cv== "")
                        castValue = null;
                    values[propertyName] = castValue;
                }
                else
                    invalidCells.Add(i);
            }

            result.InvalidCells = invalidCells.ToArray();
            result.Row = row;
            if (invalidCells.Count == 0)
                result.RowMapped = (T)parseOnType.CreateFromDictionary(values);

            return (true, result);
        }

        public async Task<ICollection<RowParseResultTyped<T>>> ReadAll()
        {
            if (!loaded)
                throw new InvalidOperationException("Sheet not loaded. Call the Load or Validate sheet first");

            var results = new List<RowParseResultTyped<T>>();
            var found = false;
            do
            {
                var (hasResult, result) = await Read();

                found = hasResult;
                if (!found) continue;
                if (result.Row.Cells.All(c => c.Value == null || c.ToString() == "")) continue;
                results.Add(result);
            } while (found);

            return results;
        }
    }
}