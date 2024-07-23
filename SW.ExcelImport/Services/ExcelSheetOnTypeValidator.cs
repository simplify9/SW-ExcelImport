using System.Linq;
using System.Collections.Generic;
using System;
using System.Threading.Tasks;
using ExcelDataReader;
using System.IO;

namespace SW.ExcelImport.Services
{

    public class ExcelSheetOnTypeValidator
    {
        private bool AllowedEnumerableInSheet (Type type) =>
            type == typeof(string) || type.IsValueType;
        public Task<SheetValidationResult> Validate(SheetOnTypeParseRequest request)
        {
            var sheet = request.Sheet;
            if (sheet.Empty || sheet.EmptyData)
                return Task.FromResult(SheetValidationResult.EmptySheet());

            var invalidName = false;
            var invalidHeaders = new List<int>();
            Type type = null;

            var sheetOptions = request.MappingOptions;
            var (ignoreFirstRow, headerMap) = GetMap(sheet, sheetOptions);

            if (sheet.Index == 0)
            {
                type = request.RootType;
                invalidName = false;
            }
            else
            {
                type = GetEnumerableType(request.RootType,
                    request.LongName ?? sheet.Name, request.NamingStrategy);
                invalidName = type == null;
            }

            if (!invalidName)
            {
                var idHeaders = new List<int>();
                if(request.MappingOptions.IdIndex.HasValue)
                    idHeaders.Add(request.MappingOptions.IdIndex.Value);
                
                if(request.MappingOptions.ParentIdIndex.HasValue)
                    idHeaders.Add(request.MappingOptions.ParentIdIndex.Value);

                for (var i = 0; i < headerMap.Length; i++)
                {
                    
                    if (idHeaders.Contains(i)) continue;
                    if (string.IsNullOrEmpty(headerMap[i]))
                    {
                        invalidHeaders.Add(i);
                        continue;
                    }
                    var propertyName = headerMap[i].Transform(request.NamingStrategy);
                    var propertyPath = PropertyPath.TryParse(type, propertyName);
                    if (propertyPath == null) 
                        invalidHeaders.Add(i);
                    else
                    {
                        var enumerableType = GetEnumerableType(type, propertyName, request.NamingStrategy);
                        if(enumerableType != null && ! AllowedEnumerableInSheet(enumerableType)  )
                            invalidHeaders.Add(i);
                    }

                }
            }

            var result = new SheetValidationResult(headerMap, ignoreFirstRow, invalidName, invalidHeaders.ToArray());
            return Task.FromResult(result);

            

        }


        private static (bool, string[]) GetMap(ISheet sheet, SheetMappingOptions options) =>
            (options.Map == null, options.Map ?? sheet.Header.Select(x => x.Value?.ToString()).ToArray());

        private Type GetEnumerableType(Type rootType, string propertyName, JsonNamingStrategy strategy)
        {

            var propertyPath = PropertyPath.TryParse(rootType, propertyName.Transform(strategy));
            if (propertyPath == null)
                return null;
            else
                return rootType.GetEnumerablePropertyType(propertyName, strategy);

        }

        public SheetValidationResult ValidateCustom<T>(SheetOnTypeParseRequest request)
        {
            var sheet = request.Sheet;
            if (sheet.Empty || sheet.EmptyData)
                return SheetValidationResult.EmptySheet();

            var invalidPairs = new List<KeyValuePair<string, string>>();
            var customMap = request.MappingOptions.CustomMap;
            foreach (var pair in customMap)
            {
                if(string.IsNullOrWhiteSpace(pair.Key) ||  string.IsNullOrWhiteSpace(pair.Value))
                    invalidPairs.Add(pair);
                else
                {
                    var propertyName = pair.Value.Transform(request.NamingStrategy);
                    var propertyPath = PropertyPath.TryParse(typeof(T), propertyName);
                    if (propertyPath == null) 
                        invalidPairs.Add(pair);
                    else
                    {
                        var enumerableType = GetEnumerableType(typeof(T), propertyName, request.NamingStrategy);
                        if(enumerableType != null && ! AllowedEnumerableInSheet(enumerableType)  )
                            invalidPairs.Add(pair);
                    }
                }
            }
            

            

            return new SheetValidationResult(invalidPairs.ToArray());
            

            

        }
    }
}