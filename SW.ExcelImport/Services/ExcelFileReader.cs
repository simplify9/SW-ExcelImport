using System.ComponentModel.Design;
using System.Xml.Linq;
using System.IO;
using System;
using SW.PrimitiveTypes;
using System.Threading.Tasks;
using ExcelDataReader;
using System.Linq;
using System.Collections.Generic;

namespace SW.ExcelImport.Services
{
    
    public class ExcelFileReader : IExcelReader 
    {
        private IExcelDataReader reader;
        private Stream stream;
        private int specificSheetIndex =-1;
        private bool inSpecificSheet;
        private int index = 0;
        private bool loaded;
        private string url;
        private string tempLocalFilePath;
        readonly ICloudFilesService cloudFilesService;

        public ExcelFileReader(ICloudFilesService cloudFilesService)
        {
            this.cloudFilesService = cloudFilesService;
        }


        public async Task<ISheet> LoadSheet(string fileUrl, int sheetIndex, string[] map = null)
        {
            url = fileUrl;
            specificSheetIndex = sheetIndex;
            tempLocalFilePath = await CopyFileLocal(fileUrl);
            await using var tempStream = File.Open(tempLocalFilePath, FileMode.Open, FileAccess.Read);
            
            using var tmpReader = ExcelReaderFactory.CreateReader(tempStream);

            if(tmpReader.ResultsCount == 0) return null;
            
            if(tmpReader.ResultsCount -1 <  specificSheetIndex) return null;
            
            var sheets = new ISheet[tmpReader.ResultsCount];
            
            var container = new SheetContainer(fileUrl, sheets);
            
            for (var i = 0; i < tmpReader.ResultsCount; i++)
            {
                if (i == specificSheetIndex)
                {
                    var ignoreHeader = map!=null;
                    tmpReader.Read();
                    CurrentSheet = new Sheet(tmpReader, container, i,ignoreHeader);
                    
                }
                else
                {
                    tmpReader.NextResult();
                }
                
            }

            return CurrentSheet;
        }

        public async Task<ISheet> LoadSheet(string fileUrl, int sheetIndex, IDictionary<string, string> map)
        {
            url = fileUrl;
            specificSheetIndex = sheetIndex;
            tempLocalFilePath = await CopyFileLocal(fileUrl);
            await using var tempStream = File.Open(tempLocalFilePath, FileMode.Open, FileAccess.Read);

            using var tmpReader = ExcelReaderFactory.CreateReader(tempStream);

            if (tmpReader.ResultsCount == 0) return null;

            if (tmpReader.ResultsCount - 1 < specificSheetIndex) return null;

            var sheets = new ISheet[tmpReader.ResultsCount];

            var container = new SheetContainer(fileUrl, sheets);

            for (var i = 0; i < tmpReader.ResultsCount; i++)
            {
                if (i == specificSheetIndex)
                {

                    tmpReader.Read();
                    CurrentSheet = new Sheet(tmpReader, container, i, true);

                }
                else
                {
                    tmpReader.NextResult();
                }

            }

            return CurrentSheet;
        }

        public async Task<ISheetContainer> Load(string fileUrl, ICollection<SheetMappingOptions> sheetsOptions)
        {
            url = fileUrl;

            tempLocalFilePath = await CopyFileLocal(fileUrl);
            await using var tempStream = File.Open(tempLocalFilePath, FileMode.Open, FileAccess.Read);
            
            using var tmpReader = ExcelReaderFactory.CreateReader(tempStream); 

            
            if(tmpReader.ResultsCount == 0) return null;

            var sheets = new ISheet[tmpReader.ResultsCount];

            var result = new SheetContainer(fileUrl, sheets);

            for (var i = 0; i < sheets.Length; i++)
            {
                var ignoreHeader = true;
                var options = sheetsOptions?.FirstOrDefault(o => o.SheetIndex == i) ?? SheetMappingOptions.Default(i);
                if(options.Map != null)
                    ignoreHeader = false;
                tmpReader.Read();
                var sheet = new Sheet(tmpReader, result, i,ignoreHeader);
                result.Sheets[i] = sheet;
                tmpReader.NextResult();
            }
            Container = result;
            CurrentSheet = Container.Sheets[0];
            
            return result;
            
        }

        public ISheet CurrentSheet { get; private set; }

        public IExcelRow Current { get; private set; }

        public ISheetContainer Container { get; private set; }


        public void Dispose()
        {
            stream?.Dispose();
            reader?.Dispose();
            if (File.Exists(tempLocalFilePath))
                File.Delete(tempLocalFilePath);
        }

        
        public async Task<bool> ReadRow()
        {
            if (!loaded)
            {
                tempLocalFilePath ??=  await CopyFileLocal(url);
                
                stream = File.Open(tempLocalFilePath, FileMode.Open, FileAccess.Read);
                
                reader = ExcelReaderFactory.CreateReader(stream);
                loaded = true;
            }

            if (specificSheetIndex >= 0)
                return await ReadRowSpecificSheet();
            
            return await ReadRowAllSheets();
            
        }
        private Task<bool> ReadRowSpecificSheet()
        {
            if (!inSpecificSheet)
            {
                if (specificSheetIndex != 0)
                {
                    var currentSheetReaderIndex = 0;
                    do
                    {
                        reader.NextResult();
                        currentSheetReaderIndex++;
                    } while (specificSheetIndex == currentSheetReaderIndex);
                }

                inSpecificSheet = true;
            }
            var found = reader.Read();
            
            if (!found)
                return Task.FromResult(false) ;
            index++;
            Current = new ExcelRow(index, CurrentSheet, reader);
            return Task.FromResult(true) ;
        }
        private async Task<bool> ReadRowAllSheets()
        {
            var found = reader.Read();
            
            if (!found)
            {
                if(reader.ResultsCount == CurrentSheet.Index + 1)
                    return false;
                CurrentSheet = Container.Sheets[CurrentSheet.Index + 1];
                reader.NextResult();
                index = 0;
                return await ReadRow();
            }
            index++;
            Current = new ExcelRow(index, CurrentSheet, reader);
            return true;
        }
        private async Task<string> CopyFileLocal(string cloudUrl)
        {
            await using var cloudFileStream = await cloudFilesService.OpenReadAsync(cloudUrl);
            var localPath = Path.GetTempFileName();
            await using var localStream = new FileStream(localPath,FileMode.Open,FileAccess.Write);
            await cloudFileStream.CopyToAsync(localStream);
            
            return localPath;
        }
        public void Reset()
        {
            specificSheetIndex = -1;
            inSpecificSheet = false;
            index = 0;
            loaded = false;
            stream?.Dispose();
            reader?.Dispose();
        }   
    }
}