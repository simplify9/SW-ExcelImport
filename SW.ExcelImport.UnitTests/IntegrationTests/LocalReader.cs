using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using SW.PrimitiveTypes;
using SW.ExcelImport;

namespace SW.ExcelImport.UnitTests
{
    public class LocalReader : ICloudFilesService
    {
        public Task<IReadOnlyDictionary<string, string>> GetMetadataAsync(string key)
        {
            throw new NotImplementedException();
        }

        public Task<bool> DeleteAsync(string key)
        {
            throw new NotImplementedException();
        }

        public string GetSignedUrl(string key, TimeSpan expiry)
        {
            throw new NotImplementedException();
        }

        public string GetUrl(string key)
        {
            throw new NotImplementedException();
        }

        public Task<IEnumerable<CloudFileInfo>> ListAsync(string prefix)
        {
            throw new NotImplementedException();
        }

        public Task<Stream> OpenReadAsync(string key)
        {
            var result = File.OpenRead ($@"IntegrationTests/Data/{key}.xlsx");
            return Task.FromResult(result as Stream) ;
        }

        public WriteWrapper OpenWrite(WriteFileSettings settings)
        {
            throw new NotImplementedException();
        }

        public Task<RemoteBlob> WriteAsync(Stream inputStream, WriteFileSettings settings)
        {
            throw new NotImplementedException();
        }

        public Task<RemoteBlob> WriteTextAsync(string text, WriteFileSettings settings)
        {
            throw new NotImplementedException();
        }
    }
}