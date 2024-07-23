using System.Collections.Generic;
using System.Data;
namespace SW.ExcelImport
{
    public class SheetMappingOptions
    {
        public static SheetMappingOptions Default(int index)
        {
            return new SheetMappingOptions()
            {
                Map = null,
                IdIndex = index == 0 ? 0 : (int?)null,
                ParentIdIndex = index == 0 ? (int?)null :0,
                IndexAsId = index != 0
            };
        }
        
        public int SheetIndex { get; set; }
        public string[] Map { get; set; }
        public int? IdIndex { get; set; }
        public int? ParentIdIndex { get; set; }
        public string SheetName { get; set; }
        public bool IndexAsId { get; set; }
        public int? ParenSheetIndex { get; set; }
        public bool Ignore { get; set; }
        public IDictionary<string,string> CustomMap  { get; set; }
    }
}
