# SimplyWorks ExcelImport

[![Build and Publish NuGet Package](https://github.com/simplify9/SW-ExcelImport/actions/workflows/nuget-publish.yml/badge.svg)](https://github.com/simplify9/SW-ExcelImport/actions/workflows/nuget-publish.yml)
[![NuGet](https://img.shields.io/nuget/v/SimplyWorks.ExcelImport.svg)](https://www.nuget.org/packages/SimplyWorks.ExcelImport/)
[![NuGet Extensions](https://img.shields.io/nuget/v/SimplyWorks.ExcelImport.Extensions.svg)](https://www.nuget.org/packages/SimplyWorks.ExcelImport.Extensions/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A .NET library for reading, parsing, and validating Excel files with type-safe operations and comprehensive error handling.

## Features

- **Excel File Reading**: Read Excel files (.xlsx, .xls) using ExcelDataReader
- **Type-Safe Parsing**: Parse Excel data into strongly-typed C# objects
- **Data Validation**: Built-in validation using data annotations
- **Batch Processing**: Support for batch imports with error tracking
- **Sheet Mapping**: Map Excel sheets to different types with flexible column mapping
- **Query Interface**: Query imported Excel data with filtering and pagination
- **Entity Framework Integration**: Store and query Excel data using Entity Framework

## Packages

This library consists of two NuGet packages:

- **SimplyWorks.ExcelImport**: Core library with Excel reading and parsing functionality
- **SimplyWorks.ExcelImport.Extensions**: ASP.NET Core dependency injection extensions

## Installation

Install the core package:
```bash
dotnet add package SimplyWorks.ExcelImport
```

For ASP.NET Core applications, also install the extensions package:
```bash
dotnet add package SimplyWorks.ExcelImport.Extensions
```

## Quick Start

### 1. Define Your Data Model

```csharp
public class Order
{
    [Required]
    public string OrderId { get; set; }
    
    [Required]
    public string CustomerName { get; set; }
    
    [Range(0, double.MaxValue)]
    public decimal Amount { get; set; }
    
    public DateTime OrderDate { get; set; }
}
```

### 2. Configure Services (ASP.NET Core)

```csharp
public void ConfigureServices(IServiceCollection services)
{
    services.AddExcelImport();
    // Configure your DbContext
    services.AddDbContext<YourDbContext>(options => 
        options.UseSqlServer(connectionString));
}
```

### 3. Import Excel Data

```csharp
public class OrderImportService
{
    private readonly ExcelService _excelService;
    
    public OrderImportService(ExcelService excelService)
    {
        _excelService = excelService;
    }
    
    public async Task ImportOrders(string excelFileUrl)
    {
        var options = new TypedParseToJsonOptions 
        {
            TypeAssemblyQualifiedName = typeof(Order).AssemblyQualifiedName,
            NamingStrategy = JsonNamingStrategy.SnakeCase
        };
        
        // Load and validate Excel file
        var container = await _excelService.LoadExcelFileInfo(excelFileUrl, options);
        
        // Check for validation errors
        if (container.Sheets.Any(sheet => sheet.HasErrors()))
        {
            // Handle validation errors
            return;
        }
        
        // Import valid data
        await _excelService.Import(excelFileUrl, options);
    }
}
```

### 4. Query Imported Data

```csharp
public class OrderQueryService
{
    private readonly IExcelQueryable _excelQueryable;
    
    public OrderQueryService(IExcelQueryable excelQueryable)
    {
        _excelQueryable = excelQueryable;
    }
    
    public async Task<IEnumerable<Order>> GetValidOrders(string reference, int pageIndex = 0, int pageSize = 100)
    {
        var options = new ExcelQueryValidatedOptions
        {
            Reference = reference,
            PageIndex = pageIndex,
            PageSize = pageSize,
            RowStatus = QueryRowStatus.Valid
        };
        
        return await _excelQueryable.Get<Order>(options);
    }
}
```

## Core Components

### ExcelService
Main service for importing Excel files with validation and error handling.

### IExcelReader
Interface for reading Excel files and sheets with support for:
- Multiple sheet processing
- Column mapping
- Row-by-row reading

### ExcelRepo
Repository pattern implementation for storing Excel import metadata and results.

### SheetReader<T>
Generic sheet reader for type-safe Excel data parsing.

### Data Validation
Built-in support for:
- Data annotation validation
- Type conversion validation
- Custom validation rules

## Configuration Options

### TypedParseToJsonOptions
- `TypeAssemblyQualifiedName`: Target type for parsing
- `NamingStrategy`: JSON naming strategy (SnakeCase, CamelCase, etc.)
- `SheetsOptions`: Configuration for multiple sheets

### SheetMappingOptions
- Column mapping configuration
- Sheet selection options
- Validation rules

## Entity Framework Integration

The library includes Entity Framework entities for tracking:
- `ExcelFileRecord`: Excel file metadata
- `SheetRecord`: Individual sheet information
- `RowRecord`: Row-level data and validation results
- `CellRecord`: Cell-level data
- `Batch` and `BatchItem`: Batch processing tracking

## Dependencies

- **ExcelDataReader**: For reading Excel files
- **Newtonsoft.Json**: JSON serialization
- **Entity Framework Core**: Data persistence
- **SimplyWorks Libraries**: Additional utilities and extensions

## Target Framework

- **.NET Standard 2.1** (Core library)
- **.NET Core 3.1** (Extensions package)

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

Please read [CONTRIBUTING.md](.github/CONTRIBUTING.md) for details on our code of conduct and the process for submitting pull requests.

## Support

For issues and questions:
- Create an issue on [GitHub](https://github.com/simplify9/SW-ExcelImport/issues)
- Check existing [documentation](https://github.com/simplify9/SW-ExcelImport)

## Related Packages

This library is part of the SimplyWorks ecosystem:
- [SimplyWorks.EfCoreExtensions](https://www.nuget.org/packages/SimplyWorks.EfCoreExtensions/)
- [SimplyWorks.ExportToExcel](https://www.nuget.org/packages/SimplyWorks.ExportToExcel/)
- [SimplyWorks.HttpExtensions](https://www.nuget.org/packages/SimplyWorks.HttpExtensions/)
- [SimplyWorks.PrimitiveTypes](https://www.nuget.org/packages/SimplyWorks.PrimitiveTypes/)