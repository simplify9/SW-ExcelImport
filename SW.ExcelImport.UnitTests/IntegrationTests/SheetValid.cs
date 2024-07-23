// using System.Linq;
// using System.Threading.Tasks;
// using Microsoft.VisualStudio.TestTools.UnitTesting;
// using SW.ExcelImport.Services;
// using SW.ExcelImport.UnitTests.TypeBased;
//
// namespace SW.ExcelImport.UnitTests.Integration
// {
//
//     [TestClass]
//     public class SheetValid
//     {
//         
//         private static SheetReader<Order> svc;
//         
//         
//         [ClassInitialize]
//         public static void ClassInitialize(TestContext tContext)
//         {
//             svc=(SheetReader<Order>)TestServerClientFactory.GetServiceProvider().GetService(typeof(SheetReader<Order>));
//         }
//         
//         [TestMethod]
//         public async Task Basic()
//         {
//             var loadResult = await svc.Validate("valid_single_sheet");
//             Assert.AreEqual(false, loadResult.HasErrors);
//
//             var result = await svc.ReadAll("valid_single_sheet");
//             
//             //Assert.IsTrue(result.Count == 4);
//             
//         }
//
//     }
// }