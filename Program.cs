using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Elasticsearch.Net;
using OfficeOpenXml;

namespace elasticsearch_NEST_demo
{
    public class Program
    {
        static async Task Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var settings = new ConnectionConfiguration(new Uri("http://localhost:9201"))
                .RequestTimeout(TimeSpan.FromMinutes(2));

            var lowlevelClient = new ElasticLowLevelClient(settings);

             // await Expert(lowlevelClient, "CSP-S江苏.xlsx", "江苏省");
             // await Expert(lowlevelClient, "CSP-J江苏.xlsx", "江苏省");
             // await Expert(lowlevelClient, "CSP-J上海.xlsx", "上海市");
             // await Expert(lowlevelClient, "CSP-S上海.xlsx", "上海市");
             // await Expert(lowlevelClient, "CSP-J安徽.xlsx", "安徽省");
             // await Expert(lowlevelClient, "CSP-S安徽.xlsx", "安徽省");
             // await Expert(lowlevelClient, "CSP-J浙江.xlsx", "浙江省");
             // await Expert(lowlevelClient, "CSP-S浙江.xlsx", "浙江省");
             await Expert(lowlevelClient, "CSP-S广东.xlsx", "广东省");

            Console.WriteLine("Hello World!");
        }

        public static async Task Expert(ElasticLowLevelClient lowlevelClient, string fileName, string province)
        {
            var file = new FileInfo(fileName);
            using (var package = new ExcelPackage(file))
            {
                var worksheet = package.Workbook.Worksheets[1]; // Tip: To access the first worksheet, try index 1, not 0
                var rowCount = worksheet.Dimension.Rows;
                var colCount = worksheet.Dimension.Columns;
                int cout = 0;
                
                for (int row = 2; row <= rowCount; row++)
                {
                    var model = new Person();
                    for (int col = 1; col <= colCount; col++)
                    {
                        switch (col)
                        {
                            case 2:
                                model.Level = worksheet.Cells[row, col].Value?.ToString();
                                break;
                            case 3:
                                model.IdNumber = worksheet.Cells[row, col].Value?.ToString();
                                break;
                            case 4:
                                model.Gender = worksheet.Cells[row, col].Value?.ToString();
                                break;
                            case 5:
                                model.City = worksheet.Cells[row, col].Value?.ToString();
                                break;
                            case 6:
                                model.School = worksheet.Cells[row, col].Value?.ToString();
                                break;
                            case 7:
                            {
                                var str = worksheet.Cells[row, col].Value?.ToString();

                                model.Score = Convert.ToDouble(str != "缺考" ? str : "-11");
                                break;
                            }

                            case 8:
                                model.IsPassed = worksheet.Cells[row, col].Value?.ToString() == "是";
                                break;
                        }
                    }

                    model.Province = province;

                    await lowlevelClient.IndexAsync<BytesResponse>("csp", PostData.Serializable(model));
                    cout++;
                }
                Console.WriteLine($"处理了{province} {cout}条记录");
            }
        }
    }

    public class Person
    {
        public string Province { get; set; }
        public string Level { get; set; }
        public string IdNumber { get; set; }
        public string Gender { get; set; }
        public string City { get; set; }
        public string School { get; set; }
        public Double Score { get; set; }
        public bool IsPassed { get; set; }
    }
}