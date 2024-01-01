using ClosedXML.Excel;
using task4.Models;

namespace task4.Services
{
	public class FileReadService
	{
        public List<Product> GetProductsFromExcel(string filePath)
        {
            var products = new List<Product>();
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                var rowNumber = 2;
                while (!worksheet.Cell(rowNumber, 1).Value.ToString().Equals(""))
                {
                    products.Add(new Product
                    {
                        Id = Convert.ToInt32(worksheet.Cell(rowNumber, 1).Value.ToString()),
                        Name = worksheet.Cell(rowNumber, 2).Value.ToString(),
                        Unit = worksheet.Cell(rowNumber, 3).Value.ToString(),
                        Price = Convert.ToDecimal(worksheet.Cell(rowNumber, 4).Value.ToString())
                    });
                    rowNumber++;
                }
            }
            return products;
        }

        public List<Client> GetClientsFromExcel(string filePath)
        {
            var clients = new List<Client>();
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(2);
                var rowNumber = 2;
                while (!worksheet.Cell(rowNumber, 1).Value.ToString().Equals(""))
                {
                    clients.Add(new Client
                    {
                        Id = Convert.ToInt32(worksheet.Cell(rowNumber, 1).Value.ToString()),
                        Name = worksheet.Cell(rowNumber, 2).Value.ToString(),
                        Address = worksheet.Cell(rowNumber, 3).Value.ToString(),
                        ContactPerson = worksheet.Cell(rowNumber, 4).Value.ToString()
                    });
                    rowNumber++;
                }
            }
            return clients;
        }

        public List<Order> GetOrdersFromExcel(string filePath)
        {
            var orders = new List<Order>();
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(3);
                var rowNumber = 2;
                while (!worksheet.Cell(rowNumber, 1).Value.ToString().Equals(""))
                {
                    orders.Add(new Order
                    {
                        Id = Convert.ToInt32(worksheet.Cell(rowNumber, 1).Value.ToString()),
                        ProductId = Convert.ToInt32(worksheet.Cell(rowNumber, 2).Value.ToString()),
                        ClientId = Convert.ToInt32(worksheet.Cell(rowNumber, 3).Value.ToString()),
                        Number = Convert.ToInt32(worksheet.Cell(rowNumber, 4).Value.ToString()),
                        Amount = Convert.ToInt32(worksheet.Cell(rowNumber, 5).Value.ToString()),
                        DateOfCreate = Convert.ToDateTime(worksheet.Cell(rowNumber, 6).Value.ToString())
                    });
                    rowNumber++;
                }
            }
            return orders;
        }
    }
}

