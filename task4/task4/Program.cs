using task4.Services;
using task4.Models;

namespace task4
{
    class Program
    {
        static void Main()
        {
            FileReadService fileReadService = new();
            FileUpdateService fileUpdateService = new();

            List<Product> products = new();
            List<Client> clients = new();
            List<Order> orders = new();

            while (true)
            {
                Console.WriteLine("Введите путь до файла с данными:");
                string filePath = Console.ReadLine();

                if (!String.IsNullOrEmpty(filePath))
                {
                    products = fileReadService.GetProductsFromExcel(filePath);
                    clients = fileReadService.GetClientsFromExcel(filePath);
                    orders = fileReadService.GetOrdersFromExcel(filePath);
                }
                else
                {
                    Console.WriteLine("Неверный путь!");
                    return;
                }

                while (true)
                {
                    Console.WriteLine("\nВведите номер действия, которое вы хотите выполнить:");
                    Console.WriteLine("1. Вывести информацию о клиенте по наименованию товара");
                    Console.WriteLine("2. Изменить данные контактного лица");
                    Console.WriteLine("3. Определить золотого клиента");
                    Console.WriteLine("4. Выход");

                    string action = Console.ReadLine();

                    switch (action)
                    {
                        case "1":
                            Console.Write("Введите наименование товара: ");
                            string productName = Console.ReadLine();
                            try
                            {
                                var product = products.FirstOrDefault(p => p.Name.Equals(productName, StringComparison.OrdinalIgnoreCase));
                                var listOfProductOrders = orders.Where(o => o.ProductId.Equals(product.Id)).ToList();

                                Console.WriteLine($"Наименование товара: {productName}");
                                foreach (var item in listOfProductOrders)
                                {
                                    var client = clients.FirstOrDefault(c => c.Id.Equals(item.ClientId));
                                    Console.WriteLine($"Наименование: {client.Name} \n " +
                                        $"Адрес: {client.Address} \n " +
                                        $"Контактное лицо: {client.ContactPerson} \n " +
                                        $"Требуемое количество: {item.Amount} \n " +
                                        $"Цена за единицу: {product.Price} \n " +
                                        $"Дата размещения: {item.DateOfCreate}");
                                }
                                break;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                                return;
                            }
                        case "2":
                            Console.Write("Название организации:");
                            string organizationName = Console.ReadLine();
                            Console.Write("ФИО нового контактного лица:");
                            string newContactName = Console.ReadLine();
                            var resultOfUpdating = fileUpdateService.UpdateExcel(filePath, organizationName, newContactName);
                            if (resultOfUpdating)
                            {
                                Console.WriteLine("Изменения сохранены!");
                                break;
                            }
                            else
                            {
                                Console.WriteLine("Не удалось сохранить изменения! Попробуйте заново.");
                                return;
                            }
                        case "3":
                            Console.Write("Введите месяц (1-12): ");
                            int month = Convert.ToInt32(Console.ReadLine());
                            Console.Write("Введите год: ");
                            int year = Convert.ToInt32(Console.ReadLine());
                            var ordersForThePeriod = orders.Where(o => o.DateOfCreate.Year == year && o.DateOfCreate.Month == month);
                            Dictionary<int, int> clientIdCounts = ordersForThePeriod.GroupBy(c => c.ClientId)
                                              .ToDictionary(g => g.Key, g => g.Count());
                            int maxCount = clientIdCounts.Values.Max();
                            var keysWithMaxCount = clientIdCounts.Keys.Where(key => clientIdCounts[key] == maxCount).ToList();
                            if (keysWithMaxCount.Count > 1)
                            {
                                Console.WriteLine($"Золотые клиенты с наибольшими заказами за {month}/{year}:");
                                keysWithMaxCount.ForEach(key => Console.WriteLine(clients.FirstOrDefault(c => c.Id.Equals(key)).Name));
                            }
                            else
                            {
                                Console.WriteLine($"Золотой клиент с наибольшими заказами за {month}/{year}:");
                                Console.WriteLine(clients.FirstOrDefault(c => c.Id.Equals(keysWithMaxCount[0])).Name);
                            }
                            break;
                        case "4":
                            return;
                        default:
                            Console.WriteLine("Неверный номер действия!");
                            break;
                    }
                }
            }
        }
    }
}