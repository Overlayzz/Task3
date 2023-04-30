using System;
using ClosedXML.Excel;

namespace task1
{
    class Program
    {
        static void Main()
        {
            string path;
            while (true)
            {
                Console.WriteLine("Введите путь к файлу:");
                path = Console.ReadLine();
                if (File.Exists(path))
                {
                    break;
                }
                Console.WriteLine("Некорректный путь к файлу. Пожалуйста, повторите ввод.");
            }            
            var parser = new DataParser(path);
            while (true)
            {
                Console.WriteLine("\nВыберите действие:\n");
                Console.WriteLine("1 - Вывести все данные");
                Console.WriteLine("2 - Вывести информацию о заказах на товар");
                Console.WriteLine("3 - Найти золотого клиента");
                Console.WriteLine("4 - Обновить контактную информацию клиента");
                Console.WriteLine("5 - Выход");

                string choice = Console.ReadLine();

                switch (choice)
                {
                    // выводим все данные
                    case "1":
                        DataPrinter dataPrinter = new DataPrinter(parser.GetClientData(), parser.GetProductData(), parser.GetRequestData());
                        dataPrinter.PrintAllData();
                        break;
                    // выводим информацию о заказах на заданный товар
                    case "2":
                        Console.WriteLine("Введите название товара:");
                        string productName = Console.ReadLine();
                        List<ClientInfo> clients = parser.GetClientData();
                        List<ProductInfo> products = parser.GetProductData();
                        List<RequestsInfo> requests = parser.GetRequestData();
                        DataPrinter datPrinter = new DataPrinter(clients, products, requests);
                        datPrinter.PrintRequestsForProduct(productName, clients, products, requests);
                        break;
                    // находим золотого клиента за заданный месяц и год
                    case "3":
                        var years = new int[] { 2021, 2022, 2023, 2024, 2025 };
                        var months = new string[] { "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь" };

                        // выводим список годов и просим пользователя выбрать
                        Console.WriteLine("Выберите год:");
                        for (int i = 0; i < years.Length; i++)
                        {
                            Console.WriteLine($"{i + 1}: {years[i]}");
                        }
                        int yearIndex = int.Parse(Console.ReadLine()) - 1;
                        int year = years[yearIndex];

                        // выводим список месяцев и просим пользователя выбрать
                        Console.WriteLine("Выберите месяц:");
                        for (int i = 0; i < months.Length; i++)
                        {
                            Console.WriteLine($"{i + 1}: {months[i]}");
                        }
                        int monthIndex = int.Parse(Console.ReadLine()) - 1;
                        int month = monthIndex + 1;
                        string goldenClientId = parser.GetGoldenClient(year, month);
                        Console.WriteLine($"Код 'золотого' клиента: {goldenClientId}");
                        break;
                    // обновляем контактную информацию для заданного клиента
                    case "4":
                        Console.WriteLine("Введите название компании клиента:");
                        string companyName = Console.ReadLine();
                        Console.WriteLine("Введите новый контакт клиента:\n");
                        string newContact = Console.ReadLine();
                        parser.UpdateClientContact(companyName, newContact);
                        break;
                    case "5":
                        return;
                    default:
                        Console.WriteLine("Некорректный ввод.");
                        break;
                }
                Console.ReadLine();
            }
        }
    }
    // класс DataParser включает методы для анализа данных из файла Excel и обновления контактной информации клиента
    public class DataParser
    {
        private readonly string _filePath;
        // конструктор, который инициализирует "_filePath"
        public DataParser(string filePath)
        {
            _filePath = filePath;
        }
        // возвращает список объектов ProductInfo, проанализированных на листе "Товары"
        public List<ProductInfo> GetProductData()
        {
            var productList = new List<ProductInfo>();
            using (var workbook = new XLWorkbook(_filePath))
            {
                var worksheet = workbook.Worksheet("Товары");
                var range = worksheet.RangeUsed();
                foreach (var row in range.RowsUsed().Skip(1))
                {
                    var productInfo = new ProductInfo();
                    productInfo.ID_Product = row.Cell(1).GetValue<string>();
                    productInfo.Name = row.Cell(2).GetValue<string>();
                    productInfo.Unit = row.Cell(3).GetValue<string>();
                    productInfo.UnitPrice = row.Cell(4).GetValue<string>();
                    productList.Add(productInfo);
                }
            }
            return productList;
        }
        //возвращает список объектов ClientInfo, проанализированных на листе "Клиенты"
        public List<ClientInfo> GetClientData()
        {
            var clientList = new List<ClientInfo>();
            using (var workbook = new XLWorkbook(_filePath))
            {
                var worksheet = workbook.Worksheet("Клиенты");
                var range = worksheet.RangeUsed();
                foreach (var row in range.RowsUsed().Skip(1))
                {
                    var clientInfo = new ClientInfo();
                    clientInfo.ID_Client = row.Cell(1).GetValue<string>();
                    clientInfo.Name_Company = row.Cell(2).GetValue<string>();
                    clientInfo.Address_Company = row.Cell(3).GetValue<string>();
                    clientInfo.Contact_Company = row.Cell(4).GetValue<string>();
                    clientList.Add(clientInfo);
                }
            }
            return clientList;
        }
        // возвращает список объектов RequestsInfo, проанализированных на листе "Заявки"
        public List<RequestsInfo> GetRequestData()
        {
            var requestList = new List<RequestsInfo>();
            using (var workbook = new XLWorkbook(_filePath))
            {
                var worksheet = workbook.Worksheet("Заявки");
                var range = worksheet.RangeUsed();
                foreach (var row in range.RowsUsed().Skip(1))
                {
                    var requestInfo = new RequestsInfo();
                    requestInfo.ID_Query = row.Cell(1).GetValue<string>();
                    requestInfo.ID_Product = row.Cell(2).GetValue<string>();
                    requestInfo.ID_Client = row.Cell(3).GetValue<string>();
                    requestInfo.Num_Query = row.Cell(4).GetValue<string>();
                    requestInfo.Required_Quantity = row.Cell(5).GetValue<string>();
                    requestInfo.Date_Publish = row.Cell(6).GetValue<string>();
                    requestList.Add(requestInfo);
                }
            }
            return requestList;
        }
        // обновляет контактную информацию клиента в файле Excel, учитывая название компании и новую контактную информацию
        public void UpdateClientContact(string companyName, string newContact)
        {
            using (var workbook = new XLWorkbook(_filePath))
            {
                var worksheet = workbook.Worksheet("Клиенты");
                var range = worksheet.RangeUsed();
                foreach (var row in range.RowsUsed().Skip(1))
                {
                    var companyNameFromSheet = row.Cell(2).GetValue<string>();
                    if (companyNameFromSheet.Equals(companyName))
                    {
                        row.Cell(4).Value = newContact;
                    }
                }
                workbook.Save();
            }
            Console.WriteLine($"Контактное лицо для организации {companyName} изменено на {newContact}.");
        }
        // возвращает ид клиента, у которого было больше всего запросов в данном месяце и году
        public string GetGoldenClient(int year, int month)
        {
            string goldenClientId = "";
            int maxRequests = 0;
            using (var workbook = new XLWorkbook(_filePath))
            {
                var requestWorksheet = workbook.Worksheet("Заявки");
                var clientWorksheet = workbook.Worksheet("Клиенты");
                var requestRange = requestWorksheet.RangeUsed();
                var clientRange = clientWorksheet.RangeUsed();


                foreach (var clientRow in clientRange.RowsUsed().Skip(1))
                {
                    var clientId = clientRow.Cell(1).GetValue<string>();
                    var clientRequests = requestRange.RowsUsed().Skip(1)
                        .Where(row => row.Cell(3).GetValue<string>() == clientId
                            && row.Cell(6).GetValue<DateTime>().Year == year
                            && row.Cell(6).GetValue<DateTime>().Month == month);

                    var numRequests = clientRequests.Count();

                    if (numRequests > maxRequests)
                    {
                        maxRequests = numRequests;
                        goldenClientId = clientId;
                    }
                }
            }
            return goldenClientId;
        }

    }

    public class ProductInfo
    {
        public string ID_Product { get; set; }
        public string Name { get; set; }
        public string Unit { get; set; }
        public string UnitPrice { get; set; }
    }

    public class ClientInfo
    {
        public string ID_Client { get; set; }
        public string Name_Company { get; set; }
        public string Address_Company { get; set; }
        public string Contact_Company { get; set; }
    }

    public class RequestsInfo
    {
        public string ID_Query { get; set; }
        public string ID_Product { get; set; }
        public string ID_Client { get; set; }
        public string Num_Query { get; set; }
        public string Required_Quantity { get; set; }
        public string Date_Publish { get; set; }
    }

    public class DataPrinter
    {
        private List<ClientInfo> clients;
        private List<ProductInfo> products;
        private List<RequestsInfo> requests;


        public DataPrinter(List<ClientInfo> clients, List<ProductInfo> products, List<RequestsInfo> requests)
        {
            this.clients = clients;
            this.products = products;
            this.requests = requests;
        }
        // перебирает объекты в каждом из списков и выводит их
        public void PrintAllData()
        {
            Console.WriteLine("Clients:");
            foreach (var client in clients)
            {
                Console.WriteLine($"ID_Client: {client.ID_Client} | Name_Company: {client.Name_Company} | Address: {client.Address_Company} | Contact_Company: {client.Contact_Company}");
            }

            Console.WriteLine("Products:");
            foreach (var product in products)
            {
                Console.WriteLine($"ID_Product: {product.ID_Product} | Name: {product.Name} | Unit: {product.Unit} | Unit_price: {product.UnitPrice}");
            }

            Console.WriteLine("Requests:");
            foreach (var request in requests)
            {
                Console.WriteLine($"ID_Query: {request.ID_Query} | ID_Product: {request.ID_Product} | ID_Client: {request.ID_Client} | Num_Query: {request.Num_Query} | Required_Quantity: {request.Required_Quantity} | Date_Publish: {request.Date_Publish}");
            }
        }
        // ищет объект ProductInfo в списке продуктов с совпадающим именем
        public void PrintRequestsForProduct(string productName, List<ClientInfo> clients, List<ProductInfo> products, List<RequestsInfo> requests)
        {
            var product = products.FirstOrDefault(p => p.Name == productName);
            if (product == null)
            {
                Console.WriteLine($"Товар {productName} не найден.");
                return;
            }


            var relevantRequests = requests.Where(o => o.ID_Product == product.ID_Product);
            var requestsInfo = new List<(string clientName, string quantity, string price, string date)>();
            foreach (var request in relevantRequests)
            {
                var client = clients.FirstOrDefault(c => c.ID_Client == request.ID_Client);
                if (client != null)
                {
                    requestsInfo.Add((client.Name_Company, request.Required_Quantity, product.UnitPrice, request.Date_Publish));
                }
            }

            if (requestsInfo.Count == 0)
            {
                Console.WriteLine($"Товар {productName} не заказывали.");
                return;
            }

            Console.WriteLine($"Информация о заказах на товар {productName}:");
            foreach (var (clientName, quantity, price, date) in requestsInfo)
            {
                Console.WriteLine($"Клиент: {clientName}, количество: {quantity}, цена: {price}, дата заказа: {date}");
            }
        }
    }
}
