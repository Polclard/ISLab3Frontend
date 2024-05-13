using AdminApplication.Models;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System.Text;

namespace AdminApplication.Controllers
{
    public class OrderController : Controller
    {
        public async Task<IActionResult> Index()
        {
            HttpClient client = new HttpClient();
            string URL = "http://localhost:5054/api/Admin/GetAllOrders";
            var response = await client.GetAsync(URL);
            var data = response.Content.ReadAsAsync<List<Order>>().Result;

            return View(data);
        }

        public IActionResult Details(Guid Id)
        {
            HttpClient client = new HttpClient();
            string URL = "http://localhost:5054/api/Admin/GetDetailsForOrder";
            var model = new
            {
                Id = Id
            };
            HttpContent content = new StringContent(JsonConvert.SerializeObject(model), Encoding.UTF8, "application/json");

            HttpResponseMessage response = client.PostAsync(URL, content).Result;

            var data = response.Content.ReadAsAsync<Order>().Result;



            HttpClient client1 = new HttpClient();
            string URL1 = "http://localhost:5054/api/Admin/GetDetailsForOrderProducts";
            var model1 = new
            {
                Id = Id
            };
            HttpContent content1 = new StringContent(JsonConvert.SerializeObject(model1), Encoding.UTF8, "application/json");

            HttpResponseMessage response1 = client1.PostAsync(URL1, content1).Result;
            var data1 = response1.Content.ReadAsAsync<List<TicketInOrder>>().Result;


            data.ProductInOrders = data1;





            for (int i = 0; i < data1.Count; i++)
            {
                HttpClient client2 = new HttpClient();
                string URL2 = "http://localhost:5054/api/Admin/GetDetailsTicket";
                var model2 = new
                {
                    Id = data1[i].TicketId
                };
                HttpContent content2 = new StringContent(JsonConvert.SerializeObject(model2), Encoding.UTF8, "application/json");

                HttpResponseMessage response2 = client2.PostAsync(URL2, content2).Result;
                var data2 = response2.Content.ReadAsAsync<Ticket>().Result;
                data1.ElementAt(i).OrderedProduct = data2;
            }





            for (int i = 0; i < data1.Count; i++)
            {
                HttpClient client3 = new HttpClient();
                string URL3 = "http://localhost:5054/api/Admin/GetDetailsConcert";
                var model3 = new
                {
                    Id = data1[i].OrderedProduct.ConcertId
                };
                HttpContent content3 = new StringContent(JsonConvert.SerializeObject(model3), Encoding.UTF8, "application/json");

                HttpResponseMessage response3 = client3.PostAsync(URL3, content3).Result;
                var data3 = response3.Content.ReadAsAsync<Concert>().Result;
                data1.ElementAt(i).OrderedProduct.Concert = data3;
            }

            data.ProductInOrders = data1;

            return View(data);
        }



        public Order GetOrderDetailsInternaly(Guid Id)
        {
            HttpClient client = new HttpClient();
            string URL = "http://localhost:5054/api/Admin/GetDetailsForOrder";
            var model = new
            {
                Id = Id
            };
            HttpContent content = new StringContent(JsonConvert.SerializeObject(model), Encoding.UTF8, "application/json");

            HttpResponseMessage response = client.PostAsync(URL, content).Result;

            var data = response.Content.ReadAsAsync<Order>().Result;



            HttpClient client1 = new HttpClient();
            string URL1 = "http://localhost:5054/api/Admin/GetDetailsForOrderProducts";
            var model1 = new
            {
                Id = Id
            };
            HttpContent content1 = new StringContent(JsonConvert.SerializeObject(model1), Encoding.UTF8, "application/json");

            HttpResponseMessage response1 = client1.PostAsync(URL1, content1).Result;
            var data1 = response1.Content.ReadAsAsync<List<TicketInOrder>>().Result;


            data.ProductInOrders = data1;





            for (int i = 0; i < data1.Count; i++)
            {
                HttpClient client2 = new HttpClient();
                string URL2 = "http://localhost:5054/api/Admin/GetDetailsTicket";
                var model2 = new
                {
                    Id = data1[i].TicketId
                };
                HttpContent content2 = new StringContent(JsonConvert.SerializeObject(model2), Encoding.UTF8, "application/json");

                HttpResponseMessage response2 = client2.PostAsync(URL2, content2).Result;
                var data2 = response2.Content.ReadAsAsync<Ticket>().Result;
                data1.ElementAt(i).OrderedProduct = data2;
            }





            for (int i = 0; i < data1.Count; i++)
            {
                HttpClient client3 = new HttpClient();
                string URL3 = "http://localhost:5054/api/Admin/GetDetailsConcert";
                var model3 = new
                {
                    Id = data1[i].OrderedProduct.ConcertId
                };
                HttpContent content3 = new StringContent(JsonConvert.SerializeObject(model3), Encoding.UTF8, "application/json");

                HttpResponseMessage response3 = client3.PostAsync(URL3, content3).Result;
                var data3 = response3.Content.ReadAsAsync<Concert>().Result;
                data1.ElementAt(i).OrderedProduct.Concert = data3;
            }

            data.ProductInOrders = data1;
            return data;
        }


        [HttpGet]
        public FileContentResult ExportAllOrders()
        {
            string fileName = "Orders.xlsx";
            string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            using (var workbook = new XLWorkbook())
            {
                IXLWorksheet worksheet = workbook.Worksheets.Add("Orders");
                worksheet.Cell(1, 1).Value = "OrderID";
                worksheet.Cell(1, 2).Value = "Customer UserName";
                worksheet.Cell(1, 3).Value = "Total Price";
                HttpClient client = new HttpClient();
                string URL = "http://localhost:5054/api/Admin/GetAllOrders";

                HttpResponseMessage response = client.GetAsync(URL).Result;
                var data = response.Content.ReadAsAsync<List<Order>>().Result;

                for (int i = 0; i < data.Count(); i++)
                {
                    var item = GetOrderDetailsInternaly(data[i].Id);
                    worksheet.Cell(i + 2, 1).Value = item.Id.ToString();
                    worksheet.Cell(i + 2, 2).Value = item.Owner.FirstName + " " + item.Owner.LastName;
                    var total = 0;
                    for (int j = 0; j < item.ProductInOrders.Count(); j++)
                    {
                        worksheet.Cell(1, 4 + j).Value = "Ticket - " + (j + 1);
                        worksheet.Cell(i + 2, 4 + j).Value = item.ProductInOrders.ElementAt(j).OrderedProduct.Concert.ConcertName;
                        total += (int)((double)item.ProductInOrders.ElementAt(j).Quantity * (double)item.ProductInOrders.ElementAt(j).OrderedProduct.Price);
                    }
                    worksheet.Cell(i + 2, 3).Value = total;
                }
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(content, contentType, fileName);
                }
            }

        }
    }
}
