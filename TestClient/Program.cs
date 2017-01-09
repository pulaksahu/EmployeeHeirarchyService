using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestClient.ServiceReference1;
using System.Net.Http;

namespace TestClient
{
    public class DataObject
    {
        public string Name { get; set; }
    }
    class Program
    {
        public static void Main(string[] args)
        {
            //Service1Client serviceClient = new Service1Client();

            //string a = serviceClient.ParseEmployeeFile();


            HttpClient client = new HttpClient();

            client.BaseAddress = new Uri("http://localhost:12618/");
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

            HttpResponseMessage response = client.GetAsync("api/employeeheirarcy").Result;  // Blocking call!
            if (response.IsSuccessStatusCode)
            {
                // Parse the response body. Blocking!
                var dataObjects = response.Content.ToString();
                //foreach (var d in dataObjects)
                //{
                //    Console.WriteLine("{0}", d.Name);
                //}
            }
            else
            {
                Console.WriteLine("{0} ({1})", (int)response.StatusCode, response.ReasonPhrase);
            }  

            Console.ReadLine();

        }
    }
}
