using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestClient.ServiceReference1;

namespace TestClient
{
    class Program
    {
        public static void Main(string[] args)
        {
            Service1Client serviceClient = new Service1Client();

            serviceClient.ParseEmployeeFile();
        }
    }
}
