using System;
using System.IO;
using System.Text;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Reflection;
using Swashbuckle.Swagger;
using OfficeOpenXml;
 using OfficeOpenXml.Style;
using System.Drawing;

class Program
{

    public static async Task Main()
    {
        // Create a new instance of the KarateCreator class
        KarateCreator kc = new KarateCreator();
        using HttpClient client = new();
        client.DefaultRequestHeaders.Accept.Clear();
        client.DefaultRequestHeaders.Accept.Add(
        new MediaTypeWithQualityHeaderValue("application/vnd.github.v3+json"));
        client.DefaultRequestHeaders.Add("User-Agent", ".NET Foundation Repository Reporter");

        // Get the swagger JSON and generate the Karate feature file
        Console.WriteLine("Please enter URL for the swagger JSON");
        string URL = Console.ReadLine();
        Console.Clear();
        Console.WriteLine("Please enter authentication key");
        string authentication = Console.ReadLine();
        Console.Clear();
        Console.WriteLine("Please enter the URL for the API");
        string ApiHost = Console.ReadLine();
        Console.Clear();
        Console.WriteLine("Please enter the name of the API");
        string APIName = Console.ReadLine();
        //Modifies the API Name to replace any spaces with _
        APIName = APIName.Replace(" ", "_");
        Console.Clear();
        await kc.GetSwaggerAndGenerateKarate(client, URL, authentication, ApiHost , APIName);
    }
}




