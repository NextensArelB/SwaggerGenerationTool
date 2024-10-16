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

/// <summary>
/// The KarateCreator class is responsible for generating Karate test scripts and configurations
/// based on Swagger API documentation. It creates necessary directories, configuration files,
/// and test scripts for different HTTP methods (POST, GET, PUT, PATCH, DELETE).
/// </summary>


using System.Text.RegularExpressions;

public class KarateCreator {

    //The function that is used to add text to a file in this case generating all the karate files
    private void AddText(FileStream fs, string value)
    {
        byte[] info = new UTF8Encoding(true).GetBytes(value);
        fs.Write(info, 0, info.Length);
    }

    //The function that is used to get the swagger JSON and generate the karate tests
    public async Task GetSwaggerAndGenerateKarate(HttpClient client, string URL, string authentication, string ApiHost, string APIName)
    {
        try
        {
            //Get the JSON from the swagger URL
            var json = await client.GetStringAsync(URL);

            var settings = new JsonSerializerSettings();
            settings.MetadataPropertyHandling = MetadataPropertyHandling.Ignore;
            SwaggerDocument JsonElements = JsonConvert.DeserializeObject<SwaggerDocument>(json, settings);
            GenerateKarateSetup(APIName);

            //Check if the swagger is version 2.0 or 3.0
            if (JsonElements.swagger == "2.0")
            {
                GenerateSwagger2(JsonElements, authentication, APIName);
            }
            else
            {
                GenerateSwagger3(JsonElements, authentication, ApiHost, APIName);
            }
        }
        catch (Exception ex)
        {
            // Handle the exception here
            Console.WriteLine("The URL: " + URL +" is incorrect.\n Please verify that the URL is of the JSON content of the swagger, not the Swagger itself. \nCheck the markdown guide for more information on how to obtain the swagger.json ");

        }
    }

    //The function that is used to generate the karate setup files
    private void GenerateKarateSetup(string APIName)
    {
        //Create the necessary directories
        string FolderPath = @"c:\repos\AutomatedTests\" + APIName + "\\Configuration";
        Directory.CreateDirectory(FolderPath);
        FolderPath = @"c:\repos\AutomatedTests\" + APIName + "\\Configuration\\Secrets";
        Directory.CreateDirectory(FolderPath);
        FolderPath = @"c:\repos\AutomatedTests\" + APIName + "\\Configuration\\Secrets.Local";
        Directory.CreateDirectory(FolderPath);
        FolderPath = @"c:\repos\AutomatedTests\" + APIName + "\\Features";
        Directory.CreateDirectory(FolderPath);
        StringBuilder karateConfig = new StringBuilder();
        StringBuilder AzureSecrets = new StringBuilder();
        StringBuilder testJs = new StringBuilder();

        //Paths for the files that will be generated
        string pathConfig = @"c:\repos\AutomatedTests\" + APIName + "\\karate-config.js";
        string pathAzureSec = @"c:\repos\AutomatedTests\" + APIName + "\\Configuration\\Secrets\\azure.json";
        string pathTestJs = @"c:\repos\AutomatedTests\" + APIName + "\\test.js";
        string pathPackageJson = @"c:\repos\AutomatedTests\" + APIName + "\\package.json";
        string pathDevDef = @"c:\repos\AutomatedTests\" + APIName + "\\Configuration\\dev.json";
        string pathAccDef = @"c:\repos\AutomatedTests\" + APIName + "\\Configuration\\acc.json";
        string pathProdDef = @"c:\repos\AutomatedTests\" + APIName + "\\Configuration\\prod.json";
        string pathDevSec = @"c:\repos\AutomatedTests\" + APIName + "\\Configuration\\Secrets.Local\\dev.json";
        string pathAccSec = @"c:\repos\AutomatedTests\" + APIName + "\\Configuration\\Secrets.Local\\acc.json";
        string pathProdSec = @"c:\repos\AutomatedTests\" + APIName + "\\Configuration\\Secrets.Local\\prod.json";

        //Generate the shared files
        generateSharedFiles(pathDevDef);
        generateSharedFiles(pathAccDef);
        generateSharedFiles(pathProdDef);
        generateSharedFiles(pathDevSec);
        generateSharedFiles(pathAccSec);
        generateSharedFiles(pathProdSec);
        generateSharedFiles(pathAzureSec);

        //Generate the karate-config.js file
        karateConfig.AppendLine("function fn() {");
        karateConfig.AppendLine("\t var env = karate.env  // Set env to karate.env if it is set");
        karateConfig.AppendLine("\n\t karate.log('env: ', env); // Below we tell the config to take the configuration file set to the chosen environment");
        karateConfig.AppendLine("\n\t var definitions = karate.read('file:Configuration/' + env.toLowerCase() + '.json');");
        karateConfig.AppendLine("\t var secretFileName = isAzure ? 'file:Configuration/Secrets/Azure' : 'file:Configuration/Secrets.Local/' + env.toLowerCase();");
        karateConfig.AppendLine("\t karate.log('secretFileName: ', secretFileName);");
        karateConfig.AppendLine("\n\t var secrets = karate.read(secretFileName + '.json');");

        karateConfig.AppendLine("\n\t var config = {");
        karateConfig.AppendLine("\t\t // Below we tell the config to take the configuration file set to the chosen environment");
        karateConfig.AppendLine("\t\t definitions: definitions,");
        karateConfig.AppendLine("\t\t secrets: secrets");
        karateConfig.AppendLine("\t\t // Variables");
        karateConfig.AppendLine("\t\t // Below we can set variables that can be used in the tests");
        karateConfig.AppendLine("\t\t };");

        karateConfig.AppendLine("\t karate.configure('ssl', true);");
        karateConfig.AppendLine("\t // Any start up scripts can be placed below and they will run before the test like login for example and in this case the authorization header and retry");
        karateConfig.AppendLine("\t // var example = karate.callSingle('file:example.feature', config);");
        karateConfig.AppendLine("\t // karate.configure('headers', { Authorization: 'Bearer ' + example.token });");
        karateConfig.AppendLine("\t // karate.configure('retry', { count: 5, interval: 5000 });");

        karateConfig.AppendLine("\n\t return config;");
        karateConfig.AppendLine("}");

        string karateConfigString = karateConfig.ToString();
        // Delete the file if it exists.
        using (FileStream fs = File.Create(pathConfig))
        {
            AddText(fs, karateConfigString);
        }

        //Generate Azure Secrets file
        AzureSecrets.AppendLine("{");
        AzureSecrets.AppendLine("\t//Place any azure secrets here that will be used in these tests. Think variable groups or pipeline variables");
        AzureSecrets.AppendLine("}");

        string AzureSecretsString = AzureSecrets.ToString();
        using (FileStream fs = File.Create(pathAzureSec))
        {
            AddText(fs, AzureSecretsString);
        }

        //Generate test.js file 
        testJs.AppendLine("const karate = require('@karatelabs/karate');");

        testJs.AppendLine("\nvar tests = process.env.npm_config_tests");
        testJs.AppendLine("var env = process.env.npm.config_env");

        testJs.AppendLine("\nconsole.log('tests: ', tests)");
        testJs.AppendLine("console.log('env: ', env)");

        testJs.AppendLine("\nif (env === null || env === undefined || env === '') {");
        testJs.AppendLine("\t//environment set to dev03 as default');");
        testJs.AppendLine("\tenv = 'dev03';");
        testJs.AppendLine("\tConsole.log('env not specified; defaulting to :, ', env);");
        testJs.AppendLine("}");

        testJs.AppendLine("\nkarate.exec(`-f cucumber:json, junit.xml ${tests} --env=${env}`);");

        string testJsString = testJs.ToString();
        
        using (FileStream fs = File.Create(pathTestJs))
        {
            AddText(fs, testJsString);
        }

        //Generate package.json
        StringBuilder packageJson = new StringBuilder();
        packageJson.AppendLine("{");
        packageJson.AppendLine("\t \"dependencies\" : {");
        packageJson.AppendLine("\t\t \"cucumber-html-reporter\": \"^5.2.0\",");
        packageJson.AppendLine("\t\t \"copyfiles\": \"^2.4.0\",");
        packageJson.AppendLine("\t\t \"minimist\": \"*\",");
        packageJson.AppendLine("\t\t \"xhr2\": \"^0.2.1\"");
        packageJson.AppendLine("\t },");
        packageJson.AppendLine("\t \"scripts\": {");
        packageJson.AppendLine("\t\t \"test\": \"node test.js\"");
        packageJson.AppendLine("\t },");
        packageJson.AppendLine("\t \"devDependencies\": {");
        packageJson.AppendLine("\t\t \"@karatelabs/karate\": \"*\"");
        packageJson.AppendLine("\t }");
        packageJson.AppendLine("}");

        string packageJsonString = packageJson.ToString();
        using (FileStream fs = File.Create(pathPackageJson))
        {
            AddText(fs, packageJsonString);
        }
    }

    //Function that generates the shared files for the different environments in the configuration folder
    private void generateSharedFiles(string path){
        if (!File.Exists(path))
        {
            StringBuilder sharedFile = new StringBuilder();
            sharedFile.AppendLine("{");
            sharedFile.AppendLine("\n}");

            string sharedFileString = sharedFile.ToString();

            using (FileStream fs = File.Create(path))
            {
                AddText(fs, sharedFileString);
            }
        }
    }

    //Function that generates the karate tests for swagger 3.0
    private void GenerateSwagger3(SwaggerDocument JsonElements, string authentication, string ApiHost, string APIName){
        
        string fileNameExcel = APIName + "-Karate Tests Overview";
         // Create a new Excel package
        ExcelPackage excelPackage = new ExcelPackage();

        //Set a count of each method type to display later
        int postCount = 0; 
        int getCount = 0;
        int putCount = 0;
        int patchCount = 0;
        int deleteCount = 0;

        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

        //Set the titles for the columns in the excel sheet
        worksheet.Cells[1, 2].Value = "Endpoint";
        worksheet.Cells[1, 3].Value = "Method";
        worksheet.Cells[1, 4].Value = "Covered";
        worksheet.Cells[1, 5].Value = "Feature";
        worksheet.Cells[1,2].Style.Fill.PatternType = ExcelFillStyle.Solid;
        worksheet.Cells[1,2].Style.Fill.BackgroundColor.SetColor(Color.Gray);  
        worksheet.Cells[1,3].Style.Fill.PatternType = ExcelFillStyle.Solid;
        worksheet.Cells[1,3].Style.Fill.BackgroundColor.SetColor(Color.Gray);
        worksheet.Cells[1,4].Style.Fill.PatternType = ExcelFillStyle.Solid;
        worksheet.Cells[1,4].Style.Fill.BackgroundColor.SetColor(Color.Gray);
        worksheet.Cells[1,5].Style.Fill.PatternType = ExcelFillStyle.Solid;
        worksheet.Cells[1,5].Style.Fill.BackgroundColor.SetColor(Color.Gray);

        int cellX = 2;

        //Loop through the paths in the swagger JSON
        foreach (var x in JsonElements.paths)
        {
            int methodType = 0;
            string fileName = "";
            string finalScript = "";
            string path = "";
            string statusCode = "";
            string parameterName = "";
            string parameterDescription = "";

            StringBuilder KarateScript = new StringBuilder();

            //Check which method type is used
            if (x.Value.post != null)
            {
                methodType = 1;
            }

            if (x.Value.get != null)
            {
                methodType = 2;
            }

            if (x.Value.put != null)
            {
                methodType = 3;
            }

            if (x.Value.patch != null)
            {
                methodType = 4;
            }

            if (x.Value.delete != null)
            {
                methodType = 5;
            }
            //Switch case for each method type
            switch (methodType)
            {
                case 1:
                    KarateScript.Clear();
                    //replace slash with dash to avoid directoryNotFoundException when creating file
                    fileName = x.Key.ToString();
                    fileName = fileName.Replace("/","-");
                    fileName = fileName.Substring(1);

                    worksheet.Cells[cellX, 2].Value = x.Key.ToString(); 
                    worksheet.Cells[cellX, 3].Value = "POST";
                    worksheet.Cells[cellX,3].Style.Font.Color.SetColor(Color.FromName("Green"));
                    worksheet.Cells[cellX, 4].Value = "-";
                    worksheet.Cells[cellX, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[cellX, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[cellX, 5].Value = fileName + "-POST.feature";

                    KarateScript.AppendLine($"Feature: Karate test for {fileName}");
                    KarateScript.AppendLine($"\n Scenario: Testing {x.Key}");

                    path = @"c:\repos\AutomatedTests\" + APIName + "\\Features\\" + fileName  + "-POST.feature";
                    statusCode = "";
                    parameterName = "";
                    parameterDescription = "";

                    // Delete the file if it exists.
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                    }

                    foreach (var responses in x.Value.post.responses)
                    {
                        if (responses.Key[0].ToString() == "2")
                            statusCode = responses.Key.ToString();
                    }

                    //Verify no null value exists to avoid null exception
                    if (x.Value.post.parameters != null)
                    {
                        foreach (var parameters in x.Value.post.parameters)
                        {
                            if (parameters != null)
                            {
                                parameterName = parameters.name;
                                parameterDescription = parameters.description;
                            }
                        }
                    }

                    //Generates tests with has own property verification
                    KarateScript.AppendLine("\n\t######################");
                    KarateScript.AppendLine("\t#POSITIVE FLOW TESTS#");
                    KarateScript.AppendLine("\t#####################");
                    KarateScript.AppendLine("\n\t* def requestArray =\n");
                    KarateScript.AppendLine("\t \"\"\" ");
                    KarateScript.AppendLine("\t\t{");
                    KarateScript.AppendLine("\t\t\t# Enter all the required parameters, followed by a comma if there's another value underneath, in here like this:");
                    KarateScript.AppendLine("\t\t\t#Testdata:test");
                    KarateScript.AppendLine("\t\t}");
                    KarateScript.AppendLine("\t \"\"\"");
                    KarateScript.AppendLine("\n\t# Verify that endpoint works with only required input");
                    KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{ApiHost}{x.Key.ToString()}'");
                    KarateScript.AppendLine("\tAnd request requestArray");
                    KarateScript.AppendLine("\tWhen method POST");
                    KarateScript.AppendLine($"\tThen status {statusCode}");

                    KarateScript.AppendLine("\n\t# Verify that the newly created entities were successfully created,by using a get call and checking that the newly created ID exists.");
                    KarateScript.AppendLine("\t# Hint: Use a get call from one of the GET templates created for any GET endpoints in this API ");

                    KarateScript.AppendLine("\n\t* def requestArray =\n");
                    KarateScript.AppendLine("\t \"\"\" ");
                    KarateScript.AppendLine("\t\t{");
                    KarateScript.AppendLine("\t\t\t# Enter all the parameters, followed by a comma if there's another value underneath, in here like this:");
                    KarateScript.AppendLine("\t\t\t#Testdata:test");
                    KarateScript.AppendLine("\t\t}");
                    KarateScript.AppendLine("\t \"\"\" ");
                    KarateScript.AppendLine("\n\t# Verify that endpoint works with all parameters");
                    KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{ApiHost}{x.Key.ToString()}'");
                    KarateScript.AppendLine("\tAnd request requestArray");
                    KarateScript.AppendLine("\tWhen method POST");
                    KarateScript.AppendLine($"\tThen status {statusCode}");
                    
                    //Start the negative flow tests
                    KarateScript.AppendLine("\n\t#####################");
                    KarateScript.AppendLine("\t#NEGATIVE FLOW TESTS#");
                    KarateScript.AppendLine("\t#####################");

                    KarateScript.AppendLine("\n\t# Verify that endpoint will give status code 400 when no request is given");
                    KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{ApiHost}{x.Key.ToString()}'");
                    KarateScript.AppendLine("\tWhen method POST");
                    KarateScript.AppendLine($"\tThen status 400");

                    KarateScript.AppendLine("\n\t#  Verify that the endpoint will give status code 400 when an empty request is given");  
                    KarateScript.AppendLine("\t* def requestArray =\n");
                    KarateScript.AppendLine("\t \"\"\" ");
                    KarateScript.AppendLine("\t{}");
                    KarateScript.AppendLine("\n\t \"\"\" ");
                    KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{ApiHost}{x.Key.ToString()}'");
                    KarateScript.AppendLine("\tAnd request requestArray");
                    KarateScript.AppendLine("\tWhen method POST");
                    KarateScript.AppendLine($"\tThen status 400");

                    KarateScript.AppendLine("\n\t# Verify 409 is given when duplicate data is given");
                    KarateScript.AppendLine("\t* def requestArray =\n");
                    KarateScript.AppendLine("\t \"\"\" ");
                    KarateScript.AppendLine(" \n\t\t{\n\t\t\t#Enter the same parameters as one of the two positive flow tests above");
                    KarateScript.AppendLine("\t\t}\n");
                    KarateScript.AppendLine("\t \"\"\" ");
                    KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{ApiHost}{x.Key.ToString()}'");
                    KarateScript.AppendLine("\tAnd request requestArray");
                    KarateScript.AppendLine("\tWhen method POST");
                    KarateScript.AppendLine($"\tThen status 409");

                    finalScript = KarateScript.ToString();

                    using (FileStream fs = File.Create(path))
                    {
                        AddText(fs, finalScript);
                    }
                    cellX++;
                    postCount++;
                    break;

                case 2:
                    KarateScript.Clear();
                    //replace slash with dash to avoid directoryNotFoundException when creating file
                    fileName = x.Key.ToString();
                    fileName = fileName.Replace("/","-");
                    fileName = fileName.Substring(1);

                    worksheet.Cells[cellX, 2].Value = x.Key.ToString(); 
                    worksheet.Cells[cellX, 3].Value = "GET";
                    worksheet.Cells[cellX, 3].Style.Font.Color.SetColor(Color.FromName("Blue"));
                    worksheet.Cells[cellX, 4].Value = "-";
                    worksheet.Cells[cellX, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[cellX, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[cellX, 5].Value = fileName + "-GET.feature";


                    KarateScript.AppendLine($"Feature: Karate test for {fileName}");
                    KarateScript.AppendLine($"\n Scenario: Testing {x.Key}");

                    path = @"c:\repos\AutomatedTests\" + APIName + "\\Features\\" + fileName  + "-GET.feature";
                    statusCode = "";
                    parameterName = "";
                    parameterDescription = "";

                    // Delete the file if it exists.
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                    }
                    foreach (var responses in x.Value.get.responses)
                    {   
                        if(responses.Key[0].ToString() == "2")
                            statusCode = responses.Key.ToString();
                        }

                        //Verify no null value exists to avoid null exception
                        if(x.Value.get.parameters != null){
                            foreach(var parameters in x.Value.get.parameters)
                            {
                                if(parameters != null){
                                    parameterName = parameters.name;
                                    parameterDescription = parameters.description;
                                }
                            }
                        }
                        
                    //Generates tests with has own property verification
                    KarateScript.AppendLine("\n\t######################");
                    KarateScript.AppendLine("\t#POSITIVE FLOW TESTS#");
                    KarateScript.AppendLine("\t#####################");
                    KarateScript.AppendLine("\n\t# Verify that endpoint works with ordinary input");
                    KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{ApiHost}{x.Key.ToString()}'");
                    KarateScript.AppendLine("\tWhen method GET");
                    KarateScript.AppendLine($"\tThen status {statusCode}\n");

                    //Generates test asserting that response does not have null value for any required parameters
                    KarateScript.AppendLine("\t# Verify that response is as expected");
                    KarateScript.AppendLine("\t# Use * assert response.insertPropertyName !# null to verify that the required request parameters are returned\n\t#For example: * assert response.example = #string or null. If it's a string for example");

                    //Assert response is JSON
                    KarateScript.AppendLine("\n\t# Assert response is JSON");
                    KarateScript.AppendLine("\tAnd match responseType == 'json'");

                    //Assert that response time is realistic
                    KarateScript.AppendLine("\n\t# Assert response time is realistic");
                    KarateScript.AppendLine("\tAnd assert responseTime < 1000");

                    KarateScript.AppendLine("\n\t# Assert response headers are as follows:");
                    KarateScript.AppendLine("\tAnd assert responseHeaders['X-Powered-By'] == '#null'");
                    KarateScript.AppendLine("\tAnd assert responseHeaders[\"Cache-Control\"] == \"no-store\"");
                    KarateScript.AppendLine("\tAnd assert responseHeaders[\"Content-Security-Policy\"] == \"frame-ancestors 'none'\"");
                    KarateScript.AppendLine("\tAnd assert responseHeaders[\"Strict-Transport-Security\"] == \"max-age=63072000; includeSubDomains; preload\"");
                    KarateScript.AppendLine("\tAnd assert responseHeaders[\"X-Content-Type-Options\"] == \"nosniff\"");
                    KarateScript.AppendLine("\tAnd assert responseHeaders[\"X-Frame-Options\"] == \"DENY\"");

                    //Start the negative flow tests
                    KarateScript.AppendLine("\n\n\t#####################");
                    KarateScript.AppendLine("\t#NEGATIVE FLOW TESTS#");
                    KarateScript.AppendLine("\t#####################");
                    KarateScript.AppendLine("\n\t# Verify that error 404 is returned when data is not found due to incorrect input");
                    KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{ApiHost}{x.Key}/191'");
                    KarateScript.AppendLine("\tWhen method GET");
                    KarateScript.AppendLine("\tThen status 404");

                    //Needs to be figured out can use URL fix
                    KarateScript.AppendLine("\n\t#Verify that error 405 when a required parameter is missing");
                    KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{ApiHost}{x.Key}'");
                    KarateScript.AppendLine("\tWhen method GET");
                    KarateScript.AppendLine("\tThen status 405");

                    if (authentication.Contains("Bearer"))
                    {
                        //Use incorrect authorization type need to figure out some sort of if or switch
                        KarateScript.AppendLine("\n\t#Verify that access is blocked with incorrect authorization type");
                        KarateScript.AppendLine($"\tGiven header X-ApiKey = '{authentication}'");
                    }
                    else
                    {
                        //Use incorrect authorization type need to figure out some sort of if or switch
                        KarateScript.AppendLine("\n\t#Verify that access is blocked with incorrect authorization type");
                        KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    }

                    KarateScript.AppendLine($"\tGiven url 'https://{ApiHost}{x.Key}'");
                    KarateScript.AppendLine("\tWhen method GET");
                    KarateScript.AppendLine("\tThen status 401");
                    
                    finalScript = KarateScript.ToString();

                    using (FileStream fs = File.Create(path))
                    {
                        AddText(fs, finalScript);
                    }
                    cellX++;
                    getCount++;
                    break;
                case 3:
                    KarateScript.Clear();
                    //replace slash with dash to avoid directoryNotFoundException when creating file
                    fileName = x.Key.ToString();
                    fileName = fileName.Replace("/","-");
                    fileName = fileName.Substring(1);

                    worksheet.Cells[cellX, 2].Value = x.Key.ToString(); 
                    worksheet.Cells[cellX, 3].Value = "PUT";
                    worksheet.Cells[cellX,3].Style.Font.Color.SetColor(Color.FromName("Gold"));
                    worksheet.Cells[cellX, 4].Value = "-";
                    worksheet.Cells[cellX, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[cellX, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[cellX, 5].Value = fileName + "-PUT.feature";


                    KarateScript.AppendLine($"Feature: Karate test for {fileName}");
                    KarateScript.AppendLine($"\n Scenario: Testing {x.Key}");

                    path = @"c:\repos\AutomatedTests\" + APIName + "\\Features\\" + fileName  + "-PUT.feature";
            
                    // Delete the file if it exists.
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                    }

                    //Generates tests with has own property verification
                    KarateScript.AppendLine("\n\t######################");
                    KarateScript.AppendLine("\t#POSITIVE FLOW TESTS#");
                    KarateScript.AppendLine("\t#####################");

                    KarateScript.AppendLine("\n\t# Use pre-existing data that is INDEPENDENT of other tests or create new data");
                    KarateScript.AppendLine("\n\t* def requestArray =\n");
                    KarateScript.AppendLine("\t \"\"\" ");
                    KarateScript.AppendLine("\t\t{");
                    KarateScript.AppendLine("\t\t\t# Update the newly created object with minimum created input");
                    KarateScript.AppendLine("\t\t}");
                    KarateScript.AppendLine("\t \"\"\" ");

                    KarateScript.AppendLine("\n\tGiven header Authorization = '" + authentication + "'");
                    KarateScript.AppendLine("\tGiven url 'https://" + ApiHost + x.Key + "'");
                    KarateScript.AppendLine("\tWhen method PUT");
                    KarateScript.AppendLine("\tThen status 204");

                    KarateScript.AppendLine("\n\t# Verify that the newly created entities were successfully created,by using a get call and checking that the newly created ID exists.");
                    KarateScript.AppendLine("\t# Hint: Use a get call from one of the GET templates created for any GET endpoints in this API ");

                    KarateScript.AppendLine("\n\t* def requestArray =");
                    KarateScript.AppendLine("\t \"\"\" ");
                    KarateScript.AppendLine("\t\t{");
                    KarateScript.AppendLine("\t\t\t# Update the newly created object with all parameters in the model");
                    KarateScript.AppendLine("\t\t}");
                    KarateScript.AppendLine("\t \"\"\" ");

                    KarateScript.AppendLine("\n\tGiven header Authorization = '" + authentication + "'");
                    KarateScript.AppendLine("\tGiven url 'https://" + ApiHost + x.Key + "'");
                    KarateScript.AppendLine("\tWhen method PUT");
                    KarateScript.AppendLine("\tThen status 204");

                    //Start the negative flow tests
                    KarateScript.AppendLine("\n\t#####################");
                    KarateScript.AppendLine("\t#NEGATIVE FLOW TESTS#");
                    KarateScript.AppendLine("\t#####################");

                    KarateScript.AppendLine("\t# Send incorrect content type for the type to get status 500, for example DateTime when string is expected");
                    KarateScript.AppendLine("\n\t* def requestArray =\n");
                    KarateScript.AppendLine("\t \"\"\" ");
                    KarateScript.AppendLine("\t\t{");
                    KarateScript.AppendLine("\t\t\t# Give a string instead of dateBirth for example ");
                    KarateScript.AppendLine("\t\t\tdateOfBirth: 'Test'");
                    KarateScript.AppendLine("\t\t}");
                    KarateScript.AppendLine("\t \"\"\" ");

                    KarateScript.AppendLine("\n\t#  Verify that the endpoint will give status code 400 when an empty request is given");  
                    KarateScript.AppendLine("\t* def requestArray =\n");
                    KarateScript.AppendLine("\t \"\"\" \n");
                    KarateScript.AppendLine("\t{}");
                    KarateScript.AppendLine("\n\t \"\"\" ");
                    KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{ApiHost}{x.Key.ToString()}'");
                    KarateScript.AppendLine("\tAnd request requestArray");
                    KarateScript.AppendLine("\tWhen method PUT");
                    KarateScript.AppendLine($"\tThen status 400");

                    finalScript = KarateScript.ToString();

                    using (FileStream fs = File.Create(path))
                    {
                        AddText(fs, finalScript);
                    }
                    cellX++;
                    putCount++;
                    break;
                case 4:
                    KarateScript.Clear();
                    //replace slash with dash to avoid directoryNotFoundException when creating file
                    fileName = x.Key.ToString();
                    fileName = fileName.Replace("/","-");
                    fileName = fileName.Substring(1);

                    worksheet.Cells[cellX, 2].Value = x.Key.ToString(); 
                    worksheet.Cells[cellX, 3].Value = "PATCH";
                    worksheet.Cells[cellX,3].Style.Font.Color.SetColor(Color.FromName("Orange"));
                    worksheet.Cells[cellX, 4].Value = "-";
                    worksheet.Cells[cellX, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[cellX, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[cellX, 5].Value = fileName + "-PATCH.feature";

                    KarateScript.AppendLine($"Feature: Karate test for {fileName}");
                    KarateScript.AppendLine($"\n Scenario: Testing {x.Key}");

                    path = @"c:\repos\AutomatedTests\" + APIName + "\\Features\\" + fileName  + "-PATCH.feature";
            
                    // Delete the file if it exists.
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                    }

                    //Generates tests with has own property verification
                    KarateScript.AppendLine("\n\t######################");
                    KarateScript.AppendLine("\t#POSITIVE FLOW TESTS#");
                    KarateScript.AppendLine("\t#####################");

                    KarateScript.AppendLine("\n\t# Use pre-existing data that is INDEPENDENT of other tests or create new data");
                    KarateScript.AppendLine("\n\t* def requestArray =\n");
                    KarateScript.AppendLine("\t \"\"\" ");
                    KarateScript.AppendLine("\t\t{");
                    KarateScript.AppendLine("\t\t\t# Update the object with minimum required input");
                    KarateScript.AppendLine("\t\t}");
                    KarateScript.AppendLine("\t \"\"\" ");

                    KarateScript.AppendLine("\n\tGiven header Authorization = '" + authentication + "'");
                    KarateScript.AppendLine("\tGiven url 'https://" + ApiHost + x.Key + "'");
                    KarateScript.AppendLine("\tWhen method PATCH");
                    KarateScript.AppendLine("\tThen status 204");

                    KarateScript.AppendLine("\n\t# Verify that the newly updated entities were successfully updated,by using a get call and checking that the newly created ID exists.");
                    KarateScript.AppendLine("\t# Hint: Use a get call from one of the GET templates created for any GET endpoints in this API ");

                    KarateScript.AppendLine("\n\t* def requestArray =");
                    KarateScript.AppendLine("\t \"\"\" ");
                    KarateScript.AppendLine("\t\t{");
                    KarateScript.AppendLine("\t\t\t# Update the newly created object with all parameters in the model");
                    KarateScript.AppendLine("\t\t}");
                    KarateScript.AppendLine("\t \"\"\" ");

                    KarateScript.AppendLine("\n\tGiven header Authorization = '" + authentication + "'");
                    KarateScript.AppendLine("\tGiven url 'https://" + ApiHost + x.Key + "'");
                    KarateScript.AppendLine("\tWhen method PATCH");
                    KarateScript.AppendLine("\tThen status 204");

                    //Start the negative flow tests
                    KarateScript.AppendLine("\n\t#####################");
                    KarateScript.AppendLine("\t#NEGATIVE FLOW TESTS#");
                    KarateScript.AppendLine("\t#####################");

                    KarateScript.AppendLine("\t# Send incorrect content type for the type to get status 500, for example DateTime when string is expected");
                    KarateScript.AppendLine("\n\t* def requestArray =\n");
                    KarateScript.AppendLine("\t \"\"\" ");
                    KarateScript.AppendLine("\t\t{");
                    KarateScript.AppendLine("\t\t\t# Give a string instead of dateBirth for example ");
                    KarateScript.AppendLine("\t\t\tdateOfBirth: 'Test'");
                    KarateScript.AppendLine("\t\t}");
                    KarateScript.AppendLine("\t \"\"\" ");

                    KarateScript.AppendLine("\n\t#  Verify that the endpoint will give status code 400 when an empty request is given");  
                    KarateScript.AppendLine("\t* def requestArray =\n");
                    KarateScript.AppendLine("\t \"\"\" \n");
                    KarateScript.AppendLine("\t{}");
                    KarateScript.AppendLine("\n\t \"\"\" ");
                    KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{ApiHost}{x.Key.ToString()}'");
                    KarateScript.AppendLine("\tAnd request requestArray");
                    KarateScript.AppendLine("\tWhen method PATCH");
                    KarateScript.AppendLine($"\tThen status 400");

                    finalScript = KarateScript.ToString();

                    using (FileStream fs = File.Create(path))
                    {
                        AddText(fs, finalScript);
                    }
                    cellX++;
                    patchCount++;
                    break;
                case 5:
                    KarateScript.Clear();
                    //replace slash with dash to avoid directoryNotFoundException when creating file
                    fileName = x.Key.ToString();
                    fileName = fileName.Replace("/","-");
                    fileName = fileName.Substring(1);

                    worksheet.Cells[cellX, 2].Value = x.Key.ToString(); 
                    worksheet.Cells[cellX, 3].Value = "DELETE";
                    worksheet.Cells[cellX, 3].Style.Font.Color.SetColor(Color.FromName("RED"));
                    worksheet.Cells[cellX, 4].Value = "-";
                    worksheet.Cells[cellX, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[cellX, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[cellX, 5].Value = fileName + "-DELETE.feature";


                    KarateScript.AppendLine($"Feature: Karate test for {fileName}");
                    KarateScript.AppendLine($"\n Scenario: Testing {x.Key}");

                    path = @"c:\repos\AutomatedTests\" + APIName + "\\Features\\" + fileName  + "-DELETE.feature";
                    statusCode = "";
                    parameterName = "";
                    parameterDescription = "";

                    // Delete the file if it exists.
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                    }
                    foreach (var responses in x.Value.get.responses)
                    {   
                        if(responses.Key[0].ToString() == "2")
                            statusCode = responses.Key.ToString();
                    }

                    //Verify no null value exists to avoid null exception
                    if(x.Value.get.parameters != null){
                        foreach(var parameters in x.Value.get.parameters)
                        {
                            if(parameters != null){
                                parameterName = parameters.name;
                                parameterDescription = parameters.description;
                            }
                        }
                    }
                    
                    //Generates tests with has own property verification
                    KarateScript.AppendLine("\n\t######################");
                    KarateScript.AppendLine("\t#POSITIVE FLOW TESTS#");
                    KarateScript.AppendLine("\t#####################");
                    KarateScript.AppendLine("\n\t# Verify that endpoint works with ordinary input");
                    KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{ApiHost}{x.Key.ToString()}'");
                    KarateScript.AppendLine("\tWhen method DELETE");
                    KarateScript.AppendLine($"\tThen status {statusCode}\n");

                    //Generates test asserting that response does not have null value for any required parameters
                    KarateScript.AppendLine("\t# Verify that response is as expected");
                    KarateScript.AppendLine("\t# Use * assert response.insertPropertyName !# null to verify that the required request parameters are returned\n\t#For example: * assert response.example = #string or null. If it's a string for example");

                    //Assert response is JSON
                    KarateScript.AppendLine("\n\t# Assert response is JSON");
                    KarateScript.AppendLine("\tAnd match responseType == 'json'");

                    //Assert that response time is realistic
                    KarateScript.AppendLine("\n\t# Assert response time is realistic");
                    KarateScript.AppendLine("\tAnd assert responseTime < 1000");

                    KarateScript.AppendLine("\n\t# Assert response headers are as follows:");
                    KarateScript.AppendLine("\tAnd assert responseHeaders['X-Powered-By'] == '#null'");
                    KarateScript.AppendLine("\tAnd assert responseHeaders[\"Cache-Control\"] == \"no-store\"");
                    KarateScript.AppendLine("\tAnd assert responseHeaders[\"Content-Security-Policy\"] == \"frame-ancestors 'none'\"");
                    KarateScript.AppendLine("\tAnd assert responseHeaders[\"Strict-Transport-Security\"] == \"max-age=63072000; includeSubDomains; preload\"");
                    KarateScript.AppendLine("\tAnd assert responseHeaders[\"X-Content-Type-Options\"] == \"nosniff\"");
                    KarateScript.AppendLine("\tAnd assert responseHeaders[\"X-Frame-Options\"] == \"DENY\"");

                    KarateScript.AppendLine("\n\t# Verify that the newly deleted entities were successfully deleted,by using a GET call and checking that the newly deleted Id no longer exists.");
                    KarateScript.AppendLine("\t# Hint: Use a GET call from one of the GET templates created for any GET endpoints in this API and check for a 404 to be sure it has been deleted. ");

                    //Start the negative flow tests
                    KarateScript.AppendLine("\n\n\t#####################");
                    KarateScript.AppendLine("\t#NEGATIVE FLOW TESTS#");
                    KarateScript.AppendLine("\t#####################");
                    KarateScript.AppendLine("\n\t# Verify that error 404 is returned when data is not found due to incorrect input");
                    KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{ApiHost}{x.Key}/191'");
                    KarateScript.AppendLine("\tWhen method DELETE");
                    KarateScript.AppendLine("\tThen status 404");

                    //Needs to be figured out can use URL fix
                    KarateScript.AppendLine("\n\t#Verify that error 405 when a required parameter is missing");
                    KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{ApiHost}{x.Key}'");
                    KarateScript.AppendLine("\tWhen method DELETE");
                    KarateScript.AppendLine("\tThen status 405");

                    if (authentication.Contains("Bearer"))
                    {
                        //Use incorrect authorization type need to figure out some sort of if or switch
                        KarateScript.AppendLine("\n\t#Verify that access is blocked with incorrect authorization type");
                        KarateScript.AppendLine($"\tGiven header X-ApiKey = '{authentication}'");
                    }
                    else
                    {
                        //Use incorrect authorization type need to figure out some sort of if or switch
                        KarateScript.AppendLine("\n\t#Verify that access is blocked with incorrect authorization type");
                        KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    }

                    KarateScript.AppendLine($"\tGiven url 'https://{ApiHost}{x.Key}'");
                    KarateScript.AppendLine("\tWhen method DELETE");
                    KarateScript.AppendLine("\tThen status 401");
                    
                    finalScript = KarateScript.ToString();

                    using (FileStream fs = File.Create(path))
                    {
                        AddText(fs, finalScript);
                    }
                    cellX++;
                    deleteCount++;
                    break;
                default:
                    break;
            }
        }

        string filePath = @"c:\repos\AutomatedTests\" + APIName +  "\\Features\\"  + fileNameExcel + ".xlsx";
        worksheet.Cells.AutoFitColumns();
        excelPackage.SaveAs(new FileInfo(filePath));
        int totalEndpoints = cellX - 2;
        Console.WriteLine("I am done generating the tests, please check the following folder: C:\\Repos\\AutomatedTests\\ " + APIName + " for the tests. I have generated " + totalEndpoints + " tests for all " + totalEndpoints + " endpoints.");
        Console.WriteLine("The following are the number of tests generated for each method type:\nGET: " + getCount + "\nPOST: " + postCount + "\nPUT: " + putCount + "\nPATCH: " + patchCount + "\nDELETE: " + deleteCount + "\n\n");
        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }

    //Method to generate the templates for Swagger 2.0
    private void GenerateSwagger2(SwaggerDocument JsonElements, string authentication, string APIName){

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial; 
        ExcelPackage excelPackage = new ExcelPackage();
        string fileNameExcel = APIName + "-Karate Tests Overview";

        //Set a count of each method type to display later
        int postCount = 0; 
        int getCount = 0;
        int putCount = 0;
        int patchCount = 0;
        int deleteCount = 0;

        // Create a new worksheet
        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
        int cellX = 2;
        //Set the cell values
        worksheet.Cells[1, 2].Value = "Api/Endpoint";
        worksheet.Cells[1, 3].Value = "Method";
        worksheet.Cells[1, 4].Value = "Covered";
        worksheet.Cells[1, 5].Value = "Feature";
        worksheet.Cells[1,2].Style.Fill.PatternType = ExcelFillStyle.Solid;
        worksheet.Cells[1,2].Style.Fill.BackgroundColor.SetColor(Color.Gray);  
        worksheet.Cells[1,3].Style.Fill.PatternType = ExcelFillStyle.Solid;
        worksheet.Cells[1,3].Style.Fill.BackgroundColor.SetColor(Color.Gray);
        worksheet.Cells[1,4].Style.Fill.PatternType = ExcelFillStyle.Solid;
        worksheet.Cells[1,4].Style.Fill.BackgroundColor.SetColor(Color.Gray);
        worksheet.Cells[1,5].Style.Fill.PatternType = ExcelFillStyle.Solid;
        worksheet.Cells[1,5].Style.Fill.BackgroundColor.SetColor(Color.Gray);

        //Iterate through the paths
        foreach (var x in JsonElements.paths)
        {

            string finalScript = "";
            string path = "";
            string fileName = "";
            string statusCode = "";

            int methodType = 0;

            StringBuilder KarateScript = new StringBuilder();

            //Check the method type
            if (x.Value.post != null)
            {
                methodType = 1;
            }

            if (x.Value.get != null)
            {
                methodType = 2;
            }

            if (x.Value.put != null)
            {
                methodType = 3;
            }

            if (x.Value.patch != null)
            {
                methodType = 4;
            }

            if (x.Value.delete != null)
            {
                methodType = 5;
            }

            //Switch statement to generate the tests
            switch (methodType)
            {
                case 1:
                    //replace slash with dash to avoid directoryNotFoundException when creating file
                    fileName = x.Key.ToString();
                    fileName = fileName.Replace("/","-");
                    fileName = fileName.Substring(1);

                    worksheet.Cells[cellX, 2].Value = x.Key.ToString(); 
                    worksheet.Cells[cellX, 3].Value = "POST";
                    worksheet.Cells[cellX, 4].Value = "-";
                    worksheet.Cells[cellX, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[cellX, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[cellX,3].Style.Font.Color.SetColor(Color.FromName("Green"));
                    worksheet.Cells[cellX, 5].Value = fileName + "-POST.feature";

                    KarateScript.Clear();
                    KarateScript.AppendLine($"Feature: Karate test for {fileName}");
                    KarateScript.AppendLine($"\n Scenario: Testing {x.Key}");

                    path = @"c:\repos\AutomatedTests\" + APIName + "\\Features\\" + fileName  + "-POST.feature";
                    statusCode = "";

                    // Delete the file if it exists.
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                    }

                    //Generates tests with has own property verification
                    KarateScript.AppendLine("\n\t######################");
                    KarateScript.AppendLine("\t#POSITIVE FLOW TESTS#");
                    KarateScript.AppendLine("\t#####################");
                    KarateScript.AppendLine("\n\t* def requestArray =\n");
                    KarateScript.AppendLine("\t \"\"\" ");
                    KarateScript.AppendLine("\t\t{");
                    KarateScript.AppendLine("\t\t\t# Enter all the required parameters, followed by a comma if there's another value underneath, in here like this:");
                    KarateScript.AppendLine("\t\t\t#Testdata:test");
                    KarateScript.AppendLine("\t\t}");
                    KarateScript.AppendLine("\t \"\"\"");
                    KarateScript.AppendLine("\n\t# Verify that endpoint works with only required input");
                    KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{JsonElements.host}{x.Key}'");
                    KarateScript.AppendLine("\tAnd request requestArray");
                    KarateScript.AppendLine("\tWhen method POST");
                    KarateScript.AppendLine($"\tThen status {statusCode}");

                    KarateScript.AppendLine("\n\t# Verify that the newly created entities were successfully created,by using a get call and checking that the newly created ID exists.");
                    KarateScript.AppendLine("\t# Hint: Use a get call from one of the GET templates created for any GET endpoints in this API ");

                    KarateScript.AppendLine("\n\t* def requestArray =\n");
                    KarateScript.AppendLine("\t \"\"\" ");
                    KarateScript.AppendLine("\t\t{");
                    KarateScript.AppendLine("\t\t\t# Enter all the parameters, followed by a comma if there's another value underneath, in here like this:");
                    KarateScript.AppendLine("\t\t\t#Testdata:test");
                    KarateScript.AppendLine("\t\t}");
                    KarateScript.AppendLine("\t \"\"\" ");
                    KarateScript.AppendLine("\n\t# Verify that endpoint works with all parameters");
                    KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{JsonElements.host}{x.Key}'");
                    KarateScript.AppendLine("\tAnd request requestArray");
                    KarateScript.AppendLine("\tWhen method POST");
                    KarateScript.AppendLine($"\tThen status {statusCode}");

                    //Start the negative flow tests
                    KarateScript.AppendLine("\n\t#####################");
                    KarateScript.AppendLine("\t#NEGATIVE FLOW TESTS#");
                    KarateScript.AppendLine("\t#####################");

                    KarateScript.AppendLine("\n\t# Verify that endpoint will give status code 400 when no request is given");
                    KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{JsonElements.host}{x.Key}'");
                    KarateScript.AppendLine("\tWhen method POST");
                    KarateScript.AppendLine($"\tThen status 400");

                    KarateScript.AppendLine("\n\t#  Verify that the endpoint will give status code 400 when an empty request is given");  
                    KarateScript.AppendLine("\t* def requestArray =\n");
                    KarateScript.AppendLine("\t \"\"\" ");
                    KarateScript.AppendLine("\t{}");
                    KarateScript.AppendLine("\n\t \"\"\" ");
                    KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{JsonElements.host}{x.Key}'");
                    KarateScript.AppendLine("\tAnd request requestArray");
                    KarateScript.AppendLine("\tWhen method POST");
                    KarateScript.AppendLine($"\tThen status 400");

                    KarateScript.AppendLine("\n\t# Verify 409 is given when duplicate data is given");
                    KarateScript.AppendLine("\t* def requestArray =\n");
                    KarateScript.AppendLine("\t \"\"\" ");
                    KarateScript.AppendLine(" \n\t\t{\n\t\t\t#Enter the same parameters as one of the two positive flow tests above");
                    KarateScript.AppendLine("\t\t}\n");
                    KarateScript.AppendLine("\t \"\"\" ");
                    KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{JsonElements.host}{x.Key}'");
                    KarateScript.AppendLine("\tAnd request requestArray");
                    KarateScript.AppendLine("\tWhen method POST");
                    KarateScript.AppendLine($"\tThen status 409");

                    finalScript = KarateScript.ToString();

                    using (FileStream fs = File.Create(path))
                    {
                        AddText(fs, finalScript);
                    }
                    cellX++;
                    postCount++;
                    break;
                case 2:
                    fileName = x.Key.ToString();
                    fileName = fileName.Replace("/","-");
                    fileName = fileName.Substring(1);

                    worksheet.Cells[cellX, 2].Value = x.Key.ToString(); 
                    worksheet.Cells[cellX, 3].Value = "GET";
                    worksheet.Cells[cellX, 4].Value = "-";
                    worksheet.Cells[cellX, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[cellX, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[cellX,3].Style.Font.Color.SetColor(Color.FromName("Blue"));
                    worksheet.Cells[cellX, 5].Value = fileName + "-GET.feature";

                    KarateScript.Clear();
                    KarateScript.AppendLine($"Feature: Karate test for {fileName}");
                    KarateScript.AppendLine($"\n Scenario: Testing {x.Key}");

                    path = @"c:\repos\AutomatedTests\" + APIName + "\\Features\\" + fileName  + "-GET.feature";
                    statusCode = "";

                    // Delete the file if it exists.
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                    }

                    //Generates tests with has own property verification
                    KarateScript.AppendLine("\n\t######################");
                    KarateScript.AppendLine("\t#POSITIVE FLOW TESTS#");
                    KarateScript.AppendLine("\t#####################");
                    KarateScript.AppendLine("\n\t# Verify that endpoint works with ordinary input");
                    KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{JsonElements.host}{x.Key}'");
                    KarateScript.AppendLine("\tWhen method GET");
                    KarateScript.AppendLine($"\tThen status {statusCode}\n");

                    //Generates test asserting that response does not have null value for any required parameters
                    KarateScript.AppendLine("\t# Verify that response is as expected");
                    KarateScript.AppendLine("\t# Use * assert response.insertPropertyName !# null to verify that the required request parameters are returned\n\t#For example: * assert response.example = #string or null. If it's a string for example");

                    //Assert response is JSON
                    KarateScript.AppendLine("\n\t# Assert response is JSON");
                    KarateScript.AppendLine("\tAnd match responseType == 'json'");

                    //Assert that response time is realistic
                    KarateScript.AppendLine("\n\t# Assert response time is realistic");
                    KarateScript.AppendLine("\tAnd assert responseTime < 1000");

                    KarateScript.AppendLine("\n\t# Assert response headers are as follows:");
                    KarateScript.AppendLine("\tAnd assert responseHeaders['X-Powered-By'] == '#null'");
                    KarateScript.AppendLine("\tAnd assert responseHeaders[\"Cache-Control\"] == \"no-store\"");
                    KarateScript.AppendLine("\tAnd assert responseHeaders[\"Content-Security-Policy\"] == \"frame-ancestors 'none'\"");
                    KarateScript.AppendLine("\tAnd assert responseHeaders[\"Strict-Transport-Security\"] == \"max-age=63072000; includeSubDomains; preload\"");
                    KarateScript.AppendLine("\tAnd assert responseHeaders[\"X-Content-Type-Options\"] == \"nosniff\"");
                    KarateScript.AppendLine("\tAnd assert responseHeaders[\"X-Frame-Options\"] == \"DENY\"");

                    //Start the negative flow tests
                    KarateScript.AppendLine("\n\n\t#####################");
                    KarateScript.AppendLine("\t#NEGATIVE FLOW TESTS#");
                    KarateScript.AppendLine("\t#####################");

                    KarateScript.AppendLine("\n\t# Verify that error 404 is returned when data is not found due to incorrect input");
                    KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{JsonElements.host}{x.Key.ToString()}/191'");
                    KarateScript.AppendLine("\tWhen method GET");
                    KarateScript.AppendLine("\tThen status 404");

                    //Needs to be figured out can use URL fix
                    KarateScript.AppendLine("\n\t#Verify that error 405 when a required parameter is missing");
                    KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{JsonElements.host}{x.Key.ToString()}'");
                    KarateScript.AppendLine("\tWhen method GET");
                    KarateScript.AppendLine("\tThen status 405");
                    
                    if (authentication.Contains("Bearer"))
                    {
                        //Use incorrect authorization type need to figure out some sort of if or switch
                        KarateScript.AppendLine("\n\t#Verify that access is blocked with incorrect authorization type");
                        KarateScript.AppendLine($"\tGiven header X-ApiKey = '{authentication}'");
                    }
                    else
                    {
                        //Use incorrect authorization type need to figure out some sort of if or switch
                        KarateScript.AppendLine("\n\t#Verify that access is blocked with incorrect authorization type");
                        KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    }
                    
                    KarateScript.AppendLine($"\tGiven url 'https://{JsonElements.host}{x.Key.ToString()}'");
                    KarateScript.AppendLine("\tWhen method GET");
                    KarateScript.AppendLine("\tThen status 401");
                    finalScript = KarateScript.ToString();

                    using (FileStream fs = File.Create(path))
                    {
                        AddText(fs, finalScript);
                    }
                    cellX++;
                    getCount++;
                    break;  
                case 3: 
                    //replace slash with dash to avoid directoryNotFoundException when creating file
                    fileName = x.Key.ToString();
                    fileName = fileName.Replace("/","-");
                    fileName = fileName.Substring(1);

                    worksheet.Cells[cellX, 2].Value = x.Key.ToString(); 
                    worksheet.Cells[cellX, 3].Value = "PUT";
                    worksheet.Cells[cellX, 3].Style.Font.Color.SetColor(Color.FromName("Gold"));
                    worksheet.Cells[cellX, 4].Value = "-";
                    worksheet.Cells[cellX, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[cellX, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[cellX, 5].Value = fileName + "-PUT.feature";

                    KarateScript.Clear();
                    KarateScript.AppendLine($"Feature: Karate test for {fileName}");
                    KarateScript.AppendLine($"\n Scenario: Testing {x.Key}");

                    path = @"c:\repos\AutomatedTests\" + APIName + "\\Features\\" + fileName  + "-PUT.feature";
            
                    // Delete the file if it exists.
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                    }

                    //Generates tests with has own property verification
                    KarateScript.AppendLine("\n\t######################");
                    KarateScript.AppendLine("\t#POSITIVE FLOW TESTS#");
                    KarateScript.AppendLine("\t#####################");

                    KarateScript.AppendLine("\n\t# Use pre-existing data that is INDEPENDENT of other tests or create new data");
                    KarateScript.AppendLine("\n\t* def requestArray =");
                    KarateScript.AppendLine("\t \"\"\" ");
                    KarateScript.AppendLine("\t\t{");
                    KarateScript.AppendLine("\t\t\t# Update the newly created object with minimum created input");
                    KarateScript.AppendLine("\t\t}");
                    KarateScript.AppendLine("\t \"\"\" ");

                    KarateScript.AppendLine($"\n\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{JsonElements.host}{x.Key}'");
                    KarateScript.AppendLine("\tWhen method PUT");
                    KarateScript.AppendLine("\tThen status 204");

                    KarateScript.AppendLine("\n\t# Verify that the newly created entities were successfully created,by using a get call and checking that the newly created ID exists.");
                    KarateScript.AppendLine("\t#Hint: Use a get call from one of the GET templates created for any GET endpoints in this API ");

                    KarateScript.AppendLine("\n\t* def requestArray =");
                    KarateScript.AppendLine("\t \"\"\" ");
                    KarateScript.AppendLine("\t\t{");
                    KarateScript.AppendLine("\t\t\t# Update the newly created object with all parameters in the model");
                    KarateScript.AppendLine("\t\t}");
                    KarateScript.AppendLine("\t \"\"\" ");

                    KarateScript.AppendLine($"\n\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{JsonElements.host}{x.Key}'");
                    KarateScript.AppendLine("\tWhen method PUT");
                    KarateScript.AppendLine("\tThen status 204");

                    //Start the negative flow tests
                    KarateScript.AppendLine("\n\n\t#####################");
                    KarateScript.AppendLine("\t#NEGATIVE FLOW TESTS#");
                    KarateScript.AppendLine("\t#####################");

                    KarateScript.AppendLine("\t# Send incorrect content type for the type to get status 500, for example DateTime when string is expected");
                    KarateScript.AppendLine("\n\t* def requestArray =\n");
                    KarateScript.AppendLine("\t \"\"\" ");
                    KarateScript.AppendLine("\t\t{");
                    KarateScript.AppendLine("\t\t\t# Give a string instead of dateBirth for example ");
                    KarateScript.AppendLine("\t\t\tdateOfBirth: 'Test'");
                    KarateScript.AppendLine("\t\t}");
                    KarateScript.AppendLine("\t \"\"\" ");

                    KarateScript.AppendLine("\n\t#  Verify that the endpoint will give status code 400 when an empty request is given");  
                    KarateScript.AppendLine("\t* def requestArray =\n");
                    KarateScript.AppendLine("\t \"\"\" ");
                    KarateScript.AppendLine("\t{}");
                    KarateScript.AppendLine("\n\t \"\"\" ");
                    KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{JsonElements.host}{x.Key.ToString()}'");
                    KarateScript.AppendLine("\tAnd request requestArray");
                    KarateScript.AppendLine("\tWhen method PUT");
                    KarateScript.AppendLine($"\tThen status 400");

                    finalScript = KarateScript.ToString();

                    using (FileStream fs = File.Create(path))
                    {
                        AddText(fs, finalScript);
                    }
                    
                    cellX++;
                    putCount++;
                    break;
                case 4:
                    //replace slash with dash to avoid directoryNotFoundException when creating file
                    fileName = x.Key.ToString();
                    fileName = fileName.Replace("/","-");
                    fileName = fileName.Substring(1);

                    worksheet.Cells[cellX, 2].Value = x.Key.ToString(); 
                    worksheet.Cells[cellX, 3].Value = "PATCH";
                    worksheet.Cells[cellX, 3].Style.Font.Color.SetColor(Color.FromName("Orange"));
                    worksheet.Cells[cellX, 4].Value = "-";
                    worksheet.Cells[cellX, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[cellX, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[cellX, 5].Value = fileName + "-PATCH.feature";

                    KarateScript.Clear();
                    KarateScript.AppendLine($"Feature: Karate test for {fileName}");
                    KarateScript.AppendLine($"\n Scenario: Testing {x.Key}");

                    path = @"c:\repos\AutomatedTests\" + APIName + "\\Features\\" + fileName  + "-PATCH.feature";
            
                    // Delete the file if it exists.
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                    }

                    //Generates tests with has own property verification
                    KarateScript.AppendLine("\n\t######################");
                    KarateScript.AppendLine("\t#POSITIVE FLOW TESTS#");
                    KarateScript.AppendLine("\t#####################");

                    KarateScript.AppendLine("\n\t# Use pre-existing data that is INDEPENDENT of other tests or create new data");
                    KarateScript.AppendLine("\n\t* def requestArray =");
                    KarateScript.AppendLine("\t \"\"\" ");
                    KarateScript.AppendLine("\t\t{");
                    KarateScript.AppendLine("\t\t\t# Update the object with minimum required input");
                    KarateScript.AppendLine("\t\t}");
                    KarateScript.AppendLine("\t \"\"\" ");

                    KarateScript.AppendLine($"\n\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{JsonElements.host}{x.Key}'");
                    KarateScript.AppendLine("\tWhen method PATCH");
                    KarateScript.AppendLine("\tThen status 204");

                    KarateScript.AppendLine("\n\t# Verify that the newly updated entities were successfully updated,by using a get call and checking that the newly created ID exists.");
                    KarateScript.AppendLine("\t#Hint: Use a get call from one of the GET templates created for any GET endpoints in this API ");

                    KarateScript.AppendLine("\n\t* def requestArray =");
                    KarateScript.AppendLine("\t \"\"\" ");
                    KarateScript.AppendLine("\t\t{");
                    KarateScript.AppendLine("\t\t\t# Update the newly created object with all parameters in the model");
                    KarateScript.AppendLine("\t\t}");
                    KarateScript.AppendLine("\t \"\"\" ");

                    KarateScript.AppendLine($"\n\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{JsonElements.host}{x.Key}'");
                    KarateScript.AppendLine("\tWhen method PATCH");
                    KarateScript.AppendLine("\tThen status 204");

                    //Start the negative flow tests
                    KarateScript.AppendLine("\n\n\t#####################");
                    KarateScript.AppendLine("\t#NEGATIVE FLOW TESTS#");
                    KarateScript.AppendLine("\t#####################");

                    KarateScript.AppendLine("\t# Send incorrect content type for the type to get status 500, for example DateTime when string is expected");
                    KarateScript.AppendLine("\n\t* def requestArray =\n");
                    KarateScript.AppendLine("\t \"\"\" ");
                    KarateScript.AppendLine("\t\t{");
                    KarateScript.AppendLine("\t\t\t# Give a string instead of dateBirth for example ");
                    KarateScript.AppendLine("\t\t\tdateOfBirth: 'Test'");
                    KarateScript.AppendLine("\t\t}");
                    KarateScript.AppendLine("\t \"\"\" ");

                    KarateScript.AppendLine("\n\t#  Verify that the endpoint will give status code 400 when an empty request is given");  
                    KarateScript.AppendLine("\t* def requestArray =\n");
                    KarateScript.AppendLine("\t \"\"\" ");
                    KarateScript.AppendLine("\t{}");
                    KarateScript.AppendLine("\n\t \"\"\" ");
                    KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{JsonElements.host}{x.Key.ToString()}'");
                    KarateScript.AppendLine("\tAnd request requestArray");
                    KarateScript.AppendLine("\tWhen method PATCH");
                    KarateScript.AppendLine($"\tThen status 400");

                    finalScript = KarateScript.ToString();

                    using (FileStream fs = File.Create(path))
                    {
                        AddText(fs, finalScript);
                    }
                    
                    cellX++;
                    patchCount++;
                    break;
                case 5:
                    fileName = x.Key.ToString();
                    fileName = fileName.Replace("/","-");
                    fileName = fileName.Substring(1);

                    worksheet.Cells[cellX, 2].Value = x.Key.ToString(); 
                    worksheet.Cells[cellX, 3].Value = "DELETE";
                    worksheet.Cells[cellX, 4].Value = "-";
                    worksheet.Cells[cellX, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[cellX, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[cellX,3].Style.Font.Color.SetColor(Color.FromName("RED"));
                    worksheet.Cells[cellX, 5].Value = fileName + "-DELETE.feature";

                    KarateScript.Clear();
                    KarateScript.AppendLine($"Feature: Karate test for {fileName}");
                    KarateScript.AppendLine($"\n Scenario: Testing {x.Key}");

                    path = @"c:\repos\AutomatedTests\" + APIName + "\\Features\\" + fileName  + "-DELETE.feature";
                    statusCode = "";

                    // Delete the file if it exists.
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                    }

                    //Generates tests with has own property verification
                    KarateScript.AppendLine("\n\t######################");
                    KarateScript.AppendLine("\t#POSITIVE FLOW TESTS#");
                    KarateScript.AppendLine("\t#####################");
                    KarateScript.AppendLine("\n\t# Verify that endpoint works with ordinary input");
                    KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{JsonElements.host}{x.Key}'");
                    KarateScript.AppendLine("\tWhen method DELETE");
                    KarateScript.AppendLine($"\tThen status {statusCode}\n");

                    //Generates test asserting that response does not have null value for any required parameters
                    KarateScript.AppendLine("\t# Verify that response is as expected");
                    KarateScript.AppendLine("\t# Use * assert response.insertPropertyName !# null to verify that the required request parameters are returned\n\t#For example: * assert response.example = #string or null. If it's a string for example");

                    //Assert response is JSON
                    KarateScript.AppendLine("\n\t# Assert response is JSON");
                    KarateScript.AppendLine("\tAnd match responseType == 'json'");

                    //Assert that response time is realistic
                    KarateScript.AppendLine("\n\t# Assert response time is realistic");
                    KarateScript.AppendLine("\tAnd assert responseTime < 1000");

                    KarateScript.AppendLine("\n\t# Assert response headers are as follows:");
                    KarateScript.AppendLine("\tAnd assert responseHeaders['X-Powered-By'] == '#null'");
                    KarateScript.AppendLine("\tAnd assert responseHeaders[\"Cache-Control\"] == \"no-store\"");
                    KarateScript.AppendLine("\tAnd assert responseHeaders[\"Content-Security-Policy\"] == \"frame-ancestors 'none'\"");
                    KarateScript.AppendLine("\tAnd assert responseHeaders[\"Strict-Transport-Security\"] == \"max-age=63072000; includeSubDomains; preload\"");
                    KarateScript.AppendLine("\tAnd assert responseHeaders[\"X-Content-Type-Options\"] == \"nosniff\"");
                    KarateScript.AppendLine("\tAnd assert responseHeaders[\"X-Frame-Options\"] == \"DENY\"");

                    KarateScript.AppendLine("\n\t# Verify that the newly deleted entities were successfully deleted,by using a get call and checking that the newly deleted ID no longer exists.");
                    KarateScript.AppendLine("\t# Hint: Use a get call from one of the GET templates created for any GET endpoints in this API and check for a 404 response as it should no longer be found ");

                    //Start the negative flow tests
                    KarateScript.AppendLine("\n\n\t#####################");
                    KarateScript.AppendLine("\t#NEGATIVE FLOW TESTS#");
                    KarateScript.AppendLine("\t#####################");

                    KarateScript.AppendLine("\n\t# Verify that error 404 is returned when data is not found due to incorrect input");
                    KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{JsonElements.host}{x.Key.ToString()}/191'");
                    KarateScript.AppendLine("\tWhen method DELETE");
                    KarateScript.AppendLine("\tThen status 404");

                    //Needs to be figured out can use URL fix
                    KarateScript.AppendLine("\n\t#Verify that error 405 when a required parameter is missing");
                    KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    KarateScript.AppendLine($"\tGiven url 'https://{JsonElements.host}{x.Key.ToString()}'");
                    KarateScript.AppendLine("\tWhen method DELETE");
                    KarateScript.AppendLine("\tThen status 405");
                    
                    if (authentication.Contains("Bearer"))
                    {
                        //Use incorrect authorization type need to figure out some sort of if or switch
                        KarateScript.AppendLine("\n\t#Verify that access is blocked with incorrect authorization type");
                        KarateScript.AppendLine($"\tGiven header X-ApiKey = '{authentication}'");
                    }
                    else
                    {
                        //Use incorrect authorization type need to figure out some sort of if or switch
                        KarateScript.AppendLine("\n\t#Verify that access is blocked with incorrect authorization type");
                        KarateScript.AppendLine($"\tGiven header Authorization = '{authentication}'");
                    }
                    
                    KarateScript.AppendLine($"\tGiven url 'https://{JsonElements.host}{x.Key.ToString()}'");
                    KarateScript.AppendLine("\tWhen method DELETE");
                    KarateScript.AppendLine("\tThen status 401");
                    finalScript = KarateScript.ToString();

                    using (FileStream fs = File.Create(path))
                    {
                        AddText(fs, finalScript);
                    }
                    cellX++;
                    deleteCount++;
                    break;
                default:
                    break;
            }
        }

        // Save the workbook
        string filePath = @"c:\repos\AutomatedTests\" + APIName +  "\\Features\\"  + fileNameExcel + ".xlsx";
        worksheet.Cells.AutoFitColumns();
        excelPackage.SaveAs(new FileInfo(filePath));
        int totalEndpoints = cellX - 2;
        Console.WriteLine("I am done generating the tests, please check the following folder: C:\\Repos\\AutomatedTests\\ " + APIName + " for the tests. I have generated " + totalEndpoints + " tests for all " + totalEndpoints + " endpoints.");
        Console.WriteLine("The following are the number of tests generated for each method type:\nGET: " + getCount + "\nPOST: " + postCount + "\nPUT: " + putCount + "\nPATCH: " + patchCount + "\nDELETE: " + deleteCount + "\n\n");
        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }
}


