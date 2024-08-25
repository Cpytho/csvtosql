using System.Diagnostics;
using System.Text;
using learning.Models;
using Microsoft.AspNetCore.Mvc;
using MySql.Data.MySqlClient;
using System.IO;

namespace learning.Controllers;

[ApiController]

[Route("api/[controller]")]

public class DataController : ControllerBase
{
    //this class will handle various apis for the data

    public IConfiguration? configuration;//this will store the configuration settings present in appsetting.json
    private readonly string? dbConnString;//this is the string required to connect to mysql db


    //initializing the rabbitmq publisher and consumer

    private readonly RabbitMQPublisher publisher;

    public DataController(IConfiguration _configuration, RabbitMQPublisher _publisher)
    {
        //constructor
        configuration = _configuration;//updating the value
        dbConnString = configuration.GetConnectionString("DBConnectionString");//updating the value

        publisher = _publisher;
    }

    [HttpGet]// identifies a action as GET
    public ActionResult<List<DataModel>> GetData()
    {
        //this method returns a actionresult (status code) and list of datamodels

        var listOfData = new List<DataModel> { };//defining a list of datamodels with initally nothing in it


        //using is used to dispose the variable once finished using it which is used only inside {}
        using (var connection = new MySqlConnection(dbConnString))
        {
            connection.Open();

            var query = "SELECT * FROM usertest2";//creating the syntax

            using (var command = new MySqlCommand(query, connection))
            {
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {

                        var data = new DataModel
                        {
                            Email_id = reader.GetString("Email_id"),
                            Name = reader.GetString("Name"),
                            Country = reader.GetString("Country"),
                            State = reader.GetString("State"),
                            City = reader.GetString("City"),
                            Telephone_number = (int)reader.GetInt64("Telephone_number"),
                            Address_line_1 = reader.GetString("Address_line_1"),
                            Address_line_2 = reader.GetString("Address_line_2"),
                            Date_of_birth = reader.GetString("Date_of_birth"),
                            Gross_salary_FY2019_20 = (int)reader.GetInt64("Gross_salary_FY2019_20"),
                            Gross_salary_FY2020_21 = (int)reader.GetInt64("Gross_salary_FY2020_21"),
                            Gross_salary_FY2021_22 = (int)reader.GetInt64("Gross_salary_FY2021_22"),
                            Gross_salary_FY2022_23 = (int)reader.GetInt64("Gross_salary_FY2022_23"),
                            Gross_salary_FY2023_24 = (int)reader.GetInt64("Gross_salary_FY2023_24")

                        };

                        listOfData.Add(data);
                    }

                }

            }


            return Ok(listOfData);
        }

    }

    // POST method to add a new DataModel
    [HttpPost("insertSingleData")]
    public async Task<IActionResult> PostData(DataModel dataModel)
    {
        if (dataModel == null)
        {
            return BadRequest("DataModel object is null.");
        }

        try
        {
            using (var connection = new MySqlConnection(dbConnString))
            {
                await connection.OpenAsync();

                var query = @"
                        INSERT INTO usertest2 
                        (Email_id, Name, Country, State, City, Telephone_number, Address_line_1, Address_line_2, Date_of_birth, Gross_salary_FY2019_20, Gross_salary_FY2020_21, Gross_salary_FY2021_22, Gross_salary_FY2022_23, Gross_salary_FY2023_24) 
                        VALUES 
                        (@Email_id, @Name, @Country, @State, @City, @Telephone_number, @Address_line_1, @Address_line_2, @Date_of_birth, @Gross_salary_FY2019_20, @Gross_salary_FY2020_21, @Gross_salary_FY2021_22, @Gross_salary_FY2022_23, @Gross_salary_FY2023_24);";

                using (var command = new MySqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Email_id", dataModel.Email_id);
                    command.Parameters.AddWithValue("@Name", dataModel.Name);
                    command.Parameters.AddWithValue("@Country", dataModel.Country);
                    command.Parameters.AddWithValue("@State", dataModel.State);
                    command.Parameters.AddWithValue("@City", dataModel.City);
                    command.Parameters.AddWithValue("@Telephone_number", dataModel.Telephone_number);
                    command.Parameters.AddWithValue("@Address_line_1", dataModel.Address_line_1);
                    command.Parameters.AddWithValue("@Address_line_2", dataModel.Address_line_2);
                    command.Parameters.AddWithValue("@Date_of_birth", dataModel.Date_of_birth);
                    command.Parameters.AddWithValue("@Gross_salary_FY2019_20", dataModel.Gross_salary_FY2019_20);
                    command.Parameters.AddWithValue("@Gross_salary_FY2020_21", dataModel.Gross_salary_FY2020_21);
                    command.Parameters.AddWithValue("@Gross_salary_FY2021_22", dataModel.Gross_salary_FY2021_22);
                    command.Parameters.AddWithValue("@Gross_salary_FY2022_23", dataModel.Gross_salary_FY2022_23);
                    command.Parameters.AddWithValue("@Gross_salary_FY2023_24", dataModel.Gross_salary_FY2023_24);

                    await command.ExecuteNonQueryAsync();
                }
            }

            return CreatedAtAction(nameof(GetData), new { email_id = dataModel.Email_id }, dataModel);
            // 1. Action name to retrieve the resource
            // 2. Route values (identifier of the newly created resource)
            // 3. The newly created resource to return in the response body
        }
        catch (Exception ex)
        {
            // Handle exceptions (e.g., log the error)
            return StatusCode(500, $"Internal server error: {ex.Message}");
        }
    }


    //POST method to add multiple DataModels


    // [HttpPost("insertMultipleData")]
    // public async Task<IActionResult> PostMultipleData(IFormFile file)
    // {
    //     Console.WriteLine("-------------");
    //     Console.WriteLine(((DateTimeOffset)DateTime.UtcNow).ToUnixTimeMilliseconds());

    //     Task[] tasks = new Task[10];
    //     int i = 0;
    //     if (file == null || file.Length == 0)
    //     {
    //         return BadRequest("No file uploaded.");
    //     }

    //     tasks[i] = Task.Run(() => ProcessCsvFile4(file.OpenReadStream()));
    //     i++; // Process the file in a streaming manner with chunks
    //     await Task.WhenAll(tasks);
    //     Console.WriteLine(((DateTimeOffset)DateTime.UtcNow).ToUnixTimeMilliseconds());

    //     return Ok();
    // }



    // private async Task ProcessCsvFile4(Stream stream)
    // {
    //     const int batchSize = 10000;
    //     using var reader = new StreamReader(stream);

    //     Console.WriteLine(((DateTimeOffset)DateTime.UtcNow).ToUnixTimeMilliseconds());

    //     await reader.ReadLineAsync(); // Skip header

    //     var records = new List<DataModel>(batchSize);
    //     string? line;

    //     while ((line = await reader.ReadLineAsync()) != null)
    //     {
    //         var columns = line.Split(',');
    //         records.Add(new DataModel
    //         {
    //             Email_id = columns[0],
    //             Name = columns[1],
    //             Country = columns[2],
    //             State = columns[3],
    //             City = columns[4],
    //             Telephone_number = (int)Int64.Parse(columns[5]),
    //             Address_line_1 = columns[6],
    //             Address_line_2 = columns[7],
    //             Date_of_birth = columns[8],
    //             Gross_salary_FY2019_20 = (int)Int64.Parse(columns[9]),
    //             Gross_salary_FY2020_21 = (int)Int64.Parse(columns[10]),
    //             Gross_salary_FY2021_22 = (int)Int64.Parse(columns[11]),
    //             Gross_salary_FY2022_23 = (int)Int64.Parse(columns[12]),
    //             Gross_salary_FY2023_24 = (int)Int64.Parse(columns[13])
    //         });

    //         if (records.Count == batchSize)
    //         {
    //             // Console.WriteLine("Ram");
    //             var recordsCopy = new List<DataModel>(records);
    //             records.Clear();
    //             // Console.WriteLine("siya");

    //            ConvertToQuery(recordsCopy);
    //         }
    //     }

    //     if (records.Count > 0)
    //     {
    //        ConvertToQuery(records);
    //     }


    //     Console.WriteLine(((DateTimeOffset)DateTime.UtcNow).ToUnixTimeMilliseconds());

    // }


    // //functions required to process the csv file ( chunking data and calling bulk query insert for this chunks)

    // private void ConvertToQuery(List<DataModel> records)
    // {
    //     //this function receives a list of records and send it chunk wise to mysql db without using Parameters for query builders


    //     var queryBuilder = new StringBuilder();
    //     queryBuilder.Append("INSERT INTO usertest2 (Email_id, Name, Country, State, City, Telephone_number, Address_line_1, Address_line_2, Date_of_birth, Gross_salary_FY2019_20, Gross_salary_FY2020_21, Gross_salary_FY2021_22, Gross_salary_FY2022_23, Gross_salary_FY2023_24) VALUES ");

    //     bool isFirstValueEntry = true;

    //     foreach (var record in records)
    //     {
    //         if (!isFirstValueEntry)
    //         {
    //             queryBuilder.Append(',');
    //         }

    //         queryBuilder.Append($"(" +
    //             $"'{MySqlHelper.EscapeString(record.Email_id)}', " +
    //             $"'{MySqlHelper.EscapeString(record.Name)}', " +
    //             $"'{MySqlHelper.EscapeString(record.Country)}', " +
    //             $"'{MySqlHelper.EscapeString(record.State)}', " +
    //             $"'{MySqlHelper.EscapeString(record.City)}', " +
    //             $"{record.Telephone_number}, " +
    //             $"'{MySqlHelper.EscapeString(record.Address_line_1)}', " +
    //             $"'{MySqlHelper.EscapeString(record.Address_line_2)}', " +
    //             $"'{MySqlHelper.EscapeString(record.Date_of_birth)}', " +
    //             $"{record.Gross_salary_FY2019_20}, " +
    //             $"{record.Gross_salary_FY2020_21}, " +
    //             $"{record.Gross_salary_FY2021_22}, " +
    //             $"{record.Gross_salary_FY2022_23}, " +
    //             $"{record.Gross_salary_FY2023_24})");

    //         isFirstValueEntry = false;
    //     }

    //     queryBuilder.Append(';');

    //     //query is generated

    //     publisher.SendChunk(queryBuilder.ToString());
    //     // Console.WriteLine("abff");

    // }


    // -------------------------------------------------------------------


    // [HttpPost("insertMultipleData")]
    // public async Task<IActionResult> insertMultipleData(IFormFile file)
    // {
    //     Console.WriteLine("-------------");
    //     Console.WriteLine(((DateTimeOffset)DateTime.UtcNow).ToUnixTimeMilliseconds());

    //     if (file == null || file.Length == 0)
    //     {
    //         return BadRequest("No file uploaded.");
    //     }

    //     using var stream = file.OpenReadStream();
    //     const int linesPerChunk = 10000;  // Set to 10,000 lines per chunk

    //     // We will process in parallel, creating a task for each chunk
    //     var tasks = new List<Task>();
    //     using var reader = new StreamReader(stream);

    //     // Skip header
    //     await reader.ReadLineAsync();

    //     while (!reader.EndOfStream)
    //     {
    //         // Process linesPerChunk lines in each task
    //         tasks.Add(ProcessCsvChunk(reader, linesPerChunk));
    //     }

    //     await Task.WhenAll(tasks);

    //     Console.WriteLine(((DateTimeOffset)DateTime.UtcNow).ToUnixTimeMilliseconds());

    //     return Ok();
    // }

    // private async Task ProcessCsvChunk(StreamReader reader, int linesPerChunk)
    // {
    //     const int batchSize = 10000;
    //     var records = new List<DataModel>(batchSize);

    //     string? line;
    //     int lineCount = 0;

    //     while (lineCount < linesPerChunk && (line = await reader.ReadLineAsync()) != null)
    //     {
    //         if (string.IsNullOrWhiteSpace(line))
    //             continue;

    //         var columns = line.Split(',');
    //         if (columns.Length < 14)
    //             continue;

    //         try
    //         {
    //             records.Add(new DataModel
    //             {
    //                 Email_id = columns[0],
    //                 Name = columns[1],
    //                 Country = columns[2],
    //                 State = columns[3],
    //                 City = columns[4],
    //                 Telephone_number = int.TryParse(columns[5], out int tel) ? tel : 0,
    //                 Address_line_1 = columns[6],
    //                 Address_line_2 = columns[7],
    //                 Date_of_birth = columns[8],
    //                 Gross_salary_FY2019_20 = int.TryParse(columns[9], out int sal19) ? sal19 : 0,
    //                 Gross_salary_FY2020_21 = int.TryParse(columns[10], out int sal20) ? sal20 : 0,
    //                 Gross_salary_FY2021_22 = int.TryParse(columns[11], out int sal21) ? sal21 : 0,
    //                 Gross_salary_FY2022_23 = int.TryParse(columns[12], out int sal22) ? sal22 : 0,
    //                 Gross_salary_FY2023_24 = int.TryParse(columns[13], out int sal23) ? sal23 : 0
    //             });

    //             lineCount++;

    //             if (records.Count == batchSize)
    //             {
    //                 await ConvertToQuery(records);
    //                 records.Clear();
    //             }
    //         }
    //         catch (Exception ex)
    //         {
    //             Console.WriteLine($"Error processing line: {ex.Message}");
    //         }
    //     }

    //     // Process any remaining records
    //     if (records.Count > 0)
    //     {
    //         await ConvertToQuery(records);
    //     }
    // }


    //-----------------------------------------------------------------------------------------------------------------------------


    [HttpPost("uploadAndProcessFile")]
    public async Task<IActionResult> UploadAndProcessFile(IFormFile file)
    {
        if (file == null || file.Length == 0)
        {
            return BadRequest("No file uploaded.");
        }

        string uploadPath = Path.Combine(Path.GetTempPath(), "uploads");
        Directory.CreateDirectory(uploadPath);

        Console.WriteLine(uploadPath);

        int numberOfParts = 10;
        var partFiles = new string[numberOfParts];

        long partSize = file.Length / numberOfParts;

        using (var stream = file.OpenReadStream())
        {
            var tasks = new Task[numberOfParts];
            using var reader = new StreamReader(stream);

            for (int i = 0; i < numberOfParts; i++)
            {
                int partIndex = i;
                partFiles[partIndex] = Path.Combine(uploadPath, $"part_{partIndex}.csv");

                tasks[partIndex] = Task.Run(async () =>
                {
                    using var fileStream = new FileStream(partFiles[partIndex], FileMode.Create);
                    using var writer = new StreamWriter(fileStream);

                    long bytesWritten = 0;
                    char[] buffer = new char[8192];
                    int charsRead;

                    while (bytesWritten < partSize && (charsRead = await reader.ReadAsync(buffer, 0, buffer.Length)) > 0)
                    {
                        await writer.WriteAsync(buffer, 0, charsRead);
                        bytesWritten += Encoding.UTF8.GetByteCount(buffer, 0, charsRead);
                    }

                    // Ensure we complete the last line before splitting
                    if (partIndex < numberOfParts - 1) // Skip for the last part
                    {
                        string? line;
                        while ((line = await reader.ReadLineAsync()) != null)
                        {
                            await writer.WriteLineAsync(line);
                            break;  // Write one additional line to complete the current part
                        }
                    }
                });
            }

            await Task.WhenAll(tasks);
        }

        // Process each part concurrently
        var processTasks = partFiles.Select(partFile =>
            Task.Run(() => ProcessFile(partFile))
        );

        await Task.WhenAll(processTasks);

        return Ok("File uploaded and processed successfully.");
    }



    [HttpPost("uploadAndSplitStream")]
    public async Task<IActionResult> UploadAndSplitStream()
    {
        Console.WriteLine("------------------------------------------");
        Stopwatch stopwatch = new Stopwatch();
        stopwatch.Start();
        string uploadPath = Path.Combine(Path.GetTempPath(), "uploads");
        Directory.CreateDirectory(uploadPath);
        Console.WriteLine(uploadPath);

        int numberOfParts = 10;
        var partFiles = new string[numberOfParts];
        var partStreams = new StreamWriter[numberOfParts];
        var processTasks = new Task[numberOfParts];

        for (int i = 0; i < numberOfParts; i++)
        {
            partFiles[i] = Path.Combine(uploadPath, $"part_{i}.csv");
            partStreams[i] = new StreamWriter(partFiles[i]);
        }

        using var reader = new StreamReader(Request.Body);
        int currentPart = 0;
        long totalLinesRead = 0;
        long linesPerPart = 10000;
        await reader.ReadLineAsync();
        await reader.ReadLineAsync();
        await reader.ReadLineAsync();
        await reader.ReadLineAsync();
        await reader.ReadLineAsync();

        string? line;
        while ((line = await reader.ReadLineAsync()) != null)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            await partStreams[currentPart].WriteLineAsync(line);
            totalLinesRead++;

            if (totalLinesRead >= linesPerPart && currentPart < numberOfParts - 1)
            {
                await partStreams[currentPart].FlushAsync();
                partStreams[currentPart].Dispose();

                int partIndex = currentPart;
                processTasks[partIndex] = Task.Run(() => ProcessFile(partFiles[partIndex]));

                currentPart++;
                totalLinesRead = 0;
            }
        }

        // Handle the last part
        await partStreams[currentPart].FlushAsync();
        partStreams[currentPart].Dispose();
        processTasks[currentPart] = Task.Run(() => ProcessFile(partFiles[currentPart]));

        // Wait for all processing tasks to complete
        await Task.WhenAll(processTasks.Where(t => t != null));
        stopwatch.Stop();
        TimeSpan elapsed = stopwatch.Elapsed;
        Console.WriteLine($"Total Elapsed time: {elapsed}");

        return Ok("File uploaded, split, and processed successfully.");
    }


    private async Task ProcessFile(string filePath)
    {
        // Simulate processing of the file
        using var reader = new StreamReader(filePath);
        string? line;
        var records = new List<DataModel>();

        while ((line = await reader.ReadLineAsync()) != null)
        {
            if (string.IsNullOrWhiteSpace(line)) continue;
            var columns = line.Split(',');

            if (columns.Length < 14) continue;  // Skip invalid lines

            records.Add(new DataModel
            {
                Email_id = columns[0],
                Name = columns[1],
                Country = columns[2],
                State = columns[3],
                City = columns[4],
                Telephone_number = (int)Int64.Parse(columns[5]),
                Address_line_1 = columns[6],
                Address_line_2 = columns[7],
                Date_of_birth = columns[8],
                Gross_salary_FY2019_20 = (int)Int64.Parse(columns[9]),
                Gross_salary_FY2020_21 = (int)Int64.Parse(columns[10]),
                Gross_salary_FY2021_22 = (int)Int64.Parse(columns[11]),
                Gross_salary_FY2022_23 = (int)Int64.Parse(columns[12]),
                Gross_salary_FY2023_24 = (int)Int64.Parse(columns[13])
            });

            if (records.Count >= 10000)  // Process in batches
            {
                await ConvertToQuery(records);
                records.Clear();
            }
        }

        if (records.Count > 0)
        {
            await ConvertToQuery(records);
        }

    }


    private async Task ConvertToQuery(List<DataModel> records)
    {

        var queryBuilder = new StringBuilder();
        queryBuilder.Append("INSERT INTO usertest2 (Email_id, Name, Country, State, City, Telephone_number, Address_line_1, Address_line_2, Date_of_birth, Gross_salary_FY2019_20, Gross_salary_FY2020_21, Gross_salary_FY2021_22, Gross_salary_FY2022_23, Gross_salary_FY2023_24) VALUES ");

        bool isFirstValueEntry = true;

        foreach (var record in records)
        {
            if (!isFirstValueEntry)
            {
                queryBuilder.Append(',');
            }

            queryBuilder.Append($"(" +
                $"'{MySqlHelper.EscapeString(record.Email_id)}', " +
                $"'{MySqlHelper.EscapeString(record.Name)}', " +
                $"'{MySqlHelper.EscapeString(record.Country)}', " +
                $"'{MySqlHelper.EscapeString(record.State)}', " +
                $"'{MySqlHelper.EscapeString(record.City)}', " +
                $"{record.Telephone_number}, " +
                $"'{MySqlHelper.EscapeString(record.Address_line_1)}', " +
                $"'{MySqlHelper.EscapeString(record.Address_line_2)}', " +
                $"'{MySqlHelper.EscapeString(record.Date_of_birth)}', " +
                $"{record.Gross_salary_FY2019_20}, " +
                $"{record.Gross_salary_FY2020_21}, " +
                $"{record.Gross_salary_FY2021_22}, " +
                $"{record.Gross_salary_FY2022_23}, " +
                $"{record.Gross_salary_FY2023_24})");

            isFirstValueEntry = false;
        }

        queryBuilder.Append(';');

        await publisher.SendChunk(queryBuilder.ToString());
    }


    [HttpPost("uploadAndSplitStream1")]
    public async Task<IActionResult> UploadAndSplitStream1()
    {
        Console.WriteLine("------------------------------------------");
        Console.WriteLine("------------------------------------------");
        Console.WriteLine("------------------------------------------");
        Console.WriteLine("------------------------------------------");
        Console.WriteLine("------------------------------------------");
        Console.WriteLine("------------------------------------------");
        Console.WriteLine(((DateTimeOffset)DateTime.UtcNow).ToUnixTimeMilliseconds());

        Stopwatch stopwatch = new Stopwatch();
        stopwatch.Start();

        int numberOfParts = 20;
        var partLists = new List<string>[numberOfParts];
        var processTasks = new Task[numberOfParts];

        for (int i = 0; i < numberOfParts; i++)
        {
            partLists[i] = new List<string>();
        }

        using var reader = new StreamReader(Request.Body);
        int currentPart = 0;
        long totalLinesRead = 0;
        long linesPerPart = 5000;
        await reader.ReadLineAsync();
        await reader.ReadLineAsync();
        await reader.ReadLineAsync();
        await reader.ReadLineAsync();
        await reader.ReadLineAsync();

        string? line;
        while ((line = await reader.ReadLineAsync()) != null)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            partLists[currentPart].Add(line);
            totalLinesRead++;

            if (totalLinesRead >= linesPerPart && currentPart < numberOfParts - 1)
            {
                int partIndex = currentPart;
                processTasks[partIndex] = Task.Run(() => ProcessList(partLists[partIndex]));

                currentPart++;
                totalLinesRead = 0;
            }
        }

        // Handle the last part
        processTasks[currentPart] = Task.Run(() => ProcessList(partLists[currentPart]));

        // Wait for all processing tasks to complete
        await Task.WhenAll(processTasks.Where(t => t != null));
        stopwatch.Stop();
        TimeSpan elapsed = stopwatch.Elapsed;
        Console.WriteLine($"Total Elapsed time after Restart: {elapsed}");

        return Ok("Data uploaded, split, and processed successfully.");
    }

    private async Task ProcessList(List<string> lines)
    {
        var records = new List<DataModel>();

        foreach (var line in lines)
        {
            var columns = line.Split(',');

            if (columns.Length < 14) continue;  // Skip invalid lines

            records.Add(new DataModel
            {
                Email_id = columns[0],
                Name = columns[1],
                Country = columns[2],
                State = columns[3],
                City = columns[4],
                Telephone_number = (int)Int64.Parse(columns[5]),
                Address_line_1 = columns[6],
                Address_line_2 = columns[7],
                Date_of_birth = columns[8],
                Gross_salary_FY2019_20 = (int)Int64.Parse(columns[9]),
                Gross_salary_FY2020_21 = (int)Int64.Parse(columns[10]),
                Gross_salary_FY2021_22 = (int)Int64.Parse(columns[11]),
                Gross_salary_FY2022_23 = (int)Int64.Parse(columns[12]),
                Gross_salary_FY2023_24 = (int)Int64.Parse(columns[13])
            });

            if (records.Count >= 10000)  // Process in batches
            {
                await ConvertToQuery(records);
                records.Clear();
            }
        }

        if (records.Count > 0)
        {
            await ConvertToQuery(records);
        }
    }




    // [HttpPost("uploadAndSplitStream2")]
    // public async Task<IActionResult> UploadAndSplitStream2()
    // {
    //     Console.WriteLine("----------------------------------------Time pass--");
    //     Stopwatch stopwatch = new Stopwatch();
    //     stopwatch.Start();


    //     string uploadPath = Path.Combine("C:\\ProgramData\\MySQL\\MySQL Server 8.0\\Uploads\\", "uploads");
    //     Directory.CreateDirectory(uploadPath);
    //     Console.WriteLine(uploadPath);

    //     int numberOfParts = 4;
    //     var partFiles = new string[numberOfParts];
    //     var partStreams = new StreamWriter[numberOfParts];
    //     var processTasks = new Task[numberOfParts];

    //     // Initialize part files and streams
    //     for (int i = 0; i < numberOfParts; i++)
    //     {
    //         partFiles[i] = Path.Combine(uploadPath, $"part_{i}.csv");
    //         partStreams[i] = new StreamWriter(partFiles[i]);
    //     }

    //     using var reader = new StreamReader(Request.Body);
    //     int currentPart = 0;
    //     long totalLinesRead = 0;
    //     long linesPerPart = 25000;

    //     // Skip the first few lines (e.g., headers)
    //     Console.WriteLine(await reader.ReadLineAsync());  // Adjust the number of skipped lines based on your file
    //     Console.WriteLine(await reader.ReadLineAsync());
    //     Console.WriteLine(await reader.ReadLineAsync());
    //     Console.WriteLine(await reader.ReadLineAsync());
    //     Console.WriteLine(await reader.ReadLineAsync());

    //     string? line;
    //     while ((line = await reader.ReadLineAsync()) != null)
    //     {
    //         if (string.IsNullOrWhiteSpace(line))
    //         {
    //             continue;
    //         }
    //         var columns = line.Split(',');

    //         if (columns.Length < 14) continue;  // Skip invalid lines

    //         await partStreams[currentPart].WriteLineAsync(line);
    //         totalLinesRead++;

    //         // Switch to next part if linesPerPart is reached
    //         if (totalLinesRead >= linesPerPart && currentPart < numberOfParts - 1)
    //         {
    //             await partStreams[currentPart].FlushAsync();
    //             partStreams[currentPart].Dispose();

    //             int partIndex = currentPart;
    //             processTasks[partIndex] = Task.Run(() => ProcessFile2(partFiles[partIndex]));

    //             currentPart++;
    //             totalLinesRead = 0;
    //         }
    //     }

    //     // Handle the last part
    //     await partStreams[currentPart].FlushAsync();
    //     partStreams[currentPart].Dispose();
    //     processTasks[currentPart] = Task.Run(() => ProcessFile2(partFiles[currentPart]));

    //     // Wait for all processing tasks to complete
    //     await Task.WhenAll(processTasks.Where(t => t != null));
    //     stopwatch.Stop();
    //     TimeSpan elapsed = stopwatch.Elapsed;
    //     Console.WriteLine($"Total Elapsed time after Restart: {elapsed}");

    //     return Ok("File uploaded, split, and processed successfully.");

    // }

    // private async Task ProcessFile2(string filePath)
    // {
    //     try
    //     {
    //         // Simulate processing using LOAD DATA INFILE
    //         var connectionString = "server=localhost;port=3306;user=root;password=2003;database=training;pooling=true";  // Replace with your actual connection string

    //         using (var connection = new MySqlConnection(connectionString))
    //         {
    //             connection.Open();

    //             var loadQuery = $@"
    //             LOAD DATA INFILE '{filePath.Replace("\\", "/")}'
    //             INTO TABLE usertest2
    //             FIELDS TERMINATED BY ',' 
    //             LINES TERMINATED BY '\r\n'; ";

    //             using (var command = new MySqlCommand(loadQuery, connection))
    //             {
    //                 Stopwatch stopwatch = new Stopwatch();
    //                 stopwatch.Start();

    //                 command.ExecuteNonQuery();
    //                 TimeSpan elapsed = stopwatch.Elapsed;
    //                 Console.WriteLine($"Total time for execution queue {elapsed}");
    //             }
    //         }
    //     }
    //     catch (Exception ex)
    //     {
    //         Console.WriteLine($"Error processing file {filePath}: {ex.Message}");
    //     }
    // }



    // [HttpPost("uploadAndCreateTable")]
    // public async Task<IActionResult> UploadAndCreateTable()
    // {
    //     try
    //     {
    //         using var reader = new StreamReader(Request.Body);
    //         Console.WriteLine(await reader.ReadLineAsync());  // Adjust the number of skipped lines based on your file
    //         Console.WriteLine(await reader.ReadLineAsync());
    //         Console.WriteLine(await reader.ReadLineAsync());
    //         Console.WriteLine(await reader.ReadLineAsync());
    //         // Read the header
    //         string? header = await reader.ReadLineAsync();
    //         if (string.IsNullOrEmpty(header))
    //         {
    //             return BadRequest("CSV file is empty or invalid.");
    //         }



    //         // Send the header back to the client
    //         return Ok(new { columns = header.Split(',') });
    //     }
    //     catch (Exception ex)
    //     {
    //         return StatusCode(500, $"Internal server error: {ex.Message}");
    //     }
    // }


    // [HttpPost("createTable")]


    // public IActionResult CreateTable([FromBody] TableCreationRequest request)
    // {
    //     try
    //     {
    //         var connectionString = "server=localhost;port=3306;user=root;password=2003;database=training;pooling=true";

    //         using (var connection = new MySqlConnection(connectionString))
    //         {
    //             connection.Open();

    //             var createTableQuery = BuildCreateTableQuery(request);
    //             Console.WriteLine(createTableQuery);

    //             using (var command = new MySqlCommand(createTableQuery, connection))
    //             {
    //                 command.ExecuteNonQuery();
    //             }
    //         }

    //         return Ok("Table created successfully.");
    //     }
    //     catch (Exception ex)
    //     {
    //         return StatusCode(500, $"Internal server error: {ex.Message}");
    //     }
    // }


    [HttpPost("uploadAndCreateTable")]
    public async Task<IActionResult> UploadAndCreateTable()
    {
        Console.WriteLine(1234);
        try
        {
            var file = Request.Form.Files[0];
            var tempPath = Path.GetTempFileName();
            using (var stream = new FileStream(tempPath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }

            // Read the header
            string? header;
            using (var reader = new StreamReader(tempPath))
            {
                header = await reader.ReadLineAsync();
            }

            if (string.IsNullOrEmpty(header))
            {
                return BadRequest("CSV file is empty or invalid.");
            }

            Console.WriteLine("saved the file");
            // Send the header back to the client
            return Ok(new { columns = header.Split(','), tempFilePath = tempPath });
        }
        catch (Exception ex)
        {
            return StatusCode(500, $"Internal server error: {ex.Message}");
        }
    }

    [HttpPost("createTableAndUpload1")]
    public async Task<IActionResult> CreateTableAndUpload1([FromBody] TableCreationRequest request)
    {
        Console.WriteLine(123);
        try
        {
            var connectionString = "server=localhost;port=3306;user=root;password=2003;database=training;pooling=true";
            Console.WriteLine(connectionString);
            using (var connection = new MySqlConnection(connectionString))
            {
                connection.Open();

                var createTableQuery = BuildCreateTableQuery(request);
                Console.WriteLine(createTableQuery);
                using (var command = new MySqlCommand(createTableQuery, connection))
                {
                    command.ExecuteNonQuery();
                }
            }

            // Now proceed with file upload and processing
            await UploadAndSplitStream(request.TableName, request.Columns.Count, request.TempFilePath);

            // Delete the temporary file
            System.IO.File.Delete(request.TempFilePath);

            return Ok("Table created and data uploaded successfully.");
        }
        catch (Exception ex)
        {
            return StatusCode(500, $"Internal server error: {ex.Message}");
        }
    }


    private async Task UploadAndSplitStream(string tableName, int expectedColumnCount, string tempFilePath)
    {
        string uploadPath = Path.Combine("C:\\ProgramData\\MySQL\\MySQL Server 8.0\\Uploads\\", "uploads");
        Directory.CreateDirectory(uploadPath);

        int numberOfParts = 4;
        var partFiles = new string[numberOfParts];
        var partStreams = new StreamWriter[numberOfParts];
        var processTasks = new Task[numberOfParts];
        var miscFile = Path.Combine(uploadPath, "miscellaneous.csv");
        var miscStream = new StreamWriter(miscFile);

        // Initialize part files and streams
        for (int i = 0; i < numberOfParts; i++)
        {
            partFiles[i] = Path.Combine(uploadPath, $"part_{i}.csv");
            partStreams[i] = new StreamWriter(partFiles[i]);
        }

        using var reader = new StreamReader(tempFilePath);
        int currentPart = 0;
        long totalLinesRead = 0;
        long linesPerPart = 25000;

        // Skip the header
        Console.WriteLine(await reader.ReadLineAsync());  // Adjust the number of skipped lines based on your file
        Console.WriteLine(await reader.ReadLineAsync());
        Console.WriteLine(await reader.ReadLineAsync());
        Console.WriteLine(await reader.ReadLineAsync());
        await reader.ReadLineAsync();

        string? line;
        while ((line = await reader.ReadLineAsync()) != null)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }
            var columns = line.Split(',');

            if (columns.Length < expectedColumnCount)
            {
                // Pad with null values
                line = string.Join(",", columns.Concat(Enumerable.Repeat("NULL", expectedColumnCount - columns.Length)));
                await partStreams[currentPart].WriteLineAsync(line);
            }
            else if (columns.Length > expectedColumnCount)
            {
                // Write to miscellaneous file
                await miscStream.WriteLineAsync(line);
            }
            else
            {
                await partStreams[currentPart].WriteLineAsync(line);
            }

            totalLinesRead++;

            // Switch to next part if linesPerPart is reached
            if (totalLinesRead >= linesPerPart && currentPart < numberOfParts - 1)
            {
                await partStreams[currentPart].FlushAsync();
                partStreams[currentPart].Dispose();

                int partIndex = currentPart;
                processTasks[partIndex] = Task.Run(() => ProcessFile(partFiles[partIndex], tableName));

                currentPart++;
                totalLinesRead = 0;
            }
        }

        // Handle the last part
        await partStreams[currentPart].FlushAsync();
        partStreams[currentPart].Dispose();
        processTasks[currentPart] = Task.Run(() => ProcessFile(partFiles[currentPart], tableName));

        // Close miscellaneous file
        await miscStream.FlushAsync();
        miscStream.Dispose();

        // Wait for all processing tasks to complete
        await Task.WhenAll(processTasks.Where(t => t != null));

        // Handle miscellaneous rows
        if (new FileInfo(miscFile).Length > 0)
        {
            // Notify the user about miscellaneous rows and ask for action
            // This could be done through a separate API endpoint or WebSocket
            Console.WriteLine("Miscellaneous rows found. User action required.");
        }
    }



    private async Task ProcessFile(string filePath, string tableName)
    {
        try
        {
            var connectionString = "server=localhost;port=3306;user=root;password=2003;database=training;pooling=true";

            using (var connection = new MySqlConnection(connectionString))
            {
                connection.Open();

                var loadQuery = $@"
            LOAD DATA INFILE '{filePath.Replace("\\", "/")}'
            INTO TABLE `{tableName}`
            FIELDS TERMINATED BY ',' 
            LINES TERMINATED BY '\r\n'; ";

                using (var command = new MySqlCommand(loadQuery, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing file {filePath}: {ex.Message}");
        }
    }

    [HttpPost("checkMiscellaneousRows")]
    public IActionResult CheckMiscellaneousRows([FromBody] string tableName)
    {
        string miscFilePath = Path.Combine("C:\\ProgramData\\MySQL\\MySQL Server 8.0\\Uploads\\", "uploads", "miscellaneous.csv");
        bool hasMiscellaneousRows = System.IO.File.Exists(miscFilePath) && new FileInfo(miscFilePath).Length > 0;

        return Ok(new { hasMiscellaneousRows });
    }

    [HttpPost("handleMiscellaneousRows")]
    public IActionResult HandleMiscellaneousRows([FromBody] MiscRowsAction action)
    {
        Console.WriteLine("Hello");
        switch (action.Action)
        {
            case "delete":
                // Delete the miscellaneous file
                System.IO.File.Delete(Path.Combine("C:\\ProgramData\\MySQL\\MySQL Server 8.0\\Uploads\\", "uploads", "miscellaneous.csv"));
                break;
            case "alterTable":
                // Alter the table to add new columns
                AlterTableAddColumns(action.TableName, action.NewColumns);
                break;
            case "truncate":
                // Upload only up to specified columns
                UploadTruncatedMiscRows(action.TableName, action.ColumnCount);
                break;
            default:
                return BadRequest("Invalid action specified");
        }

        return Ok("Miscellaneous rows handled successfully");
    }

    private void AlterTableAddColumns(string tableName, List<string> newColumns)
    {
        var connectionString = "server=localhost;port=3306;user=root;password=2003;database=training;pooling=true";

        using (var connection = new MySqlConnection(connectionString))
        {
            connection.Open();

            foreach (var column in newColumns)
            {
                var alterQuery = $"ALTER TABLE `{tableName}` ADD COLUMN `{NameCorrection(column)}` VARCHAR(255)";
                using (var command = new MySqlCommand(alterQuery, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }
    }

    private void UploadTruncatedMiscRows(string tableName, int columnCount)
    {
        var miscFile = Path.Combine("C:\\ProgramData\\MySQL\\MySQL Server 8.0\\Uploads\\", "uploads", "miscellaneous.csv");
        var truncatedFile = Path.Combine("C:\\ProgramData\\MySQL\\MySQL Server 8.0\\Uploads\\", "uploads", "truncated.csv");

        using (var reader = new StreamReader(miscFile))
        using (var writer = new StreamWriter(truncatedFile))
        {
            string? line;
            while ((line = reader.ReadLine()) != null)
            {
                var columns = line.Split(',');
                var truncatedLine = string.Join(",", columns.Take(columnCount));
                writer.WriteLine(truncatedLine);
            }
        }

        ProcessFile(truncatedFile, tableName).Wait();
        System.IO.File.Delete(truncatedFile);
    }

    public class MiscRowsAction
    {
        public string? Action { get; set; }
        public string? TableName { get; set; }
        public List<string>? NewColumns { get; set; }
        public int ColumnCount { get; set; }
    }


    private string BuildCreateTableQuery(TableCreationRequest request)
    {
        // Sanitize table name
        string CorrectTableName = NameCorrection(request.TableName);
        var query = new StringBuilder($"CREATE TABLE `{CorrectTableName}` (");

        bool hasPrimaryKey = false;

        for (int i = 0; i < request.Columns.Count; i++)
        {
            var column = request.Columns[i];

            // Sanitize column name
            string CorrectColumnName = NameCorrection(column.Name);

            query.Append($"`{CorrectColumnName}` {column.Type}");

            if (!column.AllowNull)
            {
                query.Append(" NOT NULL");
            }

            if (column.IsPrimaryKey)
            {
                if (hasPrimaryKey)
                {
                    throw new InvalidOperationException("Only one primary key is allowed.");
                }
                query.Append(" PRIMARY KEY");
                hasPrimaryKey = true;
            }

            if (i < request.Columns.Count - 1)
            {
                query.Append(", ");
            }
        }

        query.Append(")");
        return query.ToString();
    }

    private string NameCorrection(string identifier)
    {
        // Replace spaces with underscores
        identifier = identifier.Replace(" ", "_");

        // Remove any characters that are not alphanumeric or underscores
        identifier = new string(identifier.Where(c => char.IsLetterOrDigit(c) || c == '_').ToArray());

        // Ensure the identifier doesn't start with a number
        if (identifier.Length > 0 && char.IsDigit(identifier[0]))
        {
            identifier = "_" + identifier;
        }

        return identifier;
    }


   public class TableCreationRequest
{
    public string? TableName { get; set; }
    public List<ColumnDefinition>? Columns { get; set; }
    public string? TempFilePath { get; set; }
}

    public class ColumnDefinition
    {
        public string? Name { get; set; }
        public string? Type { get; set; }
        public bool AllowNull { get; set; }
        public bool IsPrimaryKey { get; set; }
    }


}