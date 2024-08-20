using ClosedXML.Excel;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Data.SqlClient;

public class ExcelToSqlService
{
    private readonly string _connectionString;

    public ExcelToSqlService(string connectionString)
    {
        _connectionString = connectionString;
    }

    public async Task<string> ReadExcelFileAsync(Stream excelStream, string tableName)
    {
        if (excelStream == null || excelStream.Length == 0)
        {
            throw new ArgumentException("Stream is null or empty.");
        }

        try
        {
            using var workbook = new XLWorkbook(excelStream);
            var worksheet = workbook.Worksheet(1); // Get the first worksheet

            // Validate if worksheet has data
            if (worksheet.RowCount() < 2)
            {
                throw new InvalidOperationException("Worksheet does not contain sufficient data.");
            }

            var dataTable = new DataTable();

            // Add columns to DataTable
            foreach (var cell in worksheet.Row(1).Cells())
            {
                dataTable.Columns.Add(cell.Value.ToString());
            }

            // Add rows to DataTable
            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                var dataRow = dataTable.NewRow();
                for (int col = 1; col <= dataTable.Columns.Count; col++)
                {
                    dataRow[col - 1] = row.Cell(col).Value;
                }
                dataTable.Rows.Add(dataRow);
            }

            // Create the table in SQL Server
            await CreateTableAsync(tableName, dataTable);

            return "File processed and table created successfully.";
        }
        catch (IOException ioEx)
        {
            // Log IO exception
            Console.WriteLine($"IO Error reading Excel file: {ioEx.Message}");
            throw new InvalidOperationException("Error reading the Excel file due to IO issue.", ioEx);
        }
        catch (Exception ex)
        {
            // Log other exceptions
            Console.WriteLine($"Error reading Excel file: {ex.Message}");
            throw;
        }
    }

    private async Task CreateTableAsync(string tableName, DataTable dataTable)
    {
        using var connection = new SqlConnection(_connectionString);
        await connection.OpenAsync();

        var createTableQuery = GenerateCreateTableQuery(tableName, dataTable);
        using var createTableCommand = new SqlCommand(createTableQuery, connection);
        await createTableCommand.ExecuteNonQueryAsync();

        var insertDataQuery = GenerateInsertDataQuery(tableName, dataTable);
        using var insertDataCommand = new SqlCommand(insertDataQuery, connection);
        foreach (DataRow row in dataTable.Rows)
        {
            foreach (var item in row.ItemArray)
            {
                insertDataCommand.Parameters.AddWithValue("@value", item);
            }
            await insertDataCommand.ExecuteNonQueryAsync();
            insertDataCommand.Parameters.Clear();
        }
    }

    private string GenerateCreateTableQuery(string tableName, DataTable dataTable)
    {
        var columns = string.Join(", ", dataTable.Columns
            .Cast<DataColumn>()
            .Select(col => $"[{col.ColumnName}] NVARCHAR(MAX)"));
        return $@"
            IF OBJECT_ID('{tableName}', 'U') IS NOT NULL
                DROP TABLE [{tableName}];
            CREATE TABLE [{tableName}] (
                {columns}
            );";
    }

    private string GenerateInsertDataQuery(string tableName, DataTable dataTable)
    {
        var columns = string.Join(", ", dataTable.Columns.Cast<DataColumn>().Select(col => $"[{col.ColumnName}]"));
        var values = string.Join(", ", dataTable.Columns.Cast<DataColumn>().Select(col => "@value"));
        return $@"
            INSERT INTO [{tableName}] ({columns})
            VALUES ({values});";
    }
}
