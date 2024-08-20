using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Threading.Tasks;

[Route("api/[controller]")]
[ApiController]
public class ExcelController : ControllerBase
{
    private readonly ExcelToSqlService _excelToSqlService;

    public ExcelController(ExcelToSqlService excelToSqlService)
    {
        _excelToSqlService = excelToSqlService;
    }

    [HttpPost("upload")]
    public async Task<IActionResult> UploadExcel(IFormFile file)
    {
        if (file == null || file.Length == 0)
        {
            return BadRequest("No file uploaded.");
        }

        var tableName = Path.GetFileNameWithoutExtension(file.FileName); // Use file name as table name

        try
        {
            using var stream = file.OpenReadStream();
            var result = await _excelToSqlService.ReadExcelFileAsync(stream, tableName);
            return Ok(result);
        }
        catch (Exception ex)
        {
            return StatusCode(500, $"Internal server error: {ex.Message}");
        }
    }
}
