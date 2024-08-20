using Microsoft.EntityFrameworkCore;

namespace DynamicExcelToSQL.Data
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
            : base(options)
        {
        }

        // Optionally add DbSet properties for your tables here if you have predefined tables
    }
}
