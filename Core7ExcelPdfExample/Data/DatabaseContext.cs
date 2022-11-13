using Core7ExcelPdfExample.Entities;
using Microsoft.EntityFrameworkCore;

namespace Core7ExcelPdfExample.Data;

public class DatabaseContext : DbContext
{
    public DatabaseContext(DbContextOptions<DatabaseContext> options) : base(options)
    {
    }

    public DbSet<Categories> Categories { get; set; }
}
