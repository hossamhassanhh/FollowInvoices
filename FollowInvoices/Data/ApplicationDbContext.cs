using Microsoft.EntityFrameworkCore;
using FollowInvoices.Models;

namespace FollowInvoices.Data
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options) : base(options)
        {
        }

        public DbSet<IsoDetails> IsoDetails { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlServer("Server=DESKTOP-BSIH5O5\\SQLEXPRESS;Database=Test;Integrated Security=True;TrustServerCertificate=true");
            //optionsBuilder.UseSqlServer("Server=DESKTOP-BSIH5O5\\SQLEXPRESS;Database=PLANDB;Integrated Security=True;TrustServerCertificate=true");
        }
    }
}