using CoedicionExcel.Models;
using System.Reflection.Emit;
using Microsoft.EntityFrameworkCore;

namespace CoedicionExcel.Data
{
    public class AppDbContext : DbContext
    {
        public AppDbContext(DbContextOptions<AppDbContext> options) : base(options)
        {
        }

        public DbSet<DocumentoExcel> DocumentosExcel { get; set; }
        public DbSet<ColumnaExcel> ColumnasExcel { get; set; }
        public DbSet<FilaExcel> FilasExcel { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            modelBuilder.Entity<DocumentoExcel>(entity =>
            {
                entity.ToTable("DocumentosExcel");
                entity.HasKey(e => e.DocumentoId);
            });

            modelBuilder.Entity<ColumnaExcel>(entity =>
            {
                entity.ToTable("ColumnasExcel");
                entity.HasKey(e => e.ColumnaId);

                entity.HasOne(e => e.Documento)
                      .WithMany(d => d.Columnas)
                      .HasForeignKey(e => e.DocumentoId)
                      .OnDelete(DeleteBehavior.Cascade);
            });

            modelBuilder.Entity<FilaExcel>(entity =>
            {
                entity.ToTable("FilasExcel");
                entity.HasKey(e => e.FilaId);

                entity.HasOne(e => e.Documento)
                      .WithMany(d => d.Filas)
                      .HasForeignKey(e => e.DocumentoId)
                      .OnDelete(DeleteBehavior.Cascade);
            });
        }
    }
}
