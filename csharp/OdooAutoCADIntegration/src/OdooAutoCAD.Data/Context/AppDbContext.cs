// OdooAutoCAD.Data/Context/AppDbContext.cs
// Entity Framework Core DbContext - equivalent to Python SQLAlchemy models

using Microsoft.EntityFrameworkCore;
using OdooAutoCAD.Data.Entities;

namespace OdooAutoCAD.Data.Context;

/// <summary>
/// Application database context using SQLite.
/// Equivalent to Python's SQLAlchemy database setup.
/// </summary>
public class AppDbContext : DbContext
{
    public AppDbContext(DbContextOptions<AppDbContext> options) : base(options)
    {
    }

    public DbSet<ServerConfig> ServerConfigs { get; set; } = null!;
    public DbSet<ProductMapping> ProductMappings { get; set; } = null!;
    public DbSet<BOQCache> BOQCache { get; set; } = null!;
    public DbSet<SyncLog> SyncLogs { get; set; } = null!;
    public DbSet<UserPreference> UserPreferences { get; set; } = null!;

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        base.OnModelCreating(modelBuilder);

        // ServerConfig
        modelBuilder.Entity<ServerConfig>(entity =>
        {
            entity.HasKey(e => e.Id);
            entity.Property(e => e.ConfigKey).IsRequired().HasMaxLength(100);
            entity.Property(e => e.ConfigValue).HasMaxLength(1000);
            entity.HasIndex(e => e.ConfigKey).IsUnique();
        });

        // ProductMapping
        modelBuilder.Entity<ProductMapping>(entity =>
        {
            entity.HasKey(e => e.Id);
            entity.Property(e => e.AutoCADName).IsRequired().HasMaxLength(200);
            entity.Property(e => e.OdooProductCode).HasMaxLength(100);
            entity.HasIndex(e => e.AutoCADName);
        });

        // BOQCache
        modelBuilder.Entity<BOQCache>(entity =>
        {
            entity.HasKey(e => e.Id);
            entity.Property(e => e.ProjectId).IsRequired();
            entity.Property(e => e.ProductName).IsRequired().HasMaxLength(200);
            entity.HasIndex(e => e.ProjectId);
        });

        // SyncLog
        modelBuilder.Entity<SyncLog>(entity =>
        {
            entity.HasKey(e => e.Id);
            entity.Property(e => e.SyncType).IsRequired().HasMaxLength(50);
            entity.HasIndex(e => e.SyncTime);
        });

        // UserPreference
        modelBuilder.Entity<UserPreference>(entity =>
        {
            entity.HasKey(e => e.Id);
            entity.Property(e => e.PreferenceKey).IsRequired().HasMaxLength(100);
            entity.HasIndex(e => e.PreferenceKey).IsUnique();
        });
    }
}

/// <summary>
/// Factory for creating DbContext instances.
/// </summary>
public class AppDbContextFactory
{
    private readonly string _connectionString;

    public AppDbContextFactory(string connectionString = "Data Source=database.db")
    {
        _connectionString = connectionString;
    }

    public AppDbContext CreateDbContext()
    {
        var optionsBuilder = new DbContextOptionsBuilder<AppDbContext>();
        optionsBuilder.UseSqlite(_connectionString);
        return new AppDbContext(optionsBuilder.Options);
    }
}
