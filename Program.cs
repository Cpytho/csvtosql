using RabbitMQ.Client;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System.Threading.Tasks;

var builder = WebApplication.CreateBuilder(args);

{
    // Add services to the container.
    builder.Services.AddControllers();
    builder.Services.AddEndpointsApiExplorer();
    builder.Services.AddSwaggerGen();

    // RabbitMQ Configuration
    builder.Services.AddSingleton<IConnectionFactory>(sp =>
    {
        return new ConnectionFactory { HostName = "localhost" };
    });

    // Database Configuration
    string? dbConnString = builder.Configuration.GetConnectionString("DBConnectionString");

    // RabbitMQ Services
    builder.Services.AddSingleton<RabbitMQPublisher>(sp =>
{
    var configuration = sp.GetRequiredService<IConfiguration>();
    return new RabbitMQPublisher("localhost");
});
    builder.Services.AddSingleton<RabbitMQConsumer>(sp => 
    {
        var factory = sp.GetRequiredService<IConnectionFactory>();
        return new RabbitMQConsumer("localhost", dbConnString);
    });

    builder.Services.AddCors(options =>
    {
        options.AddPolicy("AllowAllOrigins",
            builder =>
            {
                builder.AllowAnyOrigin()
                       .AllowAnyMethod()
                       .AllowAnyHeader();
            });
    });

    // Background Service for Consumer
    builder.Services.AddHostedService<RabbitMQConsumerService>();
}

var app = builder.Build();

{
    // Configure the HTTP request pipeline.
    if (app.Environment.IsDevelopment())
    {
        app.UseSwagger();
        app.UseSwaggerUI();
    }

    app.UseHttpsRedirection();
    app.UseAuthorization();
    app.UseCors("AllowAllOrigins");
    app.MapControllers();
}

app.Run();

// RabbitMQConsumerService.cs
public class RabbitMQConsumerService : BackgroundService
{
    private readonly RabbitMQConsumer _consumer;
    private readonly ILogger<RabbitMQConsumerService> _logger;

    public RabbitMQConsumerService(RabbitMQConsumer consumer, ILogger<RabbitMQConsumerService> logger)
    {
        _consumer = consumer;
        _logger = logger;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        try
        {
            await _consumer.StartListeningToAllQueues(stoppingToken);
        }
        catch (OperationCanceledException)
        {
            _logger.LogInformation("RabbitMQ consumer stopped.");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "An error occurred in the RabbitMQ consumer.");
        }
    }
}