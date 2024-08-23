using RabbitMQ.Client;
using System;
using System.Text;
using System.Threading.Tasks;

public class RabbitMQPublisher : IDisposable
{
    private readonly IConnection _connection;
    private readonly IModel _channel;
    private int _currentQueueIndex = 0;
    private const int QueueCount = 5;

    public RabbitMQPublisher(string hostName)
    {
        var factory = new ConnectionFactory() { HostName = hostName };
        _connection = factory.CreateConnection();
        _channel = _connection.CreateModel();

        // Declare queues
        for (int i = 0; i < QueueCount; i++)
        {
            _channel.QueueDeclare(queue: $"queue{i}", durable: false, exclusive: false, autoDelete: false, arguments: null);
        }
    }

    public async Task SendChunk(string chunk)
    {
        var body = Encoding.UTF8.GetBytes(chunk);

        _channel.BasicPublish(exchange: string.Empty,
                              routingKey: $"queue{_currentQueueIndex}",
                              basicProperties: null,
                              body: body);

        Console.WriteLine($"Sent chunk to queue{_currentQueueIndex}");

        _currentQueueIndex = (_currentQueueIndex + 1) % QueueCount;

        // Simulate some processing time
        await Task.Delay(100);
    }

    public void Dispose()
    {
        _channel?.Dispose();
        _connection?.Dispose();
    }
}