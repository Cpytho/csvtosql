using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;

public class RabbitMQConsumer : IDisposable
{
    private readonly IConnection _connection;
    private readonly IModel _channel;
    private readonly string _dbConnString;
    private const int MaxParallelTasks = 10; // Adjust based on your needs
    private static SemaphoreSlim _semaphore = new SemaphoreSlim(10); // Adjust based on your database's capacity

    public RabbitMQConsumer(string hostName, string dbConnString)
    {
        var factory = new ConnectionFactory() { HostName = hostName };
        _connection = factory.CreateConnection();
        _channel = _connection.CreateModel();
        _dbConnString = dbConnString;

        // Set prefetch count
        _channel.BasicQos(prefetchSize: 0, prefetchCount: 10, global: false);
    }

    public async Task StartListeningToAllQueues(CancellationToken cancellationToken)
    {
        var tasks = new Task[10];
        for (int i = 0; i < 10; i++)
        {
            int queueNumber = i;
            tasks[i] = StartListeningToQueue(queueNumber, cancellationToken);
        }
        await Task.WhenAll(tasks);
    }

    private async Task StartListeningToQueue(int queueNumber, CancellationToken cancellationToken)
    {
        _channel.QueueDeclare(queue: $"queue{queueNumber}", durable: false, exclusive: false, autoDelete: false, arguments: null);

        var consumer = new EventingBasicConsumer(_channel);
        var pendingTasks = new List<Task>();
        Stopwatch stopwatch = new Stopwatch();
        stopwatch.Start();

        consumer.Received += (model, ea) =>
        {
            var body = ea.Body.ToArray();
            var message = Encoding.UTF8.GetString(body);

            var insertTask = Task.Run(async () =>
            {
                try
                {
                    await InsertToDB(message, queueNumber);
                    _channel.BasicAck(ea.DeliveryTag, false);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing message: {ex.Message}");
                    _channel.BasicNack(ea.DeliveryTag, false, true);
                }
            });

            pendingTasks.Add(insertTask);

            // Ensure tasks are executed in parallel up to a certain level
            if (pendingTasks.Count >= MaxParallelTasks)
            {
                Task.WaitAny(pendingTasks.ToArray());
                pendingTasks.RemoveAll(t => t.IsCompleted); // Clean up completed tasks
            }
        };

        string consumerTag = _channel.BasicConsume(queue: $"queue{queueNumber}", autoAck: false, consumer: consumer);

        try
        {
            await Task.Delay(Timeout.Infinite, cancellationToken);
        }
        catch (OperationCanceledException)
        {
            _channel.BasicCancel(consumerTag);
        }

        // Wait for any remaining tasks to complete after cancellation
        await Task.WhenAll(pendingTasks);
        stopwatch.Stop();
        TimeSpan elapsed = stopwatch.Elapsed;
        Console.WriteLine($"Total Elapsed time for queue{queueNumber}: {elapsed}");
    }

    private async Task InsertToDB(string query, int queueN)
    {
        await _semaphore.WaitAsync();
        Console.WriteLine($"Inserting batch query from queue{queueN}");
        try
        {
            using (var connection = new MySqlConnection(_dbConnString))
            {
                await connection.OpenAsync();

                // Use MySQL Transaction for batching and atomic operations
                using (var transaction = await connection.BeginTransactionAsync())
                {
                    try
                    {
                        using (var command = new MySqlCommand(query, connection, transaction))
                        {
                            command.CommandTimeout = 60; // Adjust as needed
                            await command.ExecuteNonQueryAsync();
                        }

                        await transaction.CommitAsync();
                        Console.WriteLine($"Executed batch query from queue{queueN}");
                    }
                    catch (Exception ex)
                    {
                        await transaction.RollbackAsync();
                        Console.WriteLine($"Error executing batch query for queue{queueN}: {ex.Message}");
                        throw;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error in InsertToDB: {ex.Message}");
            throw;
        }
        finally
        {
            _semaphore.Release();
        }
    }

    public void Dispose()
    {
        _channel?.Dispose();
        _connection?.Dispose();
    }
}