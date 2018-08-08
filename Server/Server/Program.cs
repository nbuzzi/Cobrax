namespace Server
{
    using System;
    using System.Threading;
    using Core.Connection;

    public static class Program
    {
        private static readonly ConnectionConfiguration _connectionConfiguration;
        private static readonly EnvironmentConfiguration _environmentConfiguration;
        private static readonly Connection _connection;

        static Program()
        {
            _environmentConfiguration = EnvironmentConfiguration.Load();

            _connectionConfiguration = new ConnectionConfiguration
            {
                Host = _environmentConfiguration.Host,
                PortNumber = _environmentConfiguration.PortNumber,
                UseProxy = _environmentConfiguration.UseProxy
            };

            _connection = new Connection(_connectionConfiguration);
        }

        static void Main(string[] args)
        {
            while (true)
            {
                try
                {
                    _connection.ConfigureClient();

                    while (true)
                    {
                        try
                        {
                            _connection.ReceiveInformation();
                        }
                        catch (Exception ex)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("Error: {0}", ex.Message);

                            // Wait 3 seconds previous to reconnect
                            Thread.Sleep(3000);
                            break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Error: {0}", ex.Message);

                    // Wait 3 seconds previous to reconnect
                    Thread.Sleep(3000);
                }
            }
        }
    }
}
