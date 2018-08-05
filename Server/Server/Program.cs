namespace Server
{
    using System;
    using System.Configuration;
    using System.Threading;
    using Core.Connection;

    class Program
    {
        static void Main(string[] args)
        {
            var connection = new Connection();

            var addressToConnect = ConfigurationManager.AppSettings["HostAddress"] as string ?? "127.0.0.1";
            var portNumber = ConfigurationManager.AppSettings["PortNumber"] as string ?? "100";

            while (true)
            {
                try
                {
                    connection.ConfigureClient(addressToConnect, int.Parse(portNumber));

                    while (true)
                    {
                        try
                        {
                            connection.ReceiveInformation();
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
