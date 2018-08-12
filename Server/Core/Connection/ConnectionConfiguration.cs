namespace Core.Connection
{
    public class ConnectionConfiguration
    {
        public string Host { get; set; }

        public string PortNumber { get; set; }

        public bool UseProxy { get; set; }

        public string LoggerPath { get; set; }
    }
}
