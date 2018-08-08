namespace Server
{
    using System;
    using System.Configuration;

    public class EnvironmentConfiguration
    {
        public string Host { get; set; }

        public string PortNumber { get; set; }

        public bool UseProxy { get; set; }

        public static EnvironmentConfiguration Load()
        {
            return new EnvironmentConfiguration
            {
                Host = ConfigurationManager.AppSettings["HostAddress"] as string ?? "127.0.0.1",
                PortNumber = ConfigurationManager.AppSettings["PortNumber"] as string ?? "100",
                UseProxy =  Convert.ToBoolean(ConfigurationManager.AppSettings["UseProxy"])
            };
        }
    }
}
