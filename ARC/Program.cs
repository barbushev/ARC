using System;
using System.Management;
using System.IO.Ports;

namespace ARC
{
    class Program
    {
        const string   devDescriptor = "Arduino";
        const int      devBaudRate = 9600;
        const int      devDataBits = 8;
        const int      devReadTimeOut = 1000;  //ms
        const Parity   devParity = Parity.None;
        const StopBits devStopBits = StopBits.One;
        
        public enum ReturnCode
        {
            OK,
            ARC_WRONG_ARGS,
            ARC_INVALID_COMMAND,
            ARC_NOT_FOUND,
            ARC_COMMUNICATION_FAIL
        }

        static int Main(string[] args)
        {
            ReturnCode ret = ReturnCode.OK;
            string arcComPort = null;

            if (args == null || args.Length != 1)
            {
                Console.WriteLine("Usage: arc.exe cmd");
                Console.WriteLine("cmd syntax L#, where (L)evel is H or L, and # is pin # 2 to 19");
                Console.WriteLine("Example \"arc.exe L2\" will set pin 2 Low, and \"arc.exe H2\" will set pin 2 high.");
                Console.WriteLine("Note: pins 0 and 1 are reserved for communication.");
                ret = ReturnCode.ARC_WRONG_ARGS;
            }

            if (ret == ReturnCode.OK)            
                ret = CommandIsValid(args[0]);

            if (ret == ReturnCode.OK)
            {
                arcComPort = FindDevicePort();
                if (arcComPort == null)
                    ret = ReturnCode.ARC_NOT_FOUND;
            }

            if (ret == ReturnCode.OK)
            {
                ret = ExecuteCommand(arcComPort, args[0]);
            }

            Console.WriteLine($"{ret}");
            return (int)ret;
        }

        private static ReturnCode CommandIsValid(string cmd)
        {
            if (cmd.Length != 2 && cmd.Length != 3)            
                return ReturnCode.ARC_INVALID_COMMAND;

            if (cmd[0] != 'L' && cmd[0] != 'H')
                return ReturnCode.ARC_INVALID_COMMAND;                           

            if (!int.TryParse(cmd.Substring(1), out int pinNum))
                return ReturnCode.ARC_INVALID_COMMAND;            
            else if (pinNum <= 1 || pinNum >= 20)
                return ReturnCode.ARC_INVALID_COMMAND;
            
            return ReturnCode.OK;
        }


        private static ReturnCode ExecuteCommand(string comPort, string cmd)
        {
            SerialPort arcSerial = new SerialPort(comPort, devBaudRate, devParity, devDataBits, devStopBits);
            arcSerial.ReadTimeout = devReadTimeOut;
            arcSerial.NewLine = "\n";
             
            try
            {
                arcSerial.Open();
                arcSerial.WriteLine(comPort);
                string resp = arcSerial.ReadLine();
                Console.WriteLine(resp);
                arcSerial.Close();
            }
            catch
            { 
                return ReturnCode.ARC_COMMUNICATION_FAIL;
            }

            return ReturnCode.OK;
        }

        private static string FindDevicePort()
        {
            ManagementScope connectionScope = new ManagementScope();
            SelectQuery serialQuery = new SelectQuery("SELECT * FROM Win32_SerialPort");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(connectionScope, serialQuery);

            try
            {
                foreach (ManagementObject item in searcher.Get())
                {
                    string desc = item["Description"].ToString();
                    string deviceId = item["DeviceID"].ToString();

                    if (desc.Contains(devDescriptor))
                    {
                        return deviceId;
                    }
                }
            }
            catch { }

            return null;
        }
    }
}
