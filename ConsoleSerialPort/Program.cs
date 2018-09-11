using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Management;

namespace ConsoleSerialPort
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string dcDeviceName = "Prolific USB-to-Serial Comm Port";
                List<string> portFullName = SerialPort.GetPortNames().ToList();
                int countPort = 0;

                //{4d36e978-e325-11ce-bfc1-08002be10318}為設備類別port（端口（COM&LPT））的GUID
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2",
                    "SELECT * FROM Win32_PnPEntity WHERE ClassGuid=\"{4d36e978-e325-11ce-bfc1-08002be10318}\"");
                Console.WriteLine("No.\t Port \t Full Name");
                foreach (ManagementObject queryObj in searcher.Get())
                {
                    countPort++;
                    string fullName = queryObj.GetPropertyValue("Name").ToString();
                    string[] aryName = fullName.Split(new char[2] { '(', ')' });
                    Console.WriteLine("{0}.\t {1} \t {2}",countPort,aryName[1],aryName[0]);
                }

            }
            catch (Exception e)
            {
                Console.WriteLine("An error occurred : " + e.Message);
            }
            finally
            {
                Console.ReadLine();
            }
        }
    }
}
