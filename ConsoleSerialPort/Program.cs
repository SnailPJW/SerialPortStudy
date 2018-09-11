using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Management;

namespace ConsoleSerialPort
{
    class Program
    {
        static string dcDeviceName = "Prolific USB-to-Serial Comm Port";
        static Dictionary<string, string> dictionaryPortName = new Dictionary<string, string>();
        static int countPort = 0;
        static string connPortName;
        static void Main(string[] args)
        {
            try
            {
                //{4d36e978-e325-11ce-bfc1-08002be10318}為設備類別port（端口（COM&LPT））的GUID
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2",
                    "SELECT * FROM Win32_PnPEntity WHERE ClassGuid=\"{4d36e978-e325-11ce-bfc1-08002be10318}\"");

                foreach (ManagementObject queryObj in searcher.Get())
                {
                    string fullName = queryObj.GetPropertyValue("Name").ToString();
                    string[] aryName = fullName.Split(new char[2] { '(', ')' });
                    dictionaryPortName.Add(aryName[1], aryName[0]);
                }

                ShowAllInDictionary();
                GetDCDevice();
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
        static void ShowAllInDictionary()
        {
            Console.WriteLine("No.\t Port \t Full Name");
            foreach (var element in dictionaryPortName)
            {
                countPort++;
                Console.WriteLine("{0}.\t {1} \t {2}", countPort, element.Key, element.Value);
            }
        }
        static void GetDCDevice()
        {
            foreach (var element in dictionaryPortName)
            {
                element.Value.Trim();
                Console.WriteLine("{1}\t{2}\tThe Result is : {0}", string.Equals(dcDeviceName, element.Value), dcDeviceName.Length, element.Value.Length);
                
                //if(!string.Equals(dcDeviceName, element.Value))
                //{
                //    connPortName = element.Key;
                //    break;
                //}
            }
            //Console.WriteLine("The Targeted port is : {0}",connPortName);
        }
    }
}
