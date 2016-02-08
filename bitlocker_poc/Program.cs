using System;
using System.IO;
using System.Management;
using System.Management.Instrumentation;
using System.Text;

namespace bitlocker_poc
{
    sealed class Program
    {
        private const string BITLOCKER_NS = @"\ROOT\CIMV2\Security\MicrosoftVolumeEncryption";
        private const string BITLOCKER_CLASS = "Win32_EncryptableVolume";

        public static void Main(string[] args)
        {
            var now = DateTime.Now.ToString();
            var filepath = Environment.CurrentDirectory + @"\output.txt";
            Console.WriteLine("Fetching information about BitLocker using WIM ..\r\n");
            using (var file = File.CreateText(filepath))
            {
                try
                {
                    var path = new ManagementPath();
                    path.NamespacePath = BITLOCKER_NS;
                    path.ClassName = BITLOCKER_CLASS;

                    var scope = new ManagementScope(path);
                    var options = new ObjectGetOptions();
                    var management = new ManagementClass(scope, path, options);

                    file.WriteLine(string.Format("{0}\r\n", now));
                    int counter = 1;
                    foreach (var instance in management.GetInstances())
                    {
                        var builder = new StringBuilder();
                        builder.AppendFormat("{0}.\r\n", counter);
                        builder.AppendFormat("  DeviceID:            {0}\r\n", instance["DeviceID"]);
                        builder.AppendFormat("  PersistentVolumeID:  {0}\r\n", instance["PersistentVolumeID"]);
                        builder.AppendFormat("  DriveLetter:         {0}\r\n", instance["DriveLetter"]);
                        builder.AppendFormat("  ProtectionStatus:    {0}\r\n", instance["ProtectionStatus"]);

                        Console.WriteLine(builder.ToString());
                        file.Write(builder.ToString());
                        counter++;
                    }
                    file.Flush();

                    Console.WriteLine("Output written to: " + filepath);
                }
                catch (Exception ex)
                {
                    file.WriteLine(ex.Message);
                    Console.WriteLine(ex.Message);
                    file.WriteLine(ex.StackTrace);
                    Console.WriteLine(ex.StackTrace);
                }
                finally
                {
                    Console.WriteLine("\r\nPress Enter to continue ..");
                    Console.ReadLine();
                }
            }
        }
    }
}

