using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

namespace LDMFormatService
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : System.Configuration.Install.Installer
    {
        public ProjectInstaller()
        {
            InitializeComponent();
        }

        private void serviceInstaller1_BeforeInstall(object sender, InstallEventArgs e)
        {
            try
            {
                RegistryKey key = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32)
                .CreateSubKey(@"SOFTWARE\LDM");
                if (key != null)
                {
                    Process _prc = new Process();
                    _prc.StartInfo.FileName = "cmd.exe";
                    _prc.StartInfo.UseShellExecute = false;
                    _prc.StartInfo.RedirectStandardOutput = true;
                    _prc.StartInfo.RedirectStandardInput = true;
                    _prc.Start();

                    ConsoleColor _color = Console.ForegroundColor;
                    Console.ForegroundColor = ConsoleColor.DarkGreen;
                    Console.WriteLine("\n\n");
                    Console.WriteLine("PLEASE ENTER FOLLOWING DETAILS TO COMPLETE SETUP");
                    Console.WriteLine("NOTE: if you enter wrong information, you will need to reinstall the application.");
                    Console.WriteLine("\n\n");
                    Console.WriteLine("Enter INPUT_DIRECTORY (FULL PATH):");
                    key.SetValue("INPUT_DIRECTORY", Console.ReadLine().Trim());
                    Console.WriteLine("Enter OUTPUT_DIRECTORY (FULL PATH):");
                    key.SetValue("OUTPUT_DIRECTORY", Console.ReadLine().Trim());
                    Console.WriteLine("Enter BACKUP_DIRECTORY (FULL PATH):");
                    key.SetValue("BACKUP_DIRECTORY", Console.ReadLine().Trim());
                    Console.WriteLine("Enter BACKUP_INCORECTFORMAT_DIRECTORY (FULL PATH):");
                    key.SetValue("BACKUP_INCORECTFORMAT_DIRECTORY:", Console.ReadLine().Trim());
                    Console.WriteLine("Enter DB_HOST:");
                    key.SetValue("DB_HOST", Console.ReadLine().Trim());
                    Console.WriteLine("Enter DB_USERNAME:");
                    key.SetValue("DB_USERNAME", Console.ReadLine().Trim());
                    Console.WriteLine("Enter DB_PASSWORD:");
                    key.SetValue("DB_PASSWORD", Console.ReadLine().Trim());
                    Console.WriteLine("Enter DATABASE NAME:");
                    key.SetValue("DB_NAME", Console.ReadLine().Trim());
                    key.Close();
                    Console.ForegroundColor = _color;
                    _prc.Close();
                }
            }
            catch (Exception ex)
            {
            }
        }
    }
}
