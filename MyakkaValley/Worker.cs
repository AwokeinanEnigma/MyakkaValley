using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.DirectoryServices.ActiveDirectory;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Management;

namespace MyakkaValley
{

    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;

        public Worker(ILogger<Worker> logger)
        {
            _logger = logger;
        }
        
        static bool IsOnDomain()
        {
            try
            {
                Domain.GetComputerDomain();
                return true;
            }
            catch (ActiveDirectoryObjectNotFoundException e)
            {
                return false;
            }
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            bool RenameComputer(string newName, string prefix = "SAL")
            {
                ProcessStartInfo info = new ProcessStartInfo();
                info.FileName = "powershell";
                info.Verb = "runas";
                info.Arguments = $"Rename-Computer -NewName {prefix + newName} -Force -Restart";
                info.CreateNoWindow = true;
                info.RedirectStandardOutput = true;
                info.UseShellExecute = false;
                var p = Process.Start(info);
                string result = p.StandardOutput.ReadToEnd();
                
                if (result.Contains("Fail") || result.Contains("Error"))
                {
                    _logger.LogError($"Failed to rename computer, full exception: {result}");
                    return false;
                }
                if (result.Contains("Success"))
                {
                    _logger.LogInformation($"Successfully renamed computer to {prefix + newName}");
                    return true;
                }
                
                _logger.LogError($"Failed to rename computer for unknown reasons, full exception: {result}");
                return false;
            }
            void IssueShutdownCommand()
            {
                ProcessStartInfo info = new ProcessStartInfo();
                info.FileName = "shutdown";
                info.Verb = "runas";
                info.Arguments = @"/g /t 300 /f /c ""The computer will be restarted in 5 minutes to finish up a name change.""";
                info.CreateNoWindow = true;
                info.RedirectStandardOutput = true;
                info.UseShellExecute = false;
                var p = Process.Start(info);
                if (p.Responding)
                {
                    Console.WriteLine("Issued shutdown command - will restart in 5 minutes.");
                }
            }

            void CreateIntuneCheckFile()
            {
                try
                {
                    Directory.CreateDirectory("C:\\MyakkaValley");
                    File.WriteAllText("C:\\MyakkaValley\\WeExist.txt",
                        "Disregard this file - its purpose to ensure intune knows we (as in, this program) exist.");
                }
                catch (Exception e)
                {
                    _logger.LogError($"Failed to create Intune check file: {e.Message}. Things will likely break (or have broken.)");
                }
            }

            string GetSerialNumber()
            {
                ProcessStartInfo info = new ProcessStartInfo();
                info.FileName = "powershell";
                info.Verb = "runas";
                info.Arguments = "Get-WmiObject Win32_Bios | Select-Object SerialNumber";
                info.CreateNoWindow = true;
                info.RedirectStandardOutput = true;
                info.UseShellExecute = false;
                var p = Process.Start(info);
                // we do this to remove any extra characters that might be in the output.
                return p.StandardOutput.ReadToEnd().Replace("SerialNumber", string.Empty).Replace("------------", string.Empty).Trim();
            }
            
            while (!stoppingToken.IsCancellationRequested)
            {
                CreateIntuneCheckFile();
                if (IsOnDomain())
                {
                    _logger.LogInformation("We're on the domain, going to rename the computer.");
                    if (RenameComputer(GetSerialNumber()))
                    {
                        _logger.LogInformation("Successfully renamed computer. Restarting in 5 minutes.");
                        IssueShutdownCommand();
                    }
                    else
                    {
                        _logger.LogError("Failed to rename computer. Check previous logs for more information.");
                    }
                }
                else
                {
                    _logger.LogError("We're not on the domain, so we're going to wait 30 seconds and check again.");
                }
                await Task.Delay(30000, stoppingToken);
            }
        }
    }
}