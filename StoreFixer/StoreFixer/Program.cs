using StoreFixer.Utils;
using System.Runtime.CompilerServices;
using System.Security.Principal;
using System.ServiceProcess;
using System.Text;

namespace StoreFixer
{
    internal class Program
    {
        private static string logFilePath = string.Empty;

        // Console color scheme
        private enum MessageType
        {
            Info,
            Success,
            Warning,
            Error,
            Header
        }

        static async Task Main(string[] args)
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            logFilePath = Path.Combine(desktopPath, "StoreFixer_Log.txt");

            string systemDrive = Environment.GetEnvironmentVariable("SYSTEMDRIVE") ?? "C:";

            LogColored($"System Drive: {systemDrive}", MessageType.Info);
            LogColored($"=== StoreFixer Execution Started ===", MessageType.Header);
            LogColored($"Log file: {logFilePath}", MessageType.Info);
            LogColored($"Current User: {WindowsIdentity.GetCurrent().Name}", MessageType.Info);
            LogColored($"Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss}", MessageType.Info);

            // CHANGE THIS BACK TO A TRUE ONLY STATEMENT. FOR TESTING ONLY
            if (IsRunAsTi())
            {
                LogColored("Executable is provided under CC0 1.0 Universal", MessageType.Info);
                LogColored("Running as Trusted Installer. Starting execution...", MessageType.Success);
                await Execution();
            }
            else
            {
                string errorMsg = "Executable was not ran as Trusted Installer.\nPlease close this executable and run it as Trusted Installer.";
                LogColored(errorMsg, MessageType.Error);
            }

            LogColored("=== StoreFixer Execution Completed ===", MessageType.Header);
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.Write("Press any key to exit...");
            Console.ResetColor();
            Console.ReadKey();
        }

        /// <summary>
        /// Logs messages to both console and file with color
        /// </summary>
        private static void LogColored(string message, MessageType type = MessageType.Info)
        {
            string timestampedMessage = $"[{DateTime.Now:HH:mm:ss}] {message}";

            // Set console color based on message type
            ConsoleColor color = type switch
            {
                MessageType.Success => ConsoleColor.Green,
                MessageType.Warning => ConsoleColor.Yellow,
                MessageType.Error => ConsoleColor.Red,
                MessageType.Header => ConsoleColor.Cyan,
                _ => ConsoleColor.White
            };

            // Write to console with color
            Console.ForegroundColor = color;
            Console.WriteLine(timestampedMessage);
            Console.ResetColor();

            // Write to file
            try
            {
                File.AppendAllText(logFilePath, timestampedMessage + Environment.NewLine);
            }
            catch
            {
                // Silently fail if logging to file fails
            }
        }

        /// <summary>
        /// Logs messages to both console and file (backward compatibility)
        /// </summary>
        private static void Log(string message)
        {
            LogColored(message, MessageType.Info);
        }

        /// <summary>
        /// Checks if exec is ran as ti. should be the main entrypoint of the app.
        /// </summary>
        /// <returns></returns>
        private static bool IsRunAsTi()
        {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            string userName = identity.Name;

            if (userName.Equals("NT AUTHORITY\\SYSTEM", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// Main execution
        /// </summary>
        private static async Task Execution()
        {
            try
            {
                // set up variables
                string targetFile = $"{Environment.GetEnvironmentVariable("SYSTEMDRIVE")}\\ProgramData\\Microsoft\\Windows\\AppRepository\\StateRepository-Deployment.srd";
                LogColored($"Target file path: {targetFile}", MessageType.Info);

                List<bool> canContinue = new();
                int maxRetries = 5;
                bool fileDeleted = false;

                LogColored("Retrieving services...", MessageType.Info);
                HashSet<ServiceController> rootServices =
                    GetServices([ "ClipSVC", "AppXSvc", "StateRepository" ]);
                LogColored($"Root services found: {rootServices.Count}", MessageType.Success);

                HashSet<ServiceController> dependentServices = GetDependentServices(rootServices);
                LogColored($"Dependent services found: {dependentServices.Count}", MessageType.Success);

                Dictionary<string, ServiceStartMode> servicesStartMode = new Dictionary<string, ServiceStartMode>();
                foreach (ServiceController serviceController in rootServices)
                {
                    try
                    {
                        servicesStartMode.TryAdd(serviceController.ServiceName, serviceController.StartType);
                        LogColored($"Backed up root service: {serviceController.ServiceName} ({serviceController.StartType})", MessageType.Info);
                    }
                    catch (Exception ex)
                    {
                        LogColored($"Failed to backup root service {serviceController.ServiceName}: {ex.Message}", MessageType.Error);
                    }
                }

                foreach (ServiceController serviceController in dependentServices)
                {
                    try
                    {
                        servicesStartMode.TryAdd(serviceController.ServiceName, serviceController.StartType);
                        LogColored($"Backed up dependent service: {serviceController.ServiceName} ({serviceController.StartType})", MessageType.Info);
                    }
                    catch (Exception ex)
                    {
                        LogColored($"Failed to backup dependent service {serviceController.ServiceName}: {ex.Message}", MessageType.Error);
                    }
                }

                LogColored($"Total services backed up: {servicesStartMode.Count}", MessageType.Success);
                LogColored("Saving backup to registry...", MessageType.Info);

                foreach(KeyValuePair<string, ServiceStartMode> kvp in servicesStartMode)
                {
                    try
                    {
                        RegistryHelper.SetValue(@"HKLM\SOFTWARE\AtlasOS\Temp", kvp.Key, kvp.Value.ToString(), Microsoft.Win32.RegistryValueKind.String);
                    }
                    catch (Exception ex)
                    {
                        LogColored($"Failed to save backup for {kvp.Key}: {ex.Message}", MessageType.Error);
                    }
                }

                // begin trying to set everything now that we have a backup of services startup type and that we have them in list
                LogColored("Disabling services...", MessageType.Info);
                /// start by dependent services
                await DisableServicesAsync(dependentServices);
                await Task.Delay(1000);
                await DisableServicesAsync(rootServices);
                await Task.Delay(1000);

                LogColored("Stopping services...", MessageType.Info);
                await StopServicesAsync(dependentServices);
                await Task.Delay(1000);
                await StopServicesAsync(rootServices);
                await Task.Delay(2000);

                LogColored("Deleting target file...", MessageType.Info);
                await DeleteTargetFileAsync(targetFile, rootServices, dependentServices);
                await Task.Delay(1000);

                LogColored("Restoring services to original state...", MessageType.Info);
                await RestoreServicesAsync(servicesStartMode);
                await Task.Delay(1000);

                // Start services (dependent services first, then root services)
                LogColored("Starting services...", MessageType.Info);
                await StartServicesAsync(dependentServices);
                await Task.Delay(1000);
                await StartServicesAsync(rootServices);
                await Task.Delay(1000);

                LogColored("Execution completed successfully!", MessageType.Success);
            }
            catch (Exception ex)
            {
                LogColored($"CRITICAL ERROR during execution: {ex.Message}", MessageType.Error);
                LogColored($"Stack trace: {ex.StackTrace}", MessageType.Error);
            }
        }

        /// <summary>
        /// Makes sure to disable every services
        /// </summary>
        /// <param name="services"></param>
        /// <returns></returns>
        private static async Task DisableServicesAsync(HashSet<ServiceController> services)
        {
            LogColored($"Disabling {services.Count} services...", MessageType.Info);

            // Disable all services
            foreach (ServiceController service in services)
            {
                try
                {
                    ServiceHelper.SetStartupType(service.ServiceName, ServiceStartMode.Disabled);
                    LogColored($"✓ Disabled service: {service.ServiceName}", MessageType.Success);
                }
                catch (Exception ex)
                {
                    LogColored($"✗ Failed to disable service {service.ServiceName}: {ex.Message}", MessageType.Error);
                }
                await Task.Delay(100);
            }

            bool allDisabled = false;
            int maxRetries = 10;
            int retryCount = 0;

            while (!allDisabled && retryCount < maxRetries)
            {
                allDisabled = true;

                foreach (ServiceController service in services)
                {
                    try
                    {
                        if (ServiceHelper.GetStartupType(service.ServiceName) != ServiceStartMode.Disabled)
                        {
                            allDisabled = false;
                            ServiceHelper.SetStartupType(service.ServiceName, ServiceStartMode.Disabled);
                            LogColored($"↻ Retrying disable for service: {service.ServiceName}", MessageType.Warning);
                        }
                    }
                    catch (Exception ex)
                    {
                        LogColored($"✗ Verification error for {service.ServiceName}: {ex.Message}", MessageType.Error);
                    }
                }

                if (!allDisabled)
                {
                    await Task.Delay(500);
                }
                retryCount++;
            }

            if (allDisabled)
                LogColored($"✓ All {services.Count} services have been successfully disabled.", MessageType.Success);
            else
                LogColored($"⚠ Warning: Not all services were disabled after {maxRetries} retries.", MessageType.Warning);
        }

        /// <summary>
        /// Restores services to their original startup mode
        /// </summary>
        /// <param name="servicesStartMode">Dictionary containing original startup modes</param>
        private static async Task RestoreServicesAsync(Dictionary<string, ServiceStartMode> servicesStartMode)
        {
            LogColored($"Restoring {servicesStartMode.Count} services to original startup modes...", MessageType.Info);

            foreach (KeyValuePair<string, ServiceStartMode> kvp in servicesStartMode)
            {
                try
                {
                    ServiceHelper.SetStartupType(kvp.Key, kvp.Value);
                    LogColored($"✓ Restored service \"{kvp.Key}\" to {kvp.Value}", MessageType.Success);
                }
                catch (Exception ex)
                {
                    LogColored($"✗ Failed to restore service \"{kvp.Key}\": {ex.Message}", MessageType.Error);
                }
                await Task.Delay(100);
            }

            // Verify restoration
            bool allRestored = false;
            int maxRetries = 10;
            int retryCount = 0;

            while (!allRestored && retryCount < maxRetries)
            {
                allRestored = true;

                foreach (KeyValuePair<string, ServiceStartMode> kvp in servicesStartMode)
                {
                    try
                    {
                        if (ServiceHelper.GetStartupType(kvp.Key) != kvp.Value)
                        {
                            allRestored = false;
                            ServiceHelper.SetStartupType(kvp.Key, kvp.Value);
                            LogColored($"↻ Retrying restore for service: {kvp.Key}", MessageType.Warning);
                        }
                    }
                    catch (Exception ex)
                    {
                        LogColored($"✗ Verification error for {kvp.Key}: {ex.Message}", MessageType.Error);
                    }
                }

                if (!allRestored)
                {
                    await Task.Delay(500);
                }
                retryCount++;
            }

            if (allRestored)
                LogColored("✓ All services have been successfully restored.", MessageType.Success);
            else
                LogColored($"⚠ Warning: Not all services were restored after {maxRetries} retries.", MessageType.Warning);
        }

        /// <summary>
        /// Deletes the target file after ensuring all services are stopped
        /// </summary>
        /// <param name="targetFile">Path to the file to delete</param>
        /// <param name="rootServices">Root services to stop before deletion</param>
        /// <param name="dependentServices">Dependent services to stop before deletion</param>
        private static async Task DeleteTargetFileAsync(string targetFile, HashSet<ServiceController> rootServices, HashSet<ServiceController> dependentServices)
        {
            LogColored("Ensuring all services are stopped before file deletion...", MessageType.Info);
            await StopServicesAsync(dependentServices);
            await StopServicesAsync(rootServices);
            await Task.Delay(1000);

            try
            {
                if (File.Exists(targetFile))
                {
                    File.Delete(targetFile);
                    LogColored($"✓ Successfully deleted file: {targetFile}", MessageType.Success);
                }
                else
                {
                    LogColored($"⚠ Target file not found: {targetFile}", MessageType.Warning);
                }
            }
            catch (Exception ex)
            {
                LogColored($"✗ Failed to delete file: {ex.Message}", MessageType.Error);
            }
        }

        /// <summary>
        /// Stops services and verifies they are stopped
        /// </summary>
        /// <param name="services">HashSet of services to stop</param>
        private static async Task StopServicesAsync(HashSet<ServiceController> services)
        {
            LogColored($"Stopping {services.Count} services...", MessageType.Info);

            // Stop all services
            foreach (ServiceController service in services)
            {
                try
                {
                    service.Refresh();
                    if (service.Status != ServiceControllerStatus.Stopped)
                    {
                        service.Stop();
                        LogColored($"⏹ Stopping service \"{service.ServiceName}\"...", MessageType.Info);
                    }
                    else
                    {
                        LogColored($"✓ Service \"{service.ServiceName}\" already stopped", MessageType.Success);
                    }
                }
                catch (Exception ex)
                {
                    LogColored($"✗ Failed to stop service \"{service.ServiceName}\": {ex.Message}", MessageType.Error);
                }
                await Task.Delay(100);
            }

            // Verify all services are stopped
            bool allStopped = false;
            int maxRetries = 10;
            int retryCount = 0;

            while (!allStopped && retryCount < maxRetries)
            {
                allStopped = true;
                await Task.Delay(500); // Wait for services to stop

                foreach (ServiceController service in services)
                {
                    try
                    {
                        service.Refresh();
                        if (service.Status != ServiceControllerStatus.Stopped)
                        {
                            allStopped = false;
                            LogColored($"⚠ Service \"{service.ServiceName}\" still running (status: {service.Status}), retrying...", MessageType.Warning);
                            if (retryCount < maxRetries - 1)
                            {
                                service.Stop();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogColored($"✗ Error checking service {service.ServiceName}: {ex.Message}", MessageType.Error);
                    }
                }
                retryCount++;
            }

            if (allStopped)
                LogColored($"✓ All {services.Count} services have been successfully stopped.", MessageType.Success);
            else
                LogColored($"⚠ Warning: Not all services are stopped after {maxRetries} retries.", MessageType.Warning);
        }

        /// <summary>
        /// Starts services and verifies they are running
        /// </summary>
        /// <param name="services">HashSet of services to start</param>
        private static async Task StartServicesAsync(HashSet<ServiceController> services)
        {
            LogColored($"Starting {services.Count} services...", MessageType.Info);

            // Start all services
            foreach (ServiceController service in services)
            {
                try
                {
                    service.Refresh();
                    if (service.Status != ServiceControllerStatus.Running)
                    {
                        service.Start();
                        LogColored($"▶ Starting service \"{service.ServiceName}\"...", MessageType.Info);
                    }
                    else
                    {
                        LogColored($"✓ Service \"{service.ServiceName}\" already running", MessageType.Success);
                    }
                }
                catch (Exception ex)
                {
                    LogColored($"✗ Failed to start service \"{service.ServiceName}\": {ex.Message}", MessageType.Error);
                }
                await Task.Delay(100);
            }

            // Verify all services are running
            bool allRunning = false;
            int maxRetries = 10;
            int retryCount = 0;

            while (!allRunning && retryCount < maxRetries)
            {
                allRunning = true;
                await Task.Delay(500); // Wait for services to start

                foreach (ServiceController service in services)
                {
                    try
                    {
                        service.Refresh();
                        if (service.Status != ServiceControllerStatus.Running)
                        {
                            allRunning = false;
                            LogColored($"⚠ Service \"{service.ServiceName}\" not running (status: {service.Status}), retrying...", MessageType.Warning);
                            if (retryCount < maxRetries - 1)
                            {
                                service.Start();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogColored($"✗ Error checking service {service.ServiceName}: {ex.Message}", MessageType.Error);
                    }
                }
                retryCount++;
            }

            if (allRunning)
                LogColored($"✓ All {services.Count} services have been successfully started.", MessageType.Success);
            else
                LogColored($"⚠ Warning: Not all services are running after {maxRetries} retries.", MessageType.Warning);
        }

        private static HashSet<ServiceController> GetDependentServices(HashSet<ServiceController> services)
        {
            HashSet<ServiceController> serviceControllers = new HashSet<ServiceController>();

            foreach (ServiceController service in services)
                serviceControllers.UnionWith(service.DependentServices);
 
            return serviceControllers;
        }
        private static HashSet<ServiceController> GetServices(string[] services)
        {
            
            HashSet<ServiceController> serviceControllers = new HashSet<ServiceController>();

            foreach (string service in services)
                serviceControllers.Add(ServiceHelper.GetServiceController(service));

            return serviceControllers;
        }
    }
}
