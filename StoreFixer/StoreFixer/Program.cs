using StoreFixer.Utils;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.ServiceProcess;
using System.Text;

namespace StoreFixer
{
    internal class Program
    {
        private static string logFilePath = string.Empty;
        private static Dictionary<string, ServiceStartMode> servicesBackup = new();
        private static HashSet<ServiceController> allServices = new();
        private static bool executionStarted = false;

        // Console color scheme
        private enum MessageType
        {
            Info,
            Success,
            Warning,
            Error,
            Header,
            Critical
        }

        static async Task Main(string[] args)
        {
            try
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                logFilePath = Path.Combine(desktopPath, "StoreFixer_Log.txt");

                string systemDrive = Environment.GetEnvironmentVariable("SYSTEMDRIVE") ?? "C:";

                // Print startup banner
                Console.WriteLine();
                LogColored("╔════════════════════════════════════════════════════════════╗", MessageType.Header);
                LogColored("║         StoreFixer - Service Restoration Utility            ║", MessageType.Header);
                LogColored("╚════════════════════════════════════════════════════════════╝", MessageType.Header);
                Console.WriteLine();

                LogColored($"System Drive:        {systemDrive}", MessageType.Info);
                LogColored($"Log File:            {logFilePath}", MessageType.Info);
                LogColored($"Current User:        {WindowsIdentity.GetCurrent().Name}", MessageType.Info);
                LogColored($"Timestamp:           {DateTime.Now:yyyy-MM-dd HH:mm:ss}", MessageType.Info);
                Console.WriteLine();

                if (IsRunAsTi())
                {
                    LogColored("Status: Running as Trusted Installer", MessageType.Success);
                    Console.WriteLine();
                    LogColored("Starting execution...", MessageType.Success);
                    Console.WriteLine();

                    try
                    {
                        await Execution();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine();
                        LogColored("════════════════════════════════════════════════════════════", MessageType.Critical);
                        LogColored($"CRITICAL ERROR: {ex.Message}", MessageType.Critical);
                        LogColored($"Stack trace: {ex.StackTrace}", MessageType.Critical);
                        LogColored("════════════════════════════════════════════════════════════", MessageType.Critical);
                        Console.WriteLine();
                        await RestoreOnCrash();
                    }
                }
                else
                {
                    Console.WriteLine();
                    LogColored("════════════════════════════════════════════════════════════", MessageType.Error);
                    LogColored("ERROR: Not running as Trusted Installer", MessageType.Error);
                    LogColored("Please close this application and run it as Trusted Installer.", MessageType.Error);
                    LogColored("════════════════════════════════════════════════════════════", MessageType.Error);
                    Console.WriteLine();
                }

                Console.WriteLine();
                LogColored("═══════════════════════════════════════════════════════════════", MessageType.Header);
                LogColored("StoreFixer Execution Completed", MessageType.Header);
                LogColored("═══════════════════════════════════════════════════════════════", MessageType.Header);
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                LogColored("════════════════════════════════════════════════════════════", MessageType.Critical);
                LogColored($"FATAL ERROR in Main: {ex.Message}", MessageType.Critical);
                LogColored($"Stack trace: {ex.StackTrace}", MessageType.Critical);
                LogColored("════════════════════════════════════════════════════════════", MessageType.Critical);
                Console.WriteLine();

                try
                {
                    await RestoreOnCrash();
                }
                catch { }
            }
            finally
            {
                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.Write("Press any key to exit...");
                Console.ResetColor();
                try { Console.ReadKey(); } catch { }
            }
        }

        /// <summary>
        /// Emergency restoration if execution crashes
        /// </summary>
        private static async Task RestoreOnCrash()
        {
            Console.WriteLine();
            LogColored("╔════════════════════════════════════════════════════════════╗", MessageType.Critical);
            LogColored("║                  EMERGENCY RESTORATION                     ║", MessageType.Critical);
            LogColored("║                    PLEASE WAIT...                          ║", MessageType.Critical);
            LogColored("╚════════════════════════════════════════════════════════════╝", MessageType.Critical);
            Console.WriteLine();

            if (servicesBackup.Count > 0)
            {
                LogColored($"Restoring {servicesBackup.Count} services from backup...", MessageType.Warning);
                Console.WriteLine();
                await RestoreServicesAsync(servicesBackup);
                await Task.Delay(1000);
                Console.WriteLine();

                if (allServices.Count > 0)
                {
                    LogColored($"Starting {allServices.Count} services...", MessageType.Warning);
                    Console.WriteLine();
                    await StartServicesAsync(allServices);
                    Console.WriteLine();
                }
            }
            else
            {
                Console.WriteLine();
                LogColored("⚠ No service backup found - manual restoration may be needed", MessageType.Warning);
                Console.WriteLine();
            }

            LogColored("════════════════════════════════════════════════════════════", MessageType.Critical);
            LogColored("Emergency restoration complete.", MessageType.Info);
            LogColored("════════════════════════════════════════════════════════════", MessageType.Critical);
            Console.WriteLine();
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
                MessageType.Critical => ConsoleColor.Magenta,
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
            try
            {
                WindowsIdentity identity = WindowsIdentity.GetCurrent();
                string userName = identity.Name;

                if (userName.Equals("NT AUTHORITY\\SYSTEM", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                LogColored($"Error checking Trusted Installer status: {ex.Message}", MessageType.Error);
                return false;
            }
        }

        /// <summary>
        /// Main execution
        /// </summary>
        private static async Task Execution()
        {
            try
            {
                executionStarted = true;

                // ====================================================================
                // PHASE 1: SERVICE BACKUP
                // ====================================================================
                Console.WriteLine();
                LogColored("╔════════════════════════════════════════════════════════════╗", MessageType.Header);
                LogColored("║ PHASE 1: Creating Emergency Service Backup                 ║", MessageType.Header);
                LogColored("╚════════════════════════════════════════════════════════════╝", MessageType.Header);
                Console.WriteLine();

                await CreateServiceBackup();
                Console.WriteLine();

                // ====================================================================
                // PHASE 2: SERVICE RETRIEVAL & PREPARATION
                // ====================================================================
                LogColored("╔════════════════════════════════════════════════════════════╗", MessageType.Header);
                LogColored("║ PHASE 2: Retrieving Services & Configuration              ║", MessageType.Header);
                LogColored("╚════════════════════════════════════════════════════════════╝", MessageType.Header);
                Console.WriteLine();

                string targetFile = $"{Environment.GetEnvironmentVariable("SYSTEMDRIVE")}\\ProgramData\\Microsoft\\Windows\\AppRepository\\StateRepository-Deployment.srd";
                LogColored($"Target File: {targetFile}", MessageType.Info);
                Console.WriteLine();

                try
                {
                    LogColored("Retrieving services...", MessageType.Info);
                    HashSet<ServiceController> rootServices = GetServices(["ClipSVC", "AppXSvc", "StateRepository"]);
                    LogColored($"  → Root services found: {rootServices.Count}", MessageType.Success);
                    Console.WriteLine();

                    HashSet<ServiceController> dependentServices = GetDependentServices(rootServices);
                    LogColored($"  → Dependent services found: {dependentServices.Count}", MessageType.Success);
                    Console.WriteLine();

                    Dictionary<string, ServiceStartMode> servicesStartMode = new();
                    LogColored("Backing up root services...", MessageType.Info);
                    foreach (ServiceController serviceController in rootServices)
                    {
                        try
                        {
                            servicesStartMode.TryAdd(serviceController.ServiceName, serviceController.StartType);
                            LogColored($"    • {serviceController.ServiceName} ({serviceController.StartType})", MessageType.Info);
                        }
                        catch (Exception ex)
                        {
                            LogColored($"    ✗ {serviceController.ServiceName}: {ex.Message}", MessageType.Error);
                        }
                    }
                    Console.WriteLine();

                    LogColored("Backing up dependent services...", MessageType.Info);
                    foreach (ServiceController serviceController in dependentServices)
                    {
                        try
                        {
                            servicesStartMode.TryAdd(serviceController.ServiceName, serviceController.StartType);
                            LogColored($"    • {serviceController.ServiceName} ({serviceController.StartType})", MessageType.Info);
                        }
                        catch (Exception ex)
                        {
                            LogColored($"    ✗ {serviceController.ServiceName}: {ex.Message}", MessageType.Error);
                        }
                    }
                    Console.WriteLine();

                    LogColored($"Total services backed up: {servicesStartMode.Count}", MessageType.Success);
                    LogColored("Saving backup to registry...", MessageType.Info);
                    Console.WriteLine();

                    foreach (KeyValuePair<string, ServiceStartMode> kvp in servicesStartMode)
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
                    Console.WriteLine();

                    // ================================================================
                    // PHASE 3: DISABLING SERVICES
                    // ================================================================
                    LogColored("╔════════════════════════════════════════════════════════════╗", MessageType.Header);
                    LogColored("║ PHASE 3: Disabling Services                                ║", MessageType.Header);
                    LogColored("╚════════════════════════════════════════════════════════════╝", MessageType.Header);
                    Console.WriteLine();

                    LogColored("Disabling dependent services...", MessageType.Info);
                    await DisableServicesAsync(dependentServices);
                    await Task.Delay(1000);
                    Console.WriteLine();

                    LogColored("Disabling root services...", MessageType.Info);
                    await DisableServicesAsync(rootServices);
                    await Task.Delay(1000);
                    Console.WriteLine();

                    // ================================================================
                    // PHASE 4: STOPPING SERVICES
                    // ================================================================
                    LogColored("╔════════════════════════════════════════════════════════════╗", MessageType.Header);
                    LogColored("║ PHASE 4: Stopping Services                                 ║", MessageType.Header);
                    LogColored("╚════════════════════════════════════════════════════════════╝", MessageType.Header);
                    Console.WriteLine();

                    LogColored("Stopping dependent services...", MessageType.Info);
                    await StopServicesAsync(dependentServices);
                    await Task.Delay(1000);
                    Console.WriteLine();

                    LogColored("Stopping root services...", MessageType.Info);
                    await StopServicesAsync(rootServices);
                    await Task.Delay(2000);
                    Console.WriteLine();

                    // ================================================================
                    // PHASE 5: FILE DELETION
                    // ================================================================
                    LogColored("╔════════════════════════════════════════════════════════════╗", MessageType.Header);
                    LogColored("║ PHASE 5: Deleting Target File                              ║", MessageType.Header);
                    LogColored("╚════════════════════════════════════════════════════════════╝", MessageType.Header);
                    Console.WriteLine();

                    await DeleteTargetFileAsync(targetFile, rootServices, dependentServices);
                    await Task.Delay(1000);
                    Console.WriteLine();

                    // ================================================================
                    // PHASE 6: SERVICE RESTORATION
                    // ================================================================
                    LogColored("╔════════════════════════════════════════════════════════════╗", MessageType.Header);
                    LogColored("║ PHASE 6: Restoring Services to Original State              ║", MessageType.Header);
                    LogColored("╚════════════════════════════════════════════════════════════╝", MessageType.Header);
                    Console.WriteLine();

                    await RestoreServicesAsync(servicesStartMode);
                    await Task.Delay(1000);
                    Console.WriteLine();

                    // ================================================================
                    // PHASE 7: STARTING SERVICES
                    // ================================================================
                    LogColored("╔════════════════════════════════════════════════════════════╗", MessageType.Header);
                    LogColored("║ PHASE 7: Starting Services                                 ║", MessageType.Header);
                    LogColored("╚════════════════════════════════════════════════════════════╝", MessageType.Header);
                    Console.WriteLine();

                    LogColored("Starting dependent services...", MessageType.Info);
                    await StartServicesAsync(dependentServices);
                    await Task.Delay(1000);
                    Console.WriteLine();

                    LogColored("Starting root services...", MessageType.Info);
                    await StartServicesAsync(rootServices);
                    await Task.Delay(1000);
                    Console.WriteLine();

                    // ================================================================
                    // SUCCESS
                    // ================================================================
                    LogColored("════════════════════════════════════════════════════════════", MessageType.Success);
                    LogColored("Execution completed successfully!", MessageType.Success);
                    LogColored("════════════════════════════════════════════════════════════", MessageType.Success);
                    Console.WriteLine();
                }
                catch (Exception ex)
                {
                    Console.WriteLine();
                    LogColored("════════════════════════════════════════════════════════════", MessageType.Error);
                    LogColored($"Error during service operations: {ex.Message}", MessageType.Error);
                    LogColored($"Stack trace: {ex.StackTrace}", MessageType.Error);
                    LogColored("════════════════════════════════════════════════════════════", MessageType.Error);
                    Console.WriteLine();

                    // If something fails during operations, immediately restore
                    LogColored("Triggering automatic restoration...", MessageType.Warning);
                    Console.WriteLine();
                    if (servicesBackup.Count > 0)
                    {
                        await RestoreServicesAsync(servicesBackup);
                        await StartServicesAsync(allServices);
                    }

                    throw;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                LogColored("════════════════════════════════════════════════════════════", MessageType.Critical);
                LogColored($"CRITICAL ERROR during execution: {ex.Message}", MessageType.Critical);
                LogColored($"Stack trace: {ex.StackTrace}", MessageType.Critical);
                LogColored("════════════════════════════════════════════════════════════", MessageType.Critical);
                Console.WriteLine();
                throw;
            }
        }

        /// <summary>
        /// Creates an immediate backup of all services before any operations
        /// </summary>
        private static async Task CreateServiceBackup()
        {
            try
            {
                HashSet<ServiceController> rootServices = GetServices(["ClipSVC", "AppXSvc", "StateRepository"]);
                HashSet<ServiceController> dependentServices = GetDependentServices(rootServices);

                // Store all services for emergency restoration
                allServices.UnionWith(rootServices);
                allServices.UnionWith(dependentServices);

                foreach (ServiceController service in allServices)
                {
                    try
                    {
                        servicesBackup.TryAdd(service.ServiceName, service.StartType);
                    }
                    catch (Exception ex)
                    {
                        LogColored($"Failed to backup {service.ServiceName}: {ex.Message}", MessageType.Error);
                    }
                }

                LogColored($"Created emergency backup for {servicesBackup.Count} services", MessageType.Success);
                await Task.Delay(500);
            }
            catch (Exception ex)
            {
                LogColored($"Failed to create service backup: {ex.Message}", MessageType.Error);
                LogColored("Continuing without backup - ensure manual restoration capability", MessageType.Warning);
            }
        }

        /// <summary>
        /// Makes sure to disable every services
        /// </summary>
        /// <param name="services"></param>
        /// <returns></returns>
        private static async Task DisableServicesAsync(HashSet<ServiceController> services)
        {
            try
            {
                LogColored($"Disabling {services.Count} services...", MessageType.Info);

                // Disable all services
                foreach (ServiceController service in services)
                {
                    try
                    {
                        ServiceHelper.SetStartupType(service.ServiceName, ServiceStartMode.Disabled);
                        LogColored($"Disabled service: {service.ServiceName}", MessageType.Success);
                    }
                    catch (Exception ex)
                    {
                        LogColored($"Failed to disable service {service.ServiceName}: {ex.Message}", MessageType.Error);
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
                                LogColored($"Retrying disable for service: {service.ServiceName}", MessageType.Warning);
                            }
                        }
                        catch (Exception ex)
                        {
                            LogColored($"Verification error for {service.ServiceName}: {ex.Message}", MessageType.Error);
                        }
                    }

                    if (!allDisabled)
                    {
                        await Task.Delay(500);
                    }
                    retryCount++;
                }

                if (allDisabled)
                    LogColored($"All {services.Count} services have been successfully disabled.", MessageType.Success);
                else
                    LogColored($"Warning: Not all services were disabled after {maxRetries} retries.", MessageType.Warning);
            }
            catch (Exception ex)
            {
                LogColored($"Error in DisableServicesAsync: {ex.Message}", MessageType.Error);
            }
        }

        /// <summary>
        /// Restores services to their original startup mode
        /// </summary>
        /// <param name="servicesStartMode">Dictionary containing original startup modes</param>
        private static async Task RestoreServicesAsync(Dictionary<string, ServiceStartMode> servicesStartMode)
        {
            try
            {
                LogColored($"Restoring {servicesStartMode.Count} services to original startup modes...", MessageType.Info);

                foreach (KeyValuePair<string, ServiceStartMode> kvp in servicesStartMode)
                {
                    try
                    {
                        ServiceHelper.SetStartupType(kvp.Key, kvp.Value);
                        LogColored($"Restored service \"{kvp.Key}\" to {kvp.Value}", MessageType.Success);
                    }
                    catch (Exception ex)
                    {
                        LogColored($"Failed to restore service \"{kvp.Key}\": {ex.Message}", MessageType.Error);
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
                                LogColored($"Retrying restore for service: {kvp.Key}", MessageType.Warning);
                            }
                        }
                        catch (Exception ex)
                        {
                            LogColored($"Verification error for {kvp.Key}: {ex.Message}", MessageType.Error);
                        }
                    }

                    if (!allRestored)
                    {
                        await Task.Delay(500);
                    }
                    retryCount++;
                }

                if (allRestored)
                    LogColored("All services have been successfully restored.", MessageType.Success);
                else
                    LogColored($"Warning: Not all services were restored after {maxRetries} retries.", MessageType.Warning);
            }
            catch (Exception ex)
            {
                LogColored($"Error in RestoreServicesAsync: {ex.Message}", MessageType.Error);
            }
        }

        /// <summary>
        /// Deletes the target file after ensuring all services are stopped
        /// </summary>
        /// <param name="targetFile">Path to the file to delete</param>
        /// <param name="rootServices">Root services to stop before deletion</param>
        /// <param name="dependentServices">Dependent services to stop before deletion</param>
        private static async Task DeleteTargetFileAsync(string targetFile, HashSet<ServiceController> rootServices, HashSet<ServiceController> dependentServices)
        {
            try
            {
                LogColored("Ensuring all services are stopped before file deletion...", MessageType.Info);
                await StopServicesAsync(dependentServices);
                await StopServicesAsync(rootServices);
                await Task.Delay(1000);

                bool fileDeleted = false;
                int deleteAttempts = 10;
                int currentAttempt = 0;

                while (!fileDeleted && currentAttempt < deleteAttempts)
                {
                    currentAttempt++;

                    try
                    {
                        if (File.Exists(targetFile))
                        {
                            try
                            {
                                File.Delete(targetFile);
                                LogColored($"Successfully deleted file: {targetFile}", MessageType.Success);
                                fileDeleted = true;
                            }
                            catch (Exception deleteEx)
                            {
                                if (currentAttempt == 1)
                                {
                                    LogColored($"Failed to delete file (file may be in use): {deleteEx.Message}", MessageType.Warning);
                                }

                                if (currentAttempt < deleteAttempts)
                                {
                                    Console.WriteLine();
                                    LogColored($"Attempt {currentAttempt}/{deleteAttempts}: Finding and terminating processes holding the file...", MessageType.Warning);

                                    List<int> processIds = FindProcessesUsingFile(targetFile);

                                    if (processIds.Count > 0)
                                    {
                                        foreach (int pid in processIds)
                                        {
                                            if (KillProcess(pid))
                                            {
                                                LogColored($"  Terminated process {pid}", MessageType.Success);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        LogColored("No processes found holding the file, retrying services stop...", MessageType.Warning);
                                        await StopServicesAsync(dependentServices);
                                        await Task.Delay(500);
                                        await StopServicesAsync(rootServices);
                                    }

                                    await Task.Delay(500);
                                    Console.WriteLine();

                                    // Immediate retry after killing processes
                                    try
                                    {
                                        if (File.Exists(targetFile))
                                        {
                                            File.Delete(targetFile);
                                            LogColored($"Successfully deleted file on attempt {currentAttempt}: {targetFile}", MessageType.Success);
                                            fileDeleted = true;
                                        }
                                    }
                                    catch (Exception retryEx)
                                    {
                                        if (currentAttempt < deleteAttempts)
                                        {
                                            // Log to file only, will retry in next iteration
                                            try
                                            {
                                                File.AppendAllText(logFilePath, $"[{DateTime.Now:HH:mm:ss}] Attempt {currentAttempt} failed: {retryEx.Message}\n");
                                            }
                                            catch { }
                                        }
                                    }
                                }
                                else
                                {
                                    // Final attempt failed
                                    LogColored($"Failed to delete file after {deleteAttempts} attempts: {deleteEx.Message}", MessageType.Error);
                                    throw;
                                }
                            }
                        }
                        else
                        {
                            LogColored($"Target file not found: {targetFile}", MessageType.Warning);
                            fileDeleted = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        if (currentAttempt >= deleteAttempts)
                        {
                            throw;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogColored($"Error in DeleteTargetFileAsync: {ex.Message}", MessageType.Error);
                throw;
            }
        }

        /// <summary>
        /// Stops services and verifies they are stopped
        /// </summary>
        /// <param name="services">HashSet of services to stop</param>
        private static async Task StopServicesAsync(HashSet<ServiceController> services)
        {
            try
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
                            LogColored($"Stopping service: {service.ServiceName}", MessageType.Info);
                        }
                        else
                        {
                            LogColored($"Service already stopped: {service.ServiceName}", MessageType.Success);
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log the error but don't show it to console
                        try
                        {
                            File.AppendAllText(logFilePath, $"[{DateTime.Now:HH:mm:ss}] Failed to stop service \"{service.ServiceName}\": {ex.Message}\n");
                        }
                        catch { }
                        LogColored($"Waiting for service: {service.ServiceName}", MessageType.Warning);
                    }
                    await Task.Delay(100);
                }

                Console.WriteLine();

                // Verify all services are stopped
                bool allStopped = false;
                int maxRetries = 10;
                int retryCount = 0;
                int barWidth = 30;

                while (!allStopped && retryCount < maxRetries)
                {
                    allStopped = true;
                    await Task.Delay(500); // Wait for services to stop

                    List<string> stuckServices = new();

                    foreach (ServiceController service in services)
                    {
                        try
                        {
                            service.Refresh();
                            if (service.Status != ServiceControllerStatus.Stopped)
                            {
                                allStopped = false;
                                stuckServices.Add(service.ServiceName);

                                // Try stopping again
                                if (retryCount < maxRetries - 1)
                                {
                                    try
                                    {
                                        service.Stop();
                                    }
                                    catch (Exception ex)
                                    {
                                        // Log but don't console output
                                        try
                                        {
                                            File.AppendAllText(logFilePath, $"[{DateTime.Now:HH:mm:ss}] Retry stop failed for \"{service.ServiceName}\": {ex.Message}\n");
                                        }
                                        catch { }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            allStopped = false;
                            stuckServices.Add(service.ServiceName);
                            // Log the error
                            try
                            {
                                File.AppendAllText(logFilePath, $"[{DateTime.Now:HH:mm:ss}] Error checking service {service.ServiceName}: {ex.Message}\n");
                            }
                            catch { }
                        }
                    }

                    // Show animated progress bar for stuck services
                    if (!allStopped && stuckServices.Count > 0)
                    {
                        int position = retryCount % (barWidth * 2 - 2);
                        if (position >= barWidth)
                            position = barWidth * 2 - 2 - position;

                        string bar = new string('-', barWidth);
                        char[] barChars = bar.ToCharArray();
                        barChars[position] = '=';
                        if (position > 0) barChars[position - 1] = '=';

                        string animatedBar = new string(barChars);
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.Write($"\r  [{animatedBar}] Waiting for {stuckServices.Count} service(s) to stop... (Attempt {retryCount + 1}/{maxRetries})");
                        Console.ResetColor();
                    }

                    retryCount++;
                }

                Console.WriteLine();
                Console.WriteLine();

                if (allStopped)
                {
                    LogColored($"All {services.Count} services have been successfully stopped.", MessageType.Success);
                }
                else
                {
                    LogColored($"Warning: Some services did not stop after {maxRetries} retries.", MessageType.Warning);
                }
            }
            catch (Exception ex)
            {
                LogColored($"Error in StopServicesAsync: {ex.Message}", MessageType.Error);
            }
        }

        /// <summary>
        /// Starts services and verifies they are running
        /// </summary>
        /// <param name="services">HashSet of services to start</param>
        private static async Task StartServicesAsync(HashSet<ServiceController> services)
        {
            try
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
                            LogColored($"Starting service: {service.ServiceName}", MessageType.Info);
                        }
                        else
                        {
                            LogColored($"Service already running: {service.ServiceName}", MessageType.Success);
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log the error but don't show it to console
                        try
                        {
                            File.AppendAllText(logFilePath, $"[{DateTime.Now:HH:mm:ss}] Failed to start service \"{service.ServiceName}\": {ex.Message}\n");
                        }
                        catch { }
                        LogColored($"Waiting for service: {service.ServiceName}", MessageType.Warning);
                    }
                    await Task.Delay(100);
                }

                Console.WriteLine();

                // Verify all services are running
                bool allRunning = false;
                int maxRetries = 10;
                int retryCount = 0;
                while (!allRunning && retryCount < maxRetries)
                {
                    allRunning = true;
                    await Task.Delay(500); // Wait for services to start

                    List<string> stuckServices = new();

                    foreach (ServiceController service in services)
                    {
                        try
                        {
                            service.Refresh();
                            if (service.Status != ServiceControllerStatus.Running)
                            {
                                allRunning = false;
                                stuckServices.Add(service.ServiceName);

                                // Try starting again
                                if (retryCount < maxRetries - 1)
                                {
                                    try
                                    {
                                        service.Start();
                                    }
                                    catch (Exception ex)
                                    {
                                        // Log but don't console output
                                        try
                                        {
                                            File.AppendAllText(logFilePath, $"[{DateTime.Now:HH:mm:ss}] Retry start failed for \"{service.ServiceName}\": {ex.Message}\n");
                                        }
                                        catch { }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            allRunning = false;
                            stuckServices.Add(service.ServiceName);
                            // Log the error
                            try
                            {
                                File.AppendAllText(logFilePath, $"[{DateTime.Now:HH:mm:ss}] Error checking service {service.ServiceName}: {ex.Message}\n");
                            }
                            catch { }
                        }
                    }

                    // Show waiting animation for stuck services
                    if (!allRunning && stuckServices.Count > 0)
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.Write($"\r  Waiting for {stuckServices.Count} service(s) to start... (Attempt {retryCount + 1}/{maxRetries})");
                        Console.ResetColor();
                    }

                    retryCount++;
                }

                Console.WriteLine();
                Console.WriteLine();

                if (allRunning)
                {
                    LogColored($"All {services.Count} services have been successfully started.", MessageType.Success);
                }
                else
                {
                    LogColored($"Warning: Not all services are running after {maxRetries} retries.", MessageType.Warning);
                }
            }
            catch (Exception ex)
            {
                LogColored($"Error in StartServicesAsync: {ex.Message}", MessageType.Error);
            }
        }

        private static HashSet<ServiceController> GetDependentServices(HashSet<ServiceController> services)
        {
            HashSet<ServiceController> allDependents = new();
            Queue<ServiceController> toProcess = new(services);
            HashSet<string> processed = new();

            try
            {
                while (toProcess.Count > 0)
                {
                    var service = toProcess.Dequeue();

                    if (processed.Contains(service.ServiceName))
                        continue;

                    processed.Add(service.ServiceName);

                    try
                    {
                        foreach (var dependent in service.DependentServices)
                        {
                            if (!processed.Contains(dependent.ServiceName))
                            {
                                allDependents.Add(dependent);
                                toProcess.Enqueue(dependent);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogColored($"Error retrieving dependent services for {service.ServiceName}: {ex.Message}", MessageType.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                LogColored($"Error retrieving dependent services: {ex.Message}", MessageType.Error);
            }

            return allDependents;
        }

        private static HashSet<ServiceController> GetServices(string[] services)
        {
            HashSet<ServiceController> serviceControllers = new();

            foreach (string service in services)
            {
                try
                {
                    serviceControllers.Add(ServiceHelper.GetServiceController(service));
                }
                catch (Exception ex)
                {
                    LogColored($"Error retrieving service {service}: {ex.Message}", MessageType.Error);
                }
            }

            return serviceControllers;
        }

        /// <summary>
        /// Finds and terminates processes holding a file lock
        /// </summary>
        private static List<int> FindProcessesUsingFile(string filePath)
        {
            List<int> processList = new();

            try
            {
                Process[] processes = Process.GetProcesses();
                foreach (Process process in processes)
                {
                    try
                    {
                        if (process.Modules != null)
                        {
                            foreach (ProcessModule module in process.Modules)
                            {
                                if (module.FileName.Equals(filePath, StringComparison.OrdinalIgnoreCase))
                                {
                                    processList.Add(process.Id);
                                    break;
                                }
                            }
                        }
                    }
                    catch
                    {
                        // Silently skip processes we can't access
                    }
                }
            }
            catch (Exception ex)
            {
                LogColored($"Error enumerating processes: {ex.Message}", MessageType.Error);
            }

            return processList;
        }

        /// <summary>
        /// Kills a process by ID
        /// </summary>
        private static bool KillProcess(int processId)
        {
            try
            {
                Process process = Process.GetProcessById(processId);
                process.Kill();
                process.WaitForExit(1000);
                return true;
            }
            catch (Exception ex)
            {
                LogColored($"Failed to kill process {processId}: {ex.Message}", MessageType.Error);
                return false;
            }
        }
    }
}
