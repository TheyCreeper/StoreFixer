using StoreFixer.Utils;
using Microsoft.Win32;
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
        private static bool isSilent = false, isSetScheduledTaskOnCrash = false, noRestart = false, isEmergencyRestore = false;
        private const string AtlasTempKey = @"HKLM\SOFTWARE\AtlasOS\Temp";
        private const string AtlasTempSubKey = @"SOFTWARE\AtlasOS\Temp";
        private const string RecoveryKey = @"HKLM\SOFTWARE\AtlasOS\StoreFixerRecovery";
        private const string RecoverySubKey = @"SOFTWARE\AtlasOS\StoreFixerRecovery";
        private const string WinlogonKey = @"HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon";
        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
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
                ParseArguments(args);

                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                logFilePath = Path.Combine(desktopPath, "StoreFixer_Log.txt");
                LogColored(logFilePath);
                Clear();

                if (!IsRunAsTi())
                {
                    Console.WriteLine();
                    LogColored("ERROR: Not running as Trusted Installer", MessageType.Error);
                    LogColored("Please close this application and run it as Trusted Installer.", MessageType.Error);
                    Console.WriteLine();
                    return;
                }

                if (isEmergencyRestore)
                {
                    await EmergencyRestoreFromSafeMode();
                }
                else
                {
                    await Execution();
                }

                Console.WriteLine();
                LogColored("StoreFixer Execution Completed", MessageType.Header);
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                MarkRecoveryError();
                LogColored($"FATAL ERROR in Main: {ex.Message}", MessageType.Critical);
                LogColored($"Stack trace: {ex.StackTrace}", MessageType.Critical);
                Console.WriteLine();

                try
                {
                    await RestoreOnCrash();
                }
                catch { }
            }
            finally
            {
                if (isSilent) Environment.Exit(0);
                if (!isSilent)
                {
                    Console.WriteLine();
                    Console.ForegroundColor = ConsoleColor.Cyan;
                    Console.Write("Press any key to exit...");
                    Console.ResetColor();
                    try { Console.ReadKey(); } catch { }
                }
            }
        }
        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        private static void ParseArguments(string[] args)
        {
            isSilent = args.Any(x => x.Equals("silent", StringComparison.OrdinalIgnoreCase));
            isSetScheduledTaskOnCrash = args.Any(x => x.Equals("isSetScheduledTaskOnCrash", StringComparison.OrdinalIgnoreCase));
            noRestart = args.Any(x => x.Equals("noRestart", StringComparison.OrdinalIgnoreCase));
            isEmergencyRestore = args.Any(x =>
                x.Equals("emergencyRestore", StringComparison.OrdinalIgnoreCase) ||
                x.Equals("restore", StringComparison.OrdinalIgnoreCase) ||
                x.Equals("repair", StringComparison.OrdinalIgnoreCase));

            if (isSilent)
            {
                IntPtr hWnd = GetConsoleWindow();
                ShowWindow(hWnd, SW_HIDE);
            }
        }

        /// <summary>
        /// In case of a crash, sets a scheduled task which asks the user
        /// if they want to retry after a restart
        /// </summary>
        /// <returns></returns>
        private static async Task SetScheduledTask()
        {
            RegistryHelper.SetValue("HKCU\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\RunOnce",
                "MsStoreFix",
                "powershell -EP RemoteSigned -NoP & \"\"\"$([Environment]::GetFolderPath('Windows'))\\AtlasModules\\Scripts\\ScriptWrappers\\StoreFixerPrompt.ps1\"\"\"");
        }

        /// <summary>
        /// Emergency restoration if execution crashes
        /// </summary>
        private static async Task RestoreOnCrash()
        {
            if (isSetScheduledTaskOnCrash)
            {
                await SetScheduledTask();
            }

            Console.WriteLine();
            LogColored("╔════════════════════════════════════════════════════════════╗", MessageType.Critical);
            LogColored("║                  EMERGENCY RESTORATION                     ║", MessageType.Critical);
            LogColored("║                    PLEASE WAIT...                          ║", MessageType.Critical);
            LogColored("╚════════════════════════════════════════════════════════════╝", MessageType.Critical);
            Console.WriteLine();

            Dictionary<string, ServiceStartMode> restoreBackup = servicesBackup.Count > 0
                ? servicesBackup
                : ReadStoredServiceBackup();

            if (restoreBackup.Count > 0)
            {
                LogColored($"Restoring {restoreBackup.Count} services from backup...", MessageType.Warning);
                Console.WriteLine();
                await RestoreServicesAsync(restoreBackup);
                await Task.Delay(1000);
                Console.WriteLine();

                LogColored("Starting restored services...", MessageType.Warning);
                Console.WriteLine();
                await StartServicesByNameAsync(restoreBackup);
                ClearSafeModeRecovery();
                Console.WriteLine();
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
            ShowSupportMessage();
        }

        private static async Task EmergencyRestoreFromSafeMode()
        {
            Console.WriteLine();
            LogColored("╔════════════════════════════════════════════════════════════╗", MessageType.Critical);
            LogColored("║             SAFE MODE EMERGENCY RESTORATION                ║", MessageType.Critical);
            LogColored("║                    PLEASE WAIT...                          ║", MessageType.Critical);
            LogColored("╚════════════════════════════════════════════════════════════╝", MessageType.Critical);
            Console.WriteLine();

            Dictionary<string, ServiceStartMode> backup = ReadStoredServiceBackup();
            if (backup.Count == 0)
            {
                LogColored("No service backup was found. Applying conservative Store service defaults.", MessageType.Warning);
                backup = GetFallbackServiceBackup();
            }

            if (backup.Count == 0)
            {
                LogColored("No services were available to repair.", MessageType.Error);
                ShowSupportMessage();
                return;
            }

            await RestoreServicesAsync(backup);
            await StartServicesByNameAsync(backup);
            ClearSafeModeRecovery();

            LogColored("Safe Mode emergency restoration complete.", MessageType.Success);
            if (!noRestart)
            {
                LogColored("Restarting back to normal Windows in 10 seconds...", MessageType.Info);
                Process.Start("shutdown", "/r /t 10");
            }
        }

        private static void ConfigureSafeModeRecovery(Dictionary<string, ServiceStartMode> servicesStartMode, bool startWatchdog = true)
        {
            if (servicesStartMode.Count == 0)
            {
                throw new InvalidOperationException("No services were backed up. Refusing to continue.");
            }

            string windir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            string runAsTiPath = Path.Combine(windir, @"AtlasModules\Scripts\RunAsTI.cmd");
            string storeFixerPath = Path.Combine(windir, @"AtlasModules\Tools\StoreFixer.exe");
            string currentShell = (RegistryHelper.GetValue(WinlogonKey, "Shell") as string) ?? "explorer.exe";
            string recoveryShell = $"explorer.exe,cmd /c \"\"{runAsTiPath}\" \"{storeFixerPath}\" emergencyRestore -wait\"";

            try
            {
                RegistryHelper.SetValue(RecoveryKey, "Active", 1, RegistryValueKind.DWord);
                RegistryHelper.SetValue(RecoveryKey, "State", "Armed", RegistryValueKind.String);
                RegistryHelper.SetValue(RecoveryKey, "OriginalShell", currentShell, RegistryValueKind.String);
                RegistryHelper.SetValue(RecoveryKey, "StartedAtUtc", DateTime.UtcNow.ToString("O"), RegistryValueKind.String);
                RegistryHelper.SetValue(WinlogonKey, "Shell", recoveryShell, RegistryValueKind.String);

                RunProcessOrThrow("bcdedit", "/set {current} safeboot minimal");
                if (startWatchdog)
                {
                    StartInterruptionWatchdog();
                }
                LogColored("Safe Mode emergency recovery has been armed.", MessageType.Warning);
            }
            catch
            {
                RegistryHelper.SetValue(WinlogonKey, "Shell", currentShell, RegistryValueKind.String);
                try { RegistryHelper.DeleteKey(RecoveryKey); } catch { }
                throw;
            }
        }

        private static void ClearSafeModeRecovery()
        {
            try
            {
                string originalShell = (RegistryHelper.GetValue(RecoveryKey, "OriginalShell") as string) ?? "explorer.exe";
                RegistryHelper.SetValue(WinlogonKey, "Shell", originalShell, RegistryValueKind.String);
            }
            catch (Exception ex)
            {
                LogColored($"Failed to restore Winlogon shell: {ex.Message}", MessageType.Error);
            }

            try
            {
                RunProcessOrLog("bcdedit", "/deletevalue {current} safeboot");
                RunProcess("bcdedit", "/deletevalue {current} safebootalternateshell");
            }
            catch { }

            try
            {
                RegistryHelper.DeleteKey(RecoveryKey);
            }
            catch { }

            LogColored("Safe Mode emergency recovery has been cleared.", MessageType.Success);
        }

        private static void MarkRecoveryError()
        {
            try
            {
                RegistryHelper.SetValue(RecoveryKey, "State", "Error", RegistryValueKind.String);
            }
            catch { }
        }

        private static void StartInterruptionWatchdog()
        {
            int pid = Process.GetCurrentProcess().Id;
            string command = $@"
$pidToWatch = {pid}
while (Get-Process -Id $pidToWatch -ErrorAction SilentlyContinue) {{
    Start-Sleep -Seconds 2
}}
$state = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\AtlasOS\StoreFixerRecovery' -Name State -ErrorAction SilentlyContinue).State
if ($state -eq 'Armed') {{
    shutdown /r /t 10 /c 'StoreFixer was interrupted. Rebooting into Safe Mode recovery.'
}}
";

            string encodedCommand = Convert.ToBase64String(Encoding.Unicode.GetBytes(command));
            Process.Start(new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -WindowStyle Hidden -EncodedCommand {encodedCommand}",
                UseShellExecute = false,
                CreateNoWindow = true
            });
        }

        private static void ShowSupportMessage()
        {
            Console.WriteLine();
            LogColored("StoreFixer hit an error and will not restart automatically.", MessageType.Critical);
            LogColored("Please contact support in the Atlas Discord and attach StoreFixer_Log.txt from your Desktop.", MessageType.Critical);
            Console.WriteLine();
        }

        private static Dictionary<string, ServiceStartMode> ReadStoredServiceBackup()
        {
            Dictionary<string, ServiceStartMode> backup = new(StringComparer.OrdinalIgnoreCase);

            try
            {
                using RegistryKey baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
                using RegistryKey? key = baseKey.OpenSubKey(AtlasTempSubKey, false);
                if (key is null)
                {
                    return backup;
                }

                foreach (string valueName in key.GetValueNames())
                {
                    object? value = key.GetValue(valueName);
                    if (TryParseServiceStartMode(value, out ServiceStartMode mode) && ServiceExists(valueName))
                    {
                        backup[valueName] = mode;
                    }
                }
            }
            catch (Exception ex)
            {
                LogColored($"Failed to read stored service backup: {ex.Message}", MessageType.Error);
            }

            return backup;
        }

        private static Dictionary<string, ServiceStartMode> GetFallbackServiceBackup()
        {
            string[] fallbackServices =
            {
                "Appinfo",
                "ClipSVC",
                "AppXSvc",
                "StateRepository",
                "InstallService",
                "LicenseManager"
            };

            Dictionary<string, ServiceStartMode> backup = new(StringComparer.OrdinalIgnoreCase);
            foreach (string serviceName in fallbackServices)
            {
                if (ServiceExists(serviceName))
                {
                    backup[serviceName] = ServiceStartMode.Manual;
                }
            }

            return backup;
        }

        private static bool TryParseServiceStartMode(object? value, out ServiceStartMode mode)
        {
            mode = ServiceStartMode.Manual;

            if (value is int intValue && Enum.IsDefined(typeof(ServiceStartMode), intValue))
            {
                mode = (ServiceStartMode)intValue;
                return true;
            }

            if (value is string stringValue && Enum.TryParse(stringValue, ignoreCase: true, out mode))
            {
                return true;
            }

            return false;
        }

        private static bool ServiceExists(string serviceName)
        {
            try
            {
                _ = ServiceHelper.GetStartupType(serviceName);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static async Task StartServicesByNameAsync(Dictionary<string, ServiceStartMode> servicesStartMode)
        {
            HashSet<ServiceController> services = new();

            foreach (KeyValuePair<string, ServiceStartMode> kvp in servicesStartMode)
            {
                if (kvp.Value == ServiceStartMode.Disabled)
                {
                    continue;
                }

                try
                {
                    services.Add(ServiceHelper.GetServiceController(kvp.Key));
                }
                catch (Exception ex)
                {
                    LogColored($"Error retrieving service {kvp.Key}: {ex.Message}", MessageType.Error);
                }
            }

            await StartServicesAsync(services);
        }

        private static void RunProcessOrThrow(string fileName, string arguments)
        {
            using Process process = RunProcess(fileName, arguments);

            if (process.ExitCode != 0)
            {
                throw new InvalidOperationException($"{fileName} {arguments} failed with exit code {process.ExitCode}.");
            }
        }

        private static Process RunProcess(string fileName, string arguments)
        {
            Process process = Process.Start(new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments,
                UseShellExecute = false,
                CreateNoWindow = true
            }) ?? throw new InvalidOperationException($"Failed to start {fileName}.");

            process.WaitForExit();
            return process;
        }

        private static void RunProcessOrLog(string fileName, string arguments)
        {
            try
            {
                RunProcessOrThrow(fileName, arguments);
            }
            catch (Exception ex)
            {
                LogColored($"{fileName} {arguments}: {ex.Message}", MessageType.Warning);
            }
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
            if (!isSilent)
            {
                Console.ForegroundColor = color;
                Console.WriteLine(timestampedMessage);
                Console.ResetColor();
            }

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
                // ====================================================================
                // PHASE 1: SERVICE BACKUP
                // ====================================================================
                LogColored("Saving service backup...");

                await CreateServiceBackup();

                // ====================================================================
                // PHASE 2: SERVICE RETRIEVAL & PREPARATION
                // ====================================================================
                LogColored("Getting services...");
                Console.WriteLine();

                string targetFile = $"{Environment.GetEnvironmentVariable("SYSTEMDRIVE")}\\ProgramData\\Microsoft\\Windows\\AppRepository\\StateRepository-Deployment.srd";

                try
                {
                    HashSet<ServiceController> rootServices = GetServices(["ClipSVC", "AppXSvc", "StateRepository"]);
                    HashSet<ServiceController> dependentServices = GetDependentServices(rootServices);

                    Dictionary<string, ServiceStartMode> servicesStartMode = new();
                    LogColored("Backing up root services...", MessageType.Header);
                    foreach (ServiceController serviceController in rootServices)
                    {
                        try
                        {
                            if (!servicesStartMode.ContainsKey(serviceController.ServiceName))
                            {
                                servicesStartMode.Add(serviceController.ServiceName, serviceController.StartType);
                            }
                            LogColored($"{serviceController.ServiceName} ({serviceController.StartType})", MessageType.Info);
                        }
                        catch (Exception ex)
                        {
                            LogColored($"{serviceController.ServiceName}: {ex.Message}", MessageType.Error);
                        }
                    }
                    Console.WriteLine();

                    LogColored("Backing up dependent services...", MessageType.Header);
                    foreach (ServiceController serviceController in dependentServices)
                    {
                        try
                        {
                            if (!servicesStartMode.ContainsKey(serviceController.ServiceName))
                            {
                                servicesStartMode.Add(serviceController.ServiceName, serviceController.StartType);
                            }
                            LogColored($"{serviceController.ServiceName} ({serviceController.StartType})", MessageType.Info);
                        }
                        catch (Exception ex)
                        {
                            LogColored($"{serviceController.ServiceName}: {ex.Message}", MessageType.Error);
                        }
                    }

                    foreach (KeyValuePair<string, ServiceStartMode> kvp in servicesStartMode)
                    {
                        try
                        {
                            RegistryHelper.SetValue(AtlasTempKey, kvp.Key, kvp.Value.ToString(), Microsoft.Win32.RegistryValueKind.String);
                        }
                        catch (Exception ex)
                        {
                            LogColored($"Failed to save backup for {kvp.Key}: {ex.Message}", MessageType.Error);
                        }
                    }
                    ConfigureSafeModeRecovery(servicesStartMode);
                    Console.WriteLine();

                    // ================================================================
                    // PHASE 3: DISABLING SERVICES
                    // ================================================================
                    LogColored("Disabling services: ", MessageType.Header);
                    LogColored("Disabling dependent services...", MessageType.Header);
                    await DisableServicesAsync(dependentServices);
                    await Task.Delay(1000);

                    LogColored("Disabling root services...", MessageType.Header);
                    await DisableServicesAsync(rootServices);
                    await Task.Delay(1000);
                    Console.WriteLine();

                    // ================================================================
                    // PHASE 4: STOPPING SERVICES
                    // ================================================================
                    LogColored("Stopping Services", MessageType.Header);
                    LogColored("Stopping dependent services...", MessageType.Header);
                    await StopServicesAsync(dependentServices);
                    await Task.Delay(1000);

                    LogColored("Stopping root services...", MessageType.Header);
                    await StopServicesAsync(rootServices);
                    await Task.Delay(2000);
                    Console.WriteLine();

                    // ================================================================
                    // PHASE 5: FILE DELETION
                    // ================================================================
                    LogColored("Deleting Target File", MessageType.Header);
                    Console.WriteLine();

                    await DeleteTargetFileAsync(targetFile, rootServices, dependentServices);
                    await Task.Delay(1000);
                    Console.WriteLine();

                    // ================================================================
                    // PHASE 6: SERVICE RESTORATION
                    // ================================================================
                    LogColored("Restoring Services to Original State", MessageType.Header);
                    Console.WriteLine();

                    await RestoreServicesAsync(servicesStartMode);
                    await Task.Delay(1000);
                    Console.WriteLine();

                    // ================================================================
                    // PHASE 7: STARTING SERVICES
                    // ================================================================
                    LogColored("Starting services", MessageType.Header);
                    Console.WriteLine();

                    LogColored("Starting dependent services...", MessageType.Header);
                    await StartServicesAsync(dependentServices);
                    await Task.Delay(1000);
                    Console.WriteLine();

                    LogColored("Starting root services...", MessageType.Header);
                    await StartServicesAsync(rootServices);
                    await Task.Delay(1000);
                    ClearSafeModeRecovery();
                    Console.WriteLine();
                    Console.WriteLine();
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
                        if (!servicesBackup.ContainsKey(service.ServiceName))
                        {
                            servicesBackup.Add(service.ServiceName, service.StartType);
                        }
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
                int maxRetries = 5;
                int retryCount = 0;

                while (!allDisabled && retryCount < maxRetries)
                {
                    LogColored($"Attemp {retryCount + 1}/{maxRetries}: Verifying services state..", MessageType.Info);
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
                int maxRetries = 5;
                int retryCount = 0;

                while (!allRestored && retryCount < maxRetries)
                {
                    LogColored($"Attemp {retryCount + 1}/{maxRetries}: Verifying service restoration...", MessageType.Info);
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
                int deleteAttempts = 5;
                int currentAttempt = 0;

                while (!fileDeleted && currentAttempt < deleteAttempts)
                {
                    LogColored($"Attemp {currentAttempt + 1}/{deleteAttempts}: Attempting to delete file...", MessageType.Info);
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
                    catch
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
                // Stop all services
                foreach (ServiceController service in services)
                {
                    try
                    {
                        service.Refresh();
                        if (service.Status != ServiceControllerStatus.Stopped)
                        {
                            service.Stop();
                        }
                        else
                        {
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
                // Verify all services are stopped
                bool allStopped = false;
                int maxRetries = 5;
                int retryCount = 0;
                int barWidth = 30;

                while (!allStopped && retryCount < maxRetries)
                {
                    LogColored($"Attemp {retryCount + 1}/{maxRetries}: Attempting to stop services...", MessageType.Info);
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
                        string fullServiceList = string.Join(",", stuckServices);
                        Console.Write($"\r  [{animatedBar}] Waiting for {stuckServices.Count} service(s) ({fullServiceList}) to stop... (Attempt {retryCount + 1}/{maxRetries})");
                        Console.ResetColor();
                    }

                    retryCount++;
                }

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

        private static void Clear()
        {
            try
            {
                Console.Clear();
            }
            catch
            {
                // Silently fail if console is hidden or unavailable
            }

            if (!isSilent)
            {
                Console.WriteLine("Project provided under CC0-Universal: https://github.com/TheyCreeper/StoreFixer\n\n");
            }
        }
    }
}
