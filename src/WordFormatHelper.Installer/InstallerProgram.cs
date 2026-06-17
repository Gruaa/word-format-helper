using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Threading;
using System.Windows.Forms;

namespace WordFormatHelper.Installer
{
    internal static class InstallerProgram
    {
        private const string AddInProgId = "WordFormatHelper.AddIn";
        private const string InstallDir = @"C:\Program Files\WordFormatHelper";
        private const string DllFileName = "WordFormatHelper.AddIn.dll";

        [STAThread]
        private static int Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            bool uninstall = args.Length > 0 &&
                (string.Equals(args[0], "/uninstall", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(args[0], "-uninstall", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(args[0], "/u", StringComparison.OrdinalIgnoreCase));

            if (!IsAdministrator())
            {
                MessageBox.Show("请以管理员身份运行此安装程序。", "需要管理员权限",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return 1;
            }

            try
            {
                if (uninstall)
                {
                    RunUninstall();
                    MessageBox.Show("格式助手已卸载。", "卸载完成",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    RunInstall();
                    MessageBox.Show(
                        "格式助手安装完成！\n\n" +
                        "请重启 Word（或重新打开 Word），即可在工具栏看到「格式助手」选项卡。",
                        "安装完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                return 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show((uninstall ? "卸载" : "安装") + "失败：" + ex.Message + "\n\n" + ex.StackTrace,
                    "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 2;
            }
        }

        private static bool IsAdministrator()
        {
            try
            {
                using (var identity = WindowsIdentity.GetCurrent())
                {
                    var principal = new WindowsPrincipal(identity);
                    return principal.IsInRole(WindowsBuiltInRole.Administrator);
                }
            }
            catch
            {
                return false;
            }
        }

        private static void RunInstall()
        {
            Directory.CreateDirectory(InstallDir);
            string dllPath = Path.Combine(InstallDir, DllFileName);

            ExtractEmbeddedDll(dllPath);

            string regasm32 = @"C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe";
            string regasm64 = @"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe";

            if (File.Exists(regasm64))
            {
                RunRegasm(regasm64, dllPath, register: true);
            }
            if (File.Exists(regasm32))
            {
                RunRegasm(regasm32, dllPath, register: true);
            }

            ClearWordResiliencyDisabledItems();
            RegisterWordAddin();
            RegisterUninstaller();
        }

        private static void ClearWordResiliencyDisabledItems()
        {
            // Word may have disabled the add-in due to a previous failed load.
            // Clear the DisabledItems list so Word will try to load it again.
            string[] resiliencyPaths = new string[]
            {
                @"SOFTWARE\Microsoft\Office\16.0\Word\Resiliency\DisabledItems",
                @"SOFTWARE\Microsoft\Office\15.0\Word\Resiliency\DisabledItems"
            };

            foreach (var path in resiliencyPaths)
            {
                try
                {
                    using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(path, true))
                    {
                        if (key != null)
                        {
                            foreach (var name in key.GetSubKeyNames())
                            {
                                try { key.DeleteSubKey(name, false); } catch { }
                            }
                        }
                    }
                }
                catch { }
            }
        }

        private static void RunUninstall()
        {
            string dllPath = Path.Combine(InstallDir, DllFileName);

            UnregisterWordAddin();
            UnregisterUninstaller();

            string regasm32 = @"C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe";
            string regasm64 = @"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe";

            if (File.Exists(dllPath))
            {
                if (File.Exists(regasm64))
                {
                    try { RunRegasm(regasm64, dllPath, register: false); } catch { }
                }
                if (File.Exists(regasm32))
                {
                    try { RunRegasm(regasm32, dllPath, register: false); } catch { }
                }
            }

            try
            {
                if (Directory.Exists(InstallDir))
                {
                    foreach (var f in Directory.GetFiles(InstallDir))
                    {
                        try { File.Delete(f); } catch { }
                    }
                    try { Directory.Delete(InstallDir, true); } catch { }
                }
            }
            catch { }
        }

        private static void ExtractEmbeddedDll(string targetPath)
        {
            var asm = Assembly.GetExecutingAssembly();
            var resourceName = "WordFormatHelper.AddIn.dll";
            using (var stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    throw new Exception("找不到内嵌的加载项 DLL 资源：" + resourceName);
                }
                using (var fs = File.Create(targetPath))
                {
                    stream.CopyTo(fs);
                }
            }
        }

        private static void RunRegasm(string regasmPath, string dllPath, bool register)
        {
            string args = register
                ? $"/codebase \"{dllPath}\""
                : $"/unregister \"{dllPath}\"";

            var psi = new ProcessStartInfo
            {
                FileName = regasmPath,
                Arguments = args,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using (var proc = Process.Start(psi))
            {
                proc.WaitForExit(30000);
                if (!proc.HasExited)
                {
                    proc.Kill();
                }
            }
        }

        private static void RegisterWordAddin()
        {
            string addinRoot = @"SOFTWARE\Microsoft\Office\Word\Addins\" + AddInProgId;

            // HKLM 64-bit (for all users, 64-bit Word)
            using (var hklm64 = Microsoft.Win32.RegistryKey.OpenBaseKey(
                Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry64))
            {
                try
                {
                    using (var key = hklm64.CreateSubKey(addinRoot))
                    {
                        key.SetValue("LoadBehavior", 3, Microsoft.Win32.RegistryValueKind.DWord);
                        key.SetValue("Description", "格式助手", Microsoft.Win32.RegistryValueKind.String);
                        key.SetValue("FriendlyName", "格式助手", Microsoft.Win32.RegistryValueKind.String);
                        key.SetValue("CommandLineSafe", 0, Microsoft.Win32.RegistryValueKind.DWord);
                    }
                }
                catch { }
            }

            // HKLM 32-bit (for all users, 32-bit Word)
            using (var hklm32 = Microsoft.Win32.RegistryKey.OpenBaseKey(
                Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry32))
            {
                try
                {
                    using (var key = hklm32.CreateSubKey(addinRoot))
                    {
                        key.SetValue("LoadBehavior", 3, Microsoft.Win32.RegistryValueKind.DWord);
                        key.SetValue("Description", "格式助手", Microsoft.Win32.RegistryValueKind.String);
                        key.SetValue("FriendlyName", "格式助手", Microsoft.Win32.RegistryValueKind.String);
                        key.SetValue("CommandLineSafe", 0, Microsoft.Win32.RegistryValueKind.DWord);
                    }
                }
                catch { }
            }

            // HKCU (current user - ensures Word auto-loads without admin issues)
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(addinRoot))
                {
                    key.SetValue("LoadBehavior", 3, Microsoft.Win32.RegistryValueKind.DWord);
                    key.SetValue("Description", "格式助手", Microsoft.Win32.RegistryValueKind.String);
                    key.SetValue("FriendlyName", "格式助手", Microsoft.Win32.RegistryValueKind.String);
                    key.SetValue("CommandLineSafe", 0, Microsoft.Win32.RegistryValueKind.DWord);
                }
            }
            catch { }
        }

        private static void UnregisterWordAddin()
        {
            string addinRoot = @"SOFTWARE\Microsoft\Office\Word\Addins\" + AddInProgId;

            using (var hklm64 = Microsoft.Win32.RegistryKey.OpenBaseKey(
                Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry64))
            {
                try { hklm64.DeleteSubKey(addinRoot, false); } catch { }
            }

            using (var hklm32 = Microsoft.Win32.RegistryKey.OpenBaseKey(
                Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry32))
            {
                try { hklm32.DeleteSubKey(addinRoot, false); } catch { }
            }

            try { Microsoft.Win32.Registry.CurrentUser.DeleteSubKey(addinRoot, false); } catch { }
        }

        private static void RegisterUninstaller()
        {
            string uninstallKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\WordFormatHelper";

            using (var hklm64 = Microsoft.Win32.RegistryKey.OpenBaseKey(
                Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry64))
            {
                try
                {
                    using (var key = hklm64.CreateSubKey(uninstallKey))
                    {
                        key.SetValue("DisplayName", "格式助手", Microsoft.Win32.RegistryValueKind.String);
                        key.SetValue("DisplayVersion", "1.0.0.0", Microsoft.Win32.RegistryValueKind.String);
                        key.SetValue("Publisher", "WordFormatHelper", Microsoft.Win32.RegistryValueKind.String);
                        key.SetValue("InstallLocation", InstallDir, Microsoft.Win32.RegistryValueKind.String);
                        key.SetValue("NoModify", 1, Microsoft.Win32.RegistryValueKind.DWord);
                        key.SetValue("NoRepair", 1, Microsoft.Win32.RegistryValueKind.DWord);

                        string thisExe = Assembly.GetExecutingAssembly().Location;
                        key.SetValue("UninstallString", "\"" + thisExe + "\" /uninstall", Microsoft.Win32.RegistryValueKind.String);
                        key.SetValue("DisplayIcon", thisExe, Microsoft.Win32.RegistryValueKind.String);
                    }
                }
                catch { }
            }
        }

        private static void UnregisterUninstaller()
        {
            string uninstallKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\WordFormatHelper";

            using (var hklm64 = Microsoft.Win32.RegistryKey.OpenBaseKey(
                Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry64))
            {
                try { hklm64.DeleteSubKey(uninstallKey, false); } catch { }
            }
        }
    }
}
