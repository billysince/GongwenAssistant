using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Extensibility;
using HarmonyLib;

namespace GongwenPatcher
{
    // ???: ???????? Local_Wps_Vsto.MyAddin ?? OnConnection??
    // WPS / Office Word ?? Addins ?????????????????????????? A_ ?????????
    // ProgId ?? "GongwenPatcher.Connect"??CLSID ???????????? install/uninstall??
    [Guid("9F1A2B3C-4D5E-6F70-8190-A1B2C3D4E5F6")]
    [ProgId("GongwenPatcher.Connect")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class Connect : IDTExtensibility2
    {
        private static bool s_patched = false;
        private static readonly object s_lock = new object();
        private static string s_logPath;

        // ????????, ????????, ?????
        private static void Log(string msg)
        {
            try
            {
                if (s_logPath == null)
                {
                    string dir = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "GongwenAssistant");
                    Directory.CreateDirectory(dir);
                    s_logPath = Path.Combine(dir, "patcher.log");
                }
                File.AppendAllText(s_logPath,
                    DateTime.Now.ToString("HH:mm:ss.fff") + " " + msg + Environment.NewLine);
            }
            catch { }
        }

        public void OnConnection(object Application, ext_ConnectMode ConnectMode,
            object AddInInst, ref Array custom)
        {
            try
            {
                Log("OnConnection start mode=" + ConnectMode);
                EnsurePatched();
                Log("OnConnection end");
            }
            catch (Exception ex)
            {
                Log("OnConnection EX " + ex);
            }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            Log("OnDisconnection " + RemoveMode);
        }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom)
        {
            // WPS 启动完成 → 之前 AddinsCL 里的 crash 计数是误报, 清掉避免累积 3 次后 wps 自动禁用
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                    @"Software\Kingsoft\Office\WPS\AddinsCL", writable: true))
                {
                    if (key != null)
                    {
                        int cleared = 0;
                        foreach (string name in key.GetValueNames())
                        {
                            if (name == "Local_Wps_Vsto.MyAddin" ||
                                name.StartsWith("A_GongwenPatcher", StringComparison.OrdinalIgnoreCase))
                            {
                                key.DeleteValue(name, false);
                                cleared++;
                            }
                        }
                        Log("OnStartupComplete cleared CL entries=" + cleared);
                    }
                }
            }
            catch (Exception ex)
            {
                Log("OnStartupComplete EX " + ex.Message);
            }
        }
        public void OnBeginShutdown(ref Array custom) { }

        // ????: ?? scan ?????? assemblies, ???? AssemblyLoad ??????????????
        internal static void EnsurePatched()
        {
            lock (s_lock)
            {
                if (s_patched) return;
                s_patched = true;

                AppDomain.CurrentDomain.AssemblyLoad += OnAssemblyLoaded;
                Log("AssemblyLoad subscribed");

                foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
                {
                    TryPatchAssembly(asm);
                }
            }
        }

        private static void OnAssemblyLoaded(object sender, AssemblyLoadEventArgs args)
        {
            try
            {
                TryPatchAssembly(args.LoadedAssembly);
            }
            catch (Exception ex)
            {
                Log("OnAssemblyLoaded EX " + ex.Message);
            }
        }

        // ???????: ??? Local_Wps_Vsto, ?? UserUtil.IsVip / HasLogin / NeedAccountValidate / NeedRegister ??? hook
        private static void TryPatchAssembly(Assembly asm)
        {
            if (asm == null) return;
            string name;
            try { name = asm.GetName().Name; } catch { return; }
            if (!string.Equals(name, "Local_Wps_Vsto", StringComparison.OrdinalIgnoreCase))
                return;

            Log("Local_Wps_Vsto detected: " + asm.FullName);
            try
            {
                Type t = asm.GetType("Local_Wps_Vsto.UserUtil", false);
                if (t == null)
                {
                    Log("UserUtil type NOT FOUND");
                    return;
                }
                Log("UserUtil OK");

                var harmony = new Harmony("io.github.billysince.gongwen.patcher");

                // ???? bool ????? -> Prefix ?? __result=true ??? return false (skip original)
                string[] returnTrue = new[] { "IsVip", "HasLogin", "NeedAccountValidate", "NeedRegister" };
                foreach (string mn in returnTrue)
                {
                    var m = t.GetMethod(mn,
                        BindingFlags.Public | BindingFlags.NonPublic |
                        BindingFlags.Static | BindingFlags.Instance);
                    if (m == null)
                    {
                        Log("  " + mn + " not found");
                        continue;
                    }
                    if (m.ReturnType != typeof(bool))
                    {
                        Log("  " + mn + " return is " + m.ReturnType.Name + ", skip");
                        continue;
                    }
                    var prefix = typeof(Patches).GetMethod("ReturnTruePrefix",
                        BindingFlags.Public | BindingFlags.Static);
                    harmony.Patch(m, prefix: new HarmonyMethod(prefix));
                    Log("  Patched " + mn);
                }
            }
            catch (Exception ex)
            {
                Log("TryPatchAssembly EX " + ex);
            }
        }
    }

    public static class Patches
    {
        // Harmony Prefix: ???? __result ???? false ??? skip ?????
        public static bool ReturnTruePrefix(ref bool __result)
        {
            __result = true;
            return false;
        }
    }
}
