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
            // WPS ??????? ?? ?? AddinsCL ??? crash ????????, ?????????? 3 ?¦Ê? wps ???????
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

                // bool returnTrue methods -> Prefix overrides __result=true and skips original
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

                // === click trace + exception finalizer (v1.0.6) ===
                // Wrap all MyAddin._Click handlers + ribbon callbacks (Get*, On*) with:
                //   - Prefix: log "CLICK <method>" so we can prove whether onAction was actually invoked
                //   - Finalizer: swallow any exception so a crash inside handler does NOT trigger
                //                WPS crash recovery dialog
                Type myAddin = asm.GetType("Local_Wps_Vsto.MyAddin", false);
                if (myAddin != null)
                {
                    int wrapped = 0;
                    var tracePrefix = typeof(Patches).GetMethod("TracePrefix",
                        BindingFlags.Public | BindingFlags.Static);
                    var swallowFinal = typeof(Patches).GetMethod("SwallowExceptionFinalizer",
                        BindingFlags.Public | BindingFlags.Static);
                    foreach (var m in myAddin.GetMethods(BindingFlags.Public | BindingFlags.NonPublic |
                        BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly))
                    {
                        string nm = m.Name;
                        bool wrap = nm.EndsWith("_Click") || nm.StartsWith("btn") ||
                                    nm.StartsWith("OnAction") || nm.StartsWith("GetLabel") ||
                                    nm.StartsWith("GetVisible") || nm.StartsWith("GetImage") ||
                                    nm.StartsWith("GetEnabled") || nm.StartsWith("OnLoad");
                        if (!wrap) continue;
                        if (m.IsAbstract || m.ContainsGenericParameters) continue;
                        try
                        {
                            harmony.Patch(m,
                                prefix: new HarmonyMethod(tracePrefix),
                                finalizer: new HarmonyMethod(swallowFinal));
                            wrapped++;
                        }
                        catch (Exception ex)
                        {
                            Log("  wrap " + nm + " failed: " + ex.Message);
                        }
                    }
                    Log("  wrapped " + wrapped + " MyAddin methods (trace + exception swallow)");
                }
                else
                {
                    Log("MyAddin type NOT FOUND, skip click-trace wrap");
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
        // Harmony Prefix: __result=true and return false skips original method body
        public static bool ReturnTruePrefix(ref bool __result)
        {
            __result = true;
            return false;
        }

        // Harmony Prefix: log every invocation of wrapped methods. Returns true so original runs.
        public static bool TracePrefix(MethodBase __originalMethod)
        {
            try
            {
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "GongwenAssistant");
                string log = Path.Combine(dir, "patcher.log");
                File.AppendAllText(log,
                    DateTime.Now.ToString("HH:mm:ss.fff") + " CLICK " +
                    __originalMethod.DeclaringType.Name + "." + __originalMethod.Name +
                    Environment.NewLine);
            }
            catch { }
            return true;
        }

        // Harmony Finalizer: swallow any exception so a crash inside a click handler
        // does NOT trigger WPS crash recovery dialog. Returning null clears __exception.
        public static Exception SwallowExceptionFinalizer(MethodBase __originalMethod, Exception __exception)
        {
            if (__exception != null)
            {
                try
                {
                    string dir = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "GongwenAssistant");
                    string log = Path.Combine(dir, "patcher.log");
                    File.AppendAllText(log,
                        DateTime.Now.ToString("HH:mm:ss.fff") + " CLICK_EX " +
                        __originalMethod.DeclaringType.Name + "." + __originalMethod.Name + ": " +
                        __exception.GetType().Name + " " + __exception.Message +
                        Environment.NewLine);
                }
                catch { }
            }
            return null;
        }
    }
}
