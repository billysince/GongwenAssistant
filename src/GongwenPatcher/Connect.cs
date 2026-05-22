using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Extensibility;
using HarmonyLib;

namespace GongwenPatcher
{
    [Guid("9F1A2B3C-4D5E-6F70-8190-A1B2C3D4E5F6")]
    [ProgId("GongwenPatcher.Connect")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class Connect : IDTExtensibility2
    {
        private static bool s_patched = false;
        private static readonly object s_lock = new object();
        private static string s_logPath;

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

                Type myAddin = asm.GetType("Local_Wps_Vsto.MyAddin", false);
                if (myAddin != null)
                {
                    int wrapped = 0;
                    var tracePrefix = typeof(Patches).GetMethod("TracePrefix",
                        BindingFlags.Public | BindingFlags.Static);
                    var swallowFinal = typeof(Patches).GetMethod("SwallowExceptionFinalizer",
                        BindingFlags.Public | BindingFlags.Static);

                    // v1.0.8: Postfix GetCustomUI to rewrite ribbon XML.
                    // WPS 12.1+ does NOT call GetLabel callback for tab name.
                    // Tab name comes from the XML returned by GetCustomUI.
                    // We replace getLabel="GetLabel" on the tab node with a static label.
                    var getCustomUI = myAddin.GetMethod("GetCustomUI",
                        BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                    if (getCustomUI != null && getCustomUI.ReturnType == typeof(string))
                    {
                        var gcuPostfix = typeof(Patches).GetMethod("GetCustomUIPostfix",
                            BindingFlags.Public | BindingFlags.Static);
                        harmony.Patch(getCustomUI, postfix: new HarmonyMethod(gcuPostfix));
                        Log("  Patched MyAddin.GetCustomUI (ribbon XML rewrite)");
                    }
                    else
                    {
                        Log("  GetCustomUI not found, tab name stays original");
                    }

                    // v1.0.8b: also Postfix GetLabel to override tab "gwgs" return value.
                    // WPS may still call GetLabel even after XML rewrite.
                    var getLabel = myAddin.GetMethod("GetLabel",
                        BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                    if (getLabel != null && getLabel.ReturnType == typeof(string))
                    {
                        var glPrefix = typeof(Patches).GetMethod("GetLabelPostfix",
                            BindingFlags.Public | BindingFlags.Static);
                        harmony.Patch(getLabel, prefix: new HarmonyMethod(glPrefix));
                        Log("  Patched MyAddin.GetLabel (postfix override for gwgs tab)");
                    }

                    foreach (var m in myAddin.GetMethods(BindingFlags.Public | BindingFlags.NonPublic |
                        BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly))
                    {
                        string nm = m.Name;
                        if (nm == "GetCustomUI" || nm == "GetLabel") continue;
                        bool wrap = nm.EndsWith("_Click") || nm.StartsWith("btn") ||
                                    nm.StartsWith("OnAction") ||
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
        public static bool ReturnTruePrefix(ref bool __result)
        {
            __result = true;
            return false;
        }

        // v1.0.8: Postfix for GetCustomUI ?? rewrite ribbon XML to change tab name.
        // Replace: tab id="gwgs" ... getLabel="GetLabel"
        // With:    tab id="gwgs" ... label="GongwenZhushou"
        // The actual Chinese characters are injected via Unicode to avoid encoding issues.
        public static void GetCustomUIPostfix(ref string __result)
        {
            if (__result == null) return;
            try
            {
                // "????????" = \u516c\u6587\u52a9\u624b
                string newLabel = "\u516c\u6587\u52a9\u624b";

                // The original XML has: <tab id="gwgs" insertBeforeMso="TabHome" getLabel="GetLabel">
                // We want:             <tab id="gwgs" insertBeforeMso="TabHome" label="????????">
                string old1 = "getLabel=\"GetLabel\"";
                string rep1 = "label=\"" + newLabel + "\"";

                if (__result.Contains(old1))
                {
                    // Only replace the FIRST occurrence (in the tab node, not buttons)
                    int idx = __result.IndexOf(old1);
                    __result = __result.Substring(0, idx) + rep1 + __result.Substring(idx + old1.Length);
                }

                // Also rename the group label "????????" (group gUser) if present
                // Already correct in v2, might be different in GAC original

                LogPatch("GetCustomUI XML rewritten, tab label -> " + newLabel);
            }
            catch (Exception ex)
            {
                LogPatch("GetCustomUI rewrite EX: " + ex.Message);
            }
        }

        // Prefix for GetLabel: if control.Id == "gwgs", override return value and skip original
        public static bool GetLabelPostfix(object __0, ref string __result)
        {
            try
            {
                if (__0 != null)
                {
                    var idProp = __0.GetType().GetProperty("Id");
                    if (idProp != null)
                    {
                        string id = idProp.GetValue(__0, null) as string;
                        if (id == "gwgs")
                        {
                            __result = "\u516c\u6587\u52a9\u624b";
                            return false; // skip original
                        }
                    }
                }
            }
            catch { }
            return true; // run original for other controls
        }

        private static void LogPatch(string msg)
        {
            try
            {
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "GongwenAssistant");
                string log = Path.Combine(dir, "patcher.log");
                File.AppendAllText(log,
                    DateTime.Now.ToString("HH:mm:ss.fff") + " " + msg + Environment.NewLine);
            }
            catch { }
        }

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
