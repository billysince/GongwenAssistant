using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using HarmonyLib;

namespace GongwenPatcher
{
    [ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    internal interface IDTExtensibility2
    {
        [DispId(1)] void OnConnection([In, MarshalAs(UnmanagedType.IDispatch)] object Application,
            [In] int ConnectMode,
            [In, MarshalAs(UnmanagedType.IDispatch)] object AddInInst,
            [In, MarshalAs(UnmanagedType.SafeArray)] ref Array custom);
        [DispId(2)] void OnDisconnection([In] int RemoveMode,
            [In, MarshalAs(UnmanagedType.SafeArray)] ref Array custom);
        [DispId(3)] void OnAddInsUpdate(
            [In, MarshalAs(UnmanagedType.SafeArray)] ref Array custom);
        [DispId(4)] void OnStartupComplete(
            [In, MarshalAs(UnmanagedType.SafeArray)] ref Array custom);
        [DispId(5)] void OnBeginShutdown(
            [In, MarshalAs(UnmanagedType.SafeArray)] ref Array custom);
    }

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

        public void OnConnection(object Application, int ConnectMode,
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

        public void OnDisconnection(int RemoveMode, ref Array custom)
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

                    // v1.00: deep diagnostics - check wordApp and wrap DocWpsUtil methods
                    var wordAppField = myAddin.GetField("wordApp",
                        BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static);
                    if (wordAppField != null)
                    {
                        object val = wordAppField.GetValue(null);
                        Log("  MyAddin.wordApp = " + (val == null ? "NULL" : val.GetType().FullName));
                    }
                    else
                    {
                        Log("  MyAddin.wordApp field NOT FOUND");
                    }

                    // Fix: btnInsert_Two/Three_Click has a bug where count==0
                    // causes sentences[0] COMException (COM is 1-based).
                    // Transpiler changes "if (count == 1)" to "if (count <= 1)".
                    var fixTranspiler = typeof(Patches).GetMethod("FixCountOneTranspiler",
                        BindingFlags.Public | BindingFlags.Static);
                    if (fixTranspiler != null)
                    {
                        string[] fixNames = new[] { "btnInsert_Two_Click", "btnInsert_Three_Click" };
                        foreach (string fn in fixNames)
                        {
                            var mf = myAddin.GetMethod(fn,
                                BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                            if (mf != null)
                            {
                                try
                                {
                                    harmony.Patch(mf, transpiler: new HarmonyMethod(fixTranspiler));
                                    Log("  Transpiler fix " + fn + " (count<=1 safe)");
                                }
                                catch (Exception ex)
                                {
                                    Log("  Transpiler fix failed " + fn + ": " + ex.Message);
                                }
                            }
                        }
                    }

                    // Deep diagnostics removed for performance (v1.01).
                    // DocWpsUtil/WpsHelper/MessageTipsUtil/CommonConfig no longer wrapped.
                    // These were only needed to debug #33 (count==1 COMException).
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
                string newLabel = "\u516c\u6587\u52a9\u624b 1.00";

                string old1 = "getLabel=\"GetLabel\"";
                string rep1 = "label=\"" + newLabel + "\"";

                if (__result.Contains(old1))
                {
                    // Only replace the FIRST occurrence (in the tab node, not buttons)
                    int idx = __result.IndexOf(old1);
                    __result = __result.Substring(0, idx) + rep1 + __result.Substring(idx + old1.Length);
                }

                // Hide WeChat customer service button (btnKeFu)
                string kefuId = "id=\"btnKeFu\"";
                if (__result.Contains(kefuId))
                {
                    int ki = __result.IndexOf(kefuId);
                    __result = __result.Insert(ki + kefuId.Length, " visible=\"false\"");
                }

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
                            __result = "\u516c\u6587\u52a9\u624b 1.00";
                            return false;
                        }
                        if (id == "btnUserMsg")
                        {
                            __result = "\u5df2\u6fc0\u6d3b";
                            return false;
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
                string nm = __originalMethod.Name;
                if (!nm.EndsWith("_Click") && !nm.StartsWith("btn") && !nm.StartsWith("OnAction"))
                    return true;

                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "GongwenAssistant");
                string log = Path.Combine(dir, "patcher.log");
                string extra = "";
                if (nm.EndsWith("_Click"))
                {
                    try
                    {
                        var waField = __originalMethod.DeclaringType.GetField("wordApp",
                            BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static);
                        if (waField != null)
                        {
                            object wa = waField.GetValue(null);
                            extra = " wordApp=" + (wa == null ? "NULL" : "OK");
                        }
                    }
                    catch { }
                }
                File.AppendAllText(log,
                    DateTime.Now.ToString("HH:mm:ss.fff") + " CLICK " +
                    __originalMethod.DeclaringType.Name + "." + __originalMethod.Name +
                    extra + Environment.NewLine);
            }
            catch { }
            return true;
        }

        // Transpiler: change "if (count == 1)" to "if (count <= 1)" in
        // btnInsert_Two_Click / btnInsert_Three_Click. Handles three IL patterns:
        //   A: ldc.i4.1; bne.un  ->  ldc.i4.1; bgt
        //   B: ldc.i4.1; ceq; brfalse  ->  ldc.i4.1; cgt; brtrue
        //   C: ldc.i4.1; beq  ->  ldc.i4.1; ble
        public static IEnumerable<CodeInstruction> FixCountOneTranspiler(
            IEnumerable<CodeInstruction> instructions)
        {
            var codes = new List<CodeInstruction>(instructions);
            int patchCount = 0;
            for (int i = 1; i < codes.Count; i++)
            {
                if (codes[i - 1].opcode == OpCodes.Ldc_I4_1)
                {
                    if (codes[i].opcode == OpCodes.Bne_Un || codes[i].opcode == OpCodes.Bne_Un_S)
                    {
                        codes[i].opcode = codes[i].opcode == OpCodes.Bne_Un_S ? OpCodes.Bgt_S : OpCodes.Bgt;
                        patchCount++;
                    }
                    else if (codes[i].opcode == OpCodes.Beq || codes[i].opcode == OpCodes.Beq_S)
                    {
                        codes[i].opcode = codes[i].opcode == OpCodes.Beq_S ? OpCodes.Ble_S : OpCodes.Ble;
                        patchCount++;
                    }
                    else if (codes[i].opcode == OpCodes.Ceq && i + 1 < codes.Count &&
                             (codes[i + 1].opcode == OpCodes.Brfalse || codes[i + 1].opcode == OpCodes.Brfalse_S))
                    {
                        codes[i].opcode = OpCodes.Cgt;
                        codes[i + 1].opcode = codes[i + 1].opcode == OpCodes.Brfalse_S ? OpCodes.Brtrue_S : OpCodes.Brtrue;
                        patchCount++;
                    }
                }
            }
            LogPatch("FixCountOneTranspiler patches=" + patchCount + " il_count=" + codes.Count);
            return codes;
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
