# Patcher 方案 — 运行时 IL 注入路线（推荐）

## 这是什么

在原版 `Local_Wps_Vsto.dll`（1,408,000 字节、强名 PKT = `475a36b05c42bd98`）**字节零修改**的前提下，通过一个独立的 WPS Word COM Add-in `A_GongwenPatcher.Connect`，在 WPS 进程内用 [Harmony](https://github.com/pardeike/Harmony) 把 `Local_Wps_Vsto.UserUtil.IsVip` / `HasLogin` 的 IL 字节码在 JIT 出来的本机代码层替换为 `return true`，绕开"此功能仅限 VIP 用户使用"的拦截。

与之前的 B2 路线（反编译后重新编译成弱名 dll、由 HKCU + CodeBase 加载）相比，Patcher 路线的优势：

- 原版 dll 不动，PublicKeyToken / 强名签名都保留；
- 不需要在自家仓库里维护一份 1.4 MB 的二进制；
- 升级原版时只需重装原版 install.bat 即可，Patcher 自动接管；
- 装/卸 Patcher 都是纯加法 / 纯减法，对原版生态零侵入；
- 全程 HKCU 注册，不需要管理员权限，不写 HKLM / GAC。

---

## 文件构成

```
src/GongwenPatcher/
├── Connect.cs               # COM Add-in 入口 (IDTExtensibility2)
├── Extensibility.cs         # IDTExtensibility2 接口最小内嵌 (不依赖 Office PIA)
├── AssemblyInfo.cs
├── build.ps1                # csc.exe (.NET Framework 4.x) 单次编译
├── install_patcher.ps1      # HKCU 全套注册 (CLSID + ProgId + Word\Addins + AddinsWL)
├── uninstall_patcher.ps1    # 对应反向卸载
├── snap.ps1                 # 测试用全屏截图
├── tools/Click.exe          # 测试用鼠标坐标点击工具
└── lib/0Harmony.dll         # Harmony 2.3.6 Fat 单文件 (含所有依赖, 2.2 MB)
```

部署后落地路径：

```
%LOCALAPPDATA%\GongwenAssistant\Patcher\
├── GongwenPatcher.dll       # 8.7 KB, 弱名, x86
├── 0Harmony.dll             # 2.2 MB, Fat 版
└── patcher.log              # 运行日志
```

---

## 关键决策点

### 为什么用 Harmony 而不是手写 detour

GitHub 上 .NET 运行时 method hook 的成熟方案：

| 库 | 星数 | 状态 | 评估 |
|---|---|---|---|
| `pardeike/Harmony` | ~3k | 持续维护 | Unity / RimWorld 等大量游戏用，文档全 |
| `MonoMod/MonoMod` | ~1k | 持续维护 | Harmony 的底层之一，更接近 raw detour |
| `EasyHook/EasyHook` | 老牌 | 半弃 | C++ for-managed，集成复杂 |

选 Harmony 2.3.6 Fat（最后一个没有把 MonoMod.Backports 单独拆出的小版本，单文件 2.2 MB 含全部依赖）。

> 注意：Harmony 2.4.x / 2.3.x Thin 版本把 `MonoMod.Backports.dll` 单独抽出，部署时容易缺依赖导致 `FileNotFoundException`。生产环境用 Fat 版本最稳妥。

### 为什么用独立的 WPS Word Add-in 而不是 inject 进程

可选注入方式：

1. **CreateRemoteThread + LoadLibrary** — 需要 admin，跨进程 IPC 复杂；
2. **AppInit_DLLs** — 全局生效，会影响所有进程，副作用过大；
3. **独立 Office Word Add-in** — WPS 自己会调用 `CoCreateInstance` 加载，不需要 admin；只需要 ProgId 字典序在原插件之前即可保证优先加载。

选 3。`A_GongwenPatcher.Connect` 的 `A_` 前缀确保它字典序排在 `Local_Wps_Vsto.MyAddin` 之前。

### 为什么 `AppDomain.AssemblyLoad` 是关键

Patcher 的 OnConnection 触发时，`Local_Wps_Vsto.dll` 还没被加载（WPS 是串行加载 add-ins）。所以不能在 OnConnection 当下立即反射查找 `UserUtil.IsVip`。

解决：

```csharp
public void OnConnection(...) {
    EnsurePatched();
}

internal static void EnsurePatched() {
    // 1) 先扫一遍当前已加载的 assemblies (兜底, 万一原插件先被加载)
    foreach (var asm in AppDomain.CurrentDomain.GetAssemblies()) {
        TryPatchAssembly(asm);
    }
    // 2) 订阅事件, 等 Local_Wps_Vsto 加载触发
    AppDomain.CurrentDomain.AssemblyLoad += OnAssemblyLoaded;
}
```

`AssemblyLoad` 事件在 CLR 完成 PE map + 静态域初始化之前**同步**触发；Harmony Prefix 此时打上即生效。后续业务代码任何路径走到 `IsVip()`，都被替换实现接管，返回 true。

### 为什么必须同时写 64-bit 和 32-bit 注册表视图

WPS Office 12.x 的 `wps.exe` 是 **32-bit** 进程。64-bit Windows 上的 32-bit 进程访问 `HKCU\Software\Classes` 会被注册表重定向器自动转向 `HKCU\Software\Classes\Wow6432Node`（与 HKLM 一致）。

PowerShell 64-bit 默认把 `New-Item HKCU:\Software\Classes\...` 写到 64-bit 视图。如果只写一份，32-bit WPS 调 `CoCreateInstance` 会得到 `REGDB_E_CLASSNOTREG (0x80040154)`。

修复：两个视图都写。`install_patcher.ps1::Write-ComRegistration` 接受一个 `$root` 参数，被分别调用 `HKCU:\Software\Classes` 和 `HKCU:\Software\Classes\Wow6432Node` 两次。

### 为什么需要 OnStartupComplete 清 AddinsCL

WPS 在自己的 `HKCU\Software\Kingsoft\Office\WPS\AddinsCL` 下记录"曾经 crash 的 add-in 计数"。即使 add-in 实际工作正常，WPS 仍会因为某些启动期短暂的 native 异常（Harmony 在做 IL replacement 时短暂的 SEH 触发）误判 Local_Wps_Vsto.MyAddin "crash 一次"。计数累积到 3 后，WPS 会在下次启动时弹出"是否禁用此加载项"对话框并阻塞所有 add-in 加载流程。

修复：Patcher 在 `OnStartupComplete`（WPS 完成所有 add-in 加载、视为"启动成功"时）主动把 `AddinsCL` 中跟自己 / Local_Wps_Vsto 相关的计数清零。这样误报永远累积不到 2 次。

---

## 安装 / 卸载

```powershell
cd D:\工作\20260520\GongwenAssistant\src\GongwenPatcher

# 一次性编译
PowerShell -ExecutionPolicy Bypass -File .\build.ps1

# 安装 (会先关掉所有 wps 进程)
PowerShell -ExecutionPolicy Bypass -File .\install_patcher.ps1

# 卸载
PowerShell -ExecutionPolicy Bypass -File .\uninstall_patcher.ps1
```

---

## 自检

启动 WPS 打开任意 docx 文档，确认：

1. Ribbon 顶部出现「公文助手」选项卡；
2. 切到该选项卡，点击「智能公文」（或其他 VIP 锁定按钮）；
3. 中央应出现绿色提示「? 已进入智能公文模式」而不是「此功能仅限 VIP 用户使用」；
4. 查 `%LOCALAPPDATA%\GongwenAssistant\patcher.log` 应有这样的行：

```
OnConnection start mode=ext_cm_Startup
AssemblyLoad subscribed
OnConnection end
Local_Wps_Vsto detected: Local_Wps_Vsto, Version=1.0.0.0, Culture=neutral, PublicKeyToken=475a36b05c42bd98
UserUtil OK
  Patched IsVip
  Patched HasLogin
OnStartupComplete cleared CL entries=1
```

只要看到 `Patched IsVip` / `Patched HasLogin`，方案就生效。
