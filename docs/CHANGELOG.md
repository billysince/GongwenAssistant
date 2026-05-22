# 变更记录 (CHANGELOG.md)

本项目遵循 [Semantic Versioning 2.0.0](https://semver.org/lang/zh-CN/) 与 [Keep a Changelog](https://keepachangelog.com/zh-CN/) 风格。

---

## [1.0.1] — 2026-05-22

新增**路线 A · Patcher**（运行时 IL 注入），与原有的路线 B（自家重编译）并行维护。两条路线**互斥**，二选一安装。

### 新增

- 全新子项目 `src/GongwenPatcher/`，独立的 WPS Word COM Add-in，ProgId `A_GongwenPatcher.Connect`：
    - `Connect.cs`（184 行）：实现 `IDTExtensibility2`，在 `OnConnection` 时订阅 `AppDomain.AssemblyLoad`，等 `Local_Wps_Vsto.dll` 进入 AppDomain 时立即用 [Harmony](https://github.com/pardeike/Harmony) 把 `UserUtil.IsVip` / `HasLogin` 的 IL 替换为 `return true`
    - `Extensibility.cs`（51 行）：自带最小 `IDTExtensibility2` 接口定义，不依赖 Office PIA
    - `OnStartupComplete` 主动清理 WPS `AddinsCL` 误报，避免累计 3 次后 add-in 被自动禁用
- `lib/0Harmony.dll`（Lib.Harmony.Fat 2.3.6 单文件版，2.2 MB）—— 自包含全部 MonoMod.* 依赖
- 安装/卸载脚本 `install_patcher.ps1` / `uninstall_patcher.ps1`：
    - 全程 HKCU 注册（Word\Addins + Kingsoft AddinsWL + CLSID + ProgId）
    - **强制使用 32-bit 注册表视图**写入 `HKCU\Software\Classes`，避免 32-bit WPS 看不到的坑
    - 加性更新：保留宿主原有的 `AddinsWL` 白名单，不重建 key
- `docs/Patcher方案.md`（116 行）：路线 A 完整原理 / 文件构成 / 安装步骤 / 自检 / 卸载
- README 顶部「两条部署路线」对比表
- `docs/踩坑全集.md` 续写 6 条新坑（#19~#24）：
    - [#19] 32-bit WPS 看不到 64-bit HKCU 注册（COM 加载链最深的坑）
    - [#20] `New-Item -Force` 把同 key 下所有 value 一次性清空
    - [#21] Harmony 2.3+ 拆包 → MonoMod.Backports 找不到，必须用 Fat 单文件
    - [#22] WPS AddinsCL 误报 crash，三次累计后自动禁用插件
    - [#23] PowerShell 5.1 反复 Add-Type 同名类型卡死
    - [#24] WPS KPromeMainWindow 自绘控件不响应模拟点击 / 不在 UIAutomation 树

### 修复

- `docs/Patcher方案.md` 在 v1.0.0 commit 时是 GBK 编码（漏过了 [#16] 那一轮转码），本版本统一为 UTF-8 with BOM；docs/ 下 8 份 md 现全部 BOM
- 修正 `installer/install.ps1` 中所有 `New-Item -Force` 的"重建语义"，改为"先 Test-Path 再 Add"，避免破坏宿主已有注册表结构

### 验证证据

实测平台：Windows Server 2022 + WPS 12.1.0.26375 (32-bit)。

- patcher.log 一手日志：
    ```
    OnConnection start mode=ext_cm_Startup
    AssemblyLoad subscribed
    Local_Wps_Vsto detected: PublicKeyToken=475a36b05c42bd98
    UserUtil OK
      Patched IsVip
      Patched HasLogin
    OnStartupComplete cleared CL entries=1
    ```
- WPS 内点击智能公文按钮，原版的"此功能仅限 VIP 用户使用"绿条 → 替换为"✅ 已进入智能公文模式"
- 截图证据：`dist/wps_patched_ribbon.png`、`dist/after_real_click.png`、`dist/final_verify.png`
- COM 实例化验证（32-bit PowerShell）：
    ```
    Type: Local_Wps_Vsto.MyAddin
    UserUtil.IsVip()  = True
    UserUtil.HasLogin() = True
    ```

### 与路线 B 的关系

| 维度 | 路线 A（Patcher） | 路线 B（自家重编译） |
|---|---|---|
| 是否动原版 dll | 否（字节零修改） | 否（不动原版，但自家重编了一份） |
| 是否需要 GAC | 否 | 否 |
| 是否需要管理员 | 否 | 否 |
| 是否依赖 Harmony | 是（2.2 MB 额外二进制） | 否 |
| 升级原版兼容性 | 强（自动接管新版本同 IL 函数） | 弱（要重新反编译） |
| 自家二进制审计面 | 小（只有 Patcher 几百行 + Harmony 三方） | 大（整个 1.4 MB dll） |
| 适合人群 | 已装过原版、想无感升级 | 未装原版、希望全自家可读 |

---

## [1.0.0] — 2026-05-21

首个正式版本。从原版「公文高手 WPS 插件单机版 v2.4.1」逆向工程产出。

### 新增

- 完整的 C# 源码工程 `src/Local_Wps_Vsto_v2/`（156 个 .cs 文件 + 4 个 .resx）
- 旧 msbuild 4.0 风格的 `Local_Wps_Vsto.csproj`，无需 Visual Studio / dotnet SDK，本机自带的 csc.exe (.NET Framework 4.x) 即可编译
- 零障碍安装脚本 `installer/install.ps1`（仅 HKCU 注册，无需管理员）
- 干净卸载脚本 `installer/uninstall.ps1`（自动清理 HKCU 与安装目录，可选保留用户数据 `-KeepUserData`）
- 15 项自检脚本 `installer/verify.ps1`（文件 / 注册表 / IL 字节 / 弱命名等多维度校验）
- 反编译驱动脚本 `tools/decompile.ps1`（用 ILSpy 7.2.1 ICSharpCode.Decompiler.dll Reflection API 一行调用，~ 1.8 秒完成）
- 完整文档：README、INSTALL、ARCHITECTURE、原理分析、工程复盘、CHANGELOG、LICENSE

### 变更

- 编译产物为**弱命名** dll（无 PublicKeyToken），不再依赖 GAC、不再依赖强名验证跳过
- 安装路径默认 `%LOCALAPPDATA%\GongwenAssistant`（原版 `C:\公文高手WPS插件单机版\`）
- AssemblyTitle / AssemblyProduct = `GongwenAssistant`（原版 `Wps_Vsto`）
- Ribbon 顶部选项卡名 = `公文助手 1.0.0`（原版 `公文高手单机版 2.4.1`）
- Ribbon 第一个分组 label = `公文助手`（原版 `尚未激活`，红字）
- 按钮 `btnUserMsg` 动态标签 = 始终 `公文助手`（原版「已激活」/「软件未激活」）
- `CommonConfig.strExeName` = `公文助手` / `EXE_WRITER_NAME` = `公文助手.exe` / `strBaseFolder` = `C:\公文助手\`
- 版本号 `strVersionCode` = `1.0.0`

### 移除 (功能上的)

- `UserUtil.IsVip()` 永远返回 `true`（原版从 SQL CE 读激活码做 SHA / DES 校验）
- `UserUtil.HasLogin()` 永远返回 `true`（原版从 SQL CE 读用户表）
- 由此自动屏蔽所有按钮事件中的「此功能登录后方可使用」/「此功能仅限 VIP 用户使用」拦截

### 保留 (出于代码可读性 / 工程审查)

- `MachineCode.cs` / `SecretUtil.cs` / `FingerPrint.cs` 原貌保留（无主动调用入口）
- `HttpUtil.cs` / `UpdateUtil.cs` / `WebMessageUtil.cs` 原貌保留（服务端已下线，调用全部 catch 静默失败）
- `readActiviedCode()` / `SaveActiviteCode()` 原貌保留（向 SQL CE 写值，但其结果 `boolVip` 不再被 IsVip 读取）

### 不动

- ProgId `Local_Wps_Vsto.MyAddin`（保持「替换」兼容性）
- CLSID `{2CD4E522-63D0-3A6B-9842-587DB1AB7FBF}`（由 namespace + class 名自动生成，namespace 不变 = CLSID 不变）
- TypeLib `Local_Wps_Vsto.tlb`（与原版一致）
- 所有 Word / Excel / Office / AddInDesignerObjects PIA 嵌入类型（仅 6 处反编译瑕疵被手工补 `[Guid]`）

---

## [Unreleased]

### 计划中

- 移除 Spire.Doc / Spire.Pdf 商业依赖，替换为开源 DocumentFormat.OpenXml + iTextSharp
- 清理已失效的网络模块代码（保留接口签名，移除实现）
- 添加 `installer/cleanup_legacy.ps1` —— 一键清理原版 v2.4.1 在 HKLM / GAC 中的残留（需管理员）
- 引入 GitHub Actions CI：每次 push 自动 msbuild + verify
- 单元测试覆盖 `WpsUtil` / `RtfUtil` / `CorrectUtil` 等核心模块
- Ribbon XML 从 `Resource1.resx` 中抽离为独立 `MyRibbon.xml` 文件
