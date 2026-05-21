# 变更记录 (CHANGELOG.md)

本项目遵循 [Semantic Versioning 2.0.0](https://semver.org/lang/zh-CN/) 与 [Keep a Changelog](https://keepachangelog.com/zh-CN/) 风格。

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
