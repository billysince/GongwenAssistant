# 公文助手 (GongwenAssistant)

> 面向中文公文写作场景的 WPS Office Add-in。基于「公文高手 WPS 插件单机版 v2.4.1」逆向工程产出的可读源码重新编译，去除原版 VIP / 激活限制，按透明开源精神开放给同样有公文写作需求的工作人员。

---

## 这是什么

一个挂在 WPS 文字处理顶部的 Ribbon 选项卡，提供：

- 公文标准排版（设为 A4、一键排版、各级标题、公文正文、文末日期）
- 公文常用元素插入（一二三级标题、符号、页码、横页、日期）
- 红头模板 / 素材搜索 / 范文搜索 / 写作提词器
- 朗读校稿、提纲检查、保存与导出 PDF
- 模板备份、恢复、外部资源导入

UI 来自原版插件，业务功能保持 100% 兼容；底层 dll 经反编译重建，去除了所有「软件未激活 / 此功能仅限 VIP 用户」类强制提示与拦截。

---

## 为什么会有这个项目

原版「公文高手 WPS 插件单机版 v2.4.1」是一款实用的本地公文助手，但运营方已停止线上服务，导致部分功能在新机器上无法激活 / 登录，正常使用者付费后也无法保证后续可用。本项目站在工程角度做了三件事：

1. **逆向工程** —— 把原 DLL 完整反编译为 C# 源码，使任何工程师都能 review 业务逻辑，去除「黑盒不可信任」属性；
2. **本地化加载路径** —— 改用 HKCU + CodeBase 加载，**不动 HKLM、不动 GAC、不需要管理员权限**，把对系统的副作用降到最低；
3. **可重复构建** —— 提供完整 csproj + 编译脚本，任何人都可以在自己机器上 `msbuild` 出一份与发布版字节一致的 dll，确认没有夹带任何后门。

---

## 与原版的关系

- 本项目**不打包**原版安装程序 `公文高手Wps插件单机版2.4.1.exe`，也**不再分发**原作者的服务器接口。
- 仅保留：插件 DLL 的源码层重构、模板资源 (`template/*`)、必需的第三方依赖二进制（`Newtonsoft.Json` / `Spire.Doc` 等）。
- 第三方商业组件（如 `Spire.Doc.dll`、`System.Data.SqlServerCe.dll`）的版权归各自原厂，本项目以二进制依赖形式引入，仅供本项目内部 COM 加载用。
- 如果您仍在使用原版，强烈建议先卸载原版再安装本项目，避免 ProgId / CLSID 冲突。

---

## 两条部署路线

本仓库同时维护两条经过完整验证的部署路线，你可以按自己的偏好二选一：

### 路线 A · Patcher（推荐）

保留原版 `Local_Wps_Vsto.dll` 字节零修改、PKT 不变、GAC 不变；额外装一个 `A_GongwenPatcher.Connect` COM Add-in，它在 WPS 进程内用 [Harmony](https://github.com/pardeike/Harmony) 把 `UserUtil.IsVip` / `HasLogin` 的 IL 替换为 `return true`。

适合的场景：
- 你已经装过原版 `公文高手Wps插件单机版2.4.1.exe`，想做无感升级绕掉 VIP 拦截；
- 不希望维护一份自己重编译的二进制；
- 希望原版有任何升级时 Patcher 还能继续兼容。

完整原理与脚本：[`docs/Patcher方案.md`](docs/Patcher方案.md)。

### 路线 B · 自家重编译（详见下面"快速开始"）

ILSpy 反编译 → 在 `src/Local_Wps_Vsto_v2/` 维护可读 C# 源码 → msbuild 出弱命名 dll → 用 HKCU + CodeBase 加载，整体不依赖 GAC、不需要管理员。

---

## 快速开始

> 适用环境：Windows 10 / 11 + WPS Office 个人版或专业版（已安装且能正常打开 Word 文档）。

### 安装

```powershell
# 1. 解压 release zip 到任意目录, 例如 D:\GongwenAssistant
# 2. 进入 installer 子目录, 右键 install.ps1 → 用 PowerShell 运行
#    或者命令行:
PowerShell -ExecutionPolicy Bypass -File D:\GongwenAssistant\installer\install.ps1
```

安装脚本会：

1. 将 `runtime\` 下所有文件复制到 `%LOCALAPPDATA%\GongwenAssistant\`（默认）；
2. 写入 4 组 HKCU 注册表项（CLSID、ProgId、Office Word Addins、WPS AddinsWL）；
3. 自动检测并关闭可能正在运行的 WPS。

安装完成后启动 WPS，顶部 Ribbon 会出现「公文助手 1.0.0」选项卡。

### 自检

```powershell
PowerShell -ExecutionPolicy Bypass -File D:\GongwenAssistant\installer\verify.ps1
```

预期输出末尾应为 `结果: PASS=15 FAIL=0`。

### 卸载

```powershell
PowerShell -ExecutionPolicy Bypass -File D:\GongwenAssistant\installer\uninstall.ps1
```

会清理 HKCU 下全部安装相关项 + 删除安装目录。如需保留模板与配置数据，可加 `-KeepUserData`。

---

## 从源码自行编译

```powershell
# 需要本机已安装 .NET Framework 4.x (Windows 10+ 默认满足)
$proj = 'src\Local_Wps_Vsto_v2\Local_Wps_Vsto.csproj'
$msb  = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe'
& $msb $proj /t:Build /p:Configuration=Release
# 产物位于 src\Local_Wps_Vsto_v2\bin\Release\Local_Wps_Vsto.dll
```

无需 Visual Studio、无需 dotnet SDK、无需 NuGet 还原。所有依赖通过 `HintPath` 直接指向 `公文高手Wps插件单机版2.4.1\免安装版本\公文高手Wps插件单机版\*.dll`（如果不需要原版目录，可手动改 `HintPath` 指向本项目 `runtime\` 目录）。

---

## 目录结构

```
GongwenAssistant/
├── src/
│   └── Local_Wps_Vsto_v2/         # ILSpy 7.2.1 反编译 + 手工修复后的 C# 源码
│       ├── Local_Wps_Vsto.csproj  # 旧 msbuild 4.x 风格, 不依赖 dotnet SDK
│       ├── Local_Wps_Vsto/        # 业务代码 (60+ 文件)
│       ├── Word/  Excel/  Office/ AddInDesignerObjects/   # PIA 嵌入类型
│       ├── Properties/AssemblyInfo.cs
│       └── *.resx                 # 资源 (ribbon XML / RTF / 图片)
├── runtime/                       # 发布运行时, 安装到 %LOCALAPPDATA%
│   ├── Local_Wps_Vsto.dll         # 新编译的弱命名 patched dll (1,408,512 bytes)
│   ├── Newtonsoft.Json.dll
│   ├── Spire.Doc.dll / Spire.Pdf.dll / Spire.License.dll
│   ├── System.Data.SqlServerCe.dll
│   ├── ICSharpCode.SharpZipLib.dll
│   ├── Microsoft.mshtml.dll
│   ├── RibbonControlsLibrary.dll
│   ├── EntityFramework.dll
│   ├── amd64/  x86/               # SQL CE native runtime
│   ├── conf/                      # 默认配置 (sdf)
│   └── template/                  # 红头模板 / 范文 / 排版样式
├── installer/
│   ├── install.ps1                # 零障碍安装 (HKCU only)
│   ├── uninstall.ps1              # 卸载
│   └── verify.ps1                 # 15 项自检
├── tools/
│   ├── ilspy72/                   # 反编译器 (ILSpy 7.2.1) - 仅开发时使用
│   ├── dnspy/                     # dnSpyEx 6.5.1 - 备用反编译器
│   └── decompile.ps1              # 反编译驱动脚本 (PowerShell + Reflection)
├── docs/
│   ├── ARCHITECTURE.md            # 系统设计
│   ├── CHANGELOG.md               # 变更记录
│   ├── INSTALL.md                 # 详细安装手册
│   ├── 原理分析.md                # WPS COM Add-in 加载链路解构
│   └── 工程复盘.md                # 决策点 + try / error 记录
├── bin/                           # msbuild 构建日志归档
├── LICENSE
└── README.md                      # 本文件
```

---

## 透明开源声明

- 本项目的所有源码（`src/Local_Wps_Vsto_v2/`）由 ILSpy 7.2.1 自原版 dll 反编译产出后人工修复，未经混淆；
- 所有注册表 / 文件 / 网络副作用都在 `installer/*.ps1` 中明文写出，可逐行 review；
- 没有任何「外部 C&C」 / 隐藏后门 / 数据回传逻辑（原版 dll 中的 `HttpUtil` / `UpdateUtil` 等代码保留，但服务端已下线，调用全部 catch 静默失败，可在源码中确认）；
- 自检脚本 `installer/verify.ps1` 用 `Assembly.ReflectionOnlyLoadFrom` 直接读取 PE 文件 IL 字节，对 `UserUtil.IsVip()` / `UserUtil.HasLogin()` 做最终校验。

如果您发现任何与上述声明不符的地方，请在 Issue 区直接指出。

---

## 致谢

- 原作者「公文高手」团队 —— 完整的功能设计、UI 与模板素材
- ILSpy / dnSpy 团队 —— 高质量的 .NET 反编译器
- 本项目维护者：billysince
