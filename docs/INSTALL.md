# 安装手册 (INSTALL.md)

针对从来没接触过 PowerShell 或 COM Add-in 的用户。

---

## 一、系统要求

| 项目 | 最低要求 |
|------|----------|
| 操作系统 | Windows 10 / 11 (x64) |
| .NET Framework | 4.5 或更高 (Win10 自带 4.8, 通常无需手动安装) |
| WPS Office | 个人版 / 专业版 / 教育版均可, 需能正常打开 Word 文档 |
| 权限 | 普通用户即可, **无需管理员** |
| 可用磁盘 | ≥ 100 MB (主要是 Spire.Pdf.dll 13 MB) |

---

## 二、安装前检查（建议）

### 1. 卸载原版「公文高手」

如果您机器上之前装过原版 `公文高手Wps插件单机版2.4.1.exe`，请先用它自带的卸载程序卸载干净。如果不确定，可以打开 PowerShell 执行：

```powershell
# 检查 HKLM 里是否有原版残留
reg query 'HKLM\SOFTWARE\Wow6432Node\Classes\CLSID\{2CD4E522-63D0-3A6B-9842-587DB1AB7FBF}' 2>&1
# 检查 GAC 里是否有原版 dll
Test-Path 'C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Local_Wps_Vsto'
```

如果上面两条任一返回有内容 / `True`，建议先把原版卸载干净再继续。

### 2. 关闭 WPS

确保所有 WPS 文档已保存并关闭，避免文件锁定。安装脚本会自动尝试关闭 WPS，但事先手动关一次更稳。

---

## 三、安装步骤

### 1. 下载并解压

把 release zip（例如 `GongwenAssistant-1.0.0.zip`）解压到任意目录，**不要包含中文或空格更稳妥**，例如：

```
D:\GongwenAssistant\
├── runtime\
├── installer\
├── docs\
├── README.md
└── ...
```

### 2. 执行安装脚本

打开 PowerShell（开始菜单搜 `powershell` 即可，不需要「以管理员身份运行」），进入 installer 目录：

```powershell
cd D:\GongwenAssistant\installer
PowerShell -ExecutionPolicy Bypass -File .\install.ps1
```

如果一切顺利，最后一行应该是：

```
[ OK  ] 安装完成! 启动 WPS, 顶部菜单栏会出现「公文助手 1.0.0」选项卡
```

### 3. 自检

在同一个 PowerShell 窗口继续运行：

```powershell
PowerShell -ExecutionPolicy Bypass -File .\verify.ps1
```

应当看到 15 条 `[PASS]`，最终：

```
结果: PASS=15 FAIL=0
```

如果有任何 `[FAIL]`，请把完整输出截图反馈到 Issue。

### 4. 打开 WPS 验证 UI

启动 WPS Office 的「文字」（Word 文档）模式，新建或打开任意 docx 文件。Ribbon 顶部应该出现一个新选项卡：「**公文助手 1.0.0**」（位于「开始」标签左侧）。

切换到「公文助手 1.0.0」选项卡，您应该看到下列分组：

- 公文助手 (`[剩余 100 天]` 按钮)
- 素材/范文
- 排版
- 插入
- 校稿
- 其他

所有按钮均可点击，**不会再弹出「软件未激活 / 此功能仅限 VIP」提示**。

---

## 四、自定义安装目录

默认安装到 `%LOCALAPPDATA%\GongwenAssistant`，约 `C:\Users\<您的用户名>\AppData\Local\GongwenAssistant`。

如果想换位置（例如固定盘符）：

```powershell
PowerShell -ExecutionPolicy Bypass -File .\install.ps1 -InstallDir 'D:\Apps\GongwenAssistant'
```

注意：CodeBase 是写死到注册表里的路径，**安装后不要再随意挪动安装目录**，否则 WPS 会找不到 dll。

如果非要挪：先 `uninstall.ps1`，再用新路径重新 `install.ps1`。

---

## 五、常见问题

### Q1：安装脚本提示「禁止执行脚本」？

PowerShell 默认禁止运行外部脚本。请按手册中给的写法 `PowerShell -ExecutionPolicy Bypass -File ...`，这是一次性的、不会改全局策略。或者一次性放开：

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

### Q2：WPS 没有出现「公文助手」选项卡？

请按顺序排查：

1. 重启 WPS（关全部 wps.exe 进程再开）；
2. 跑 `verify.ps1` 看 15 项是否全 PASS；
3. 打开 WPS → 文件 → 选项 → 自定义功能区，看右侧勾选列表中是否有「公文助手」、是否被勾选；
4. 打开 WPS → 文件 → 选项 → 加载项，看「不活动的应用程序加载项」里是否有 `Local_Wps_Vsto.MyAddin` 被禁用，如有则在底部「管理 → COM 加载项」里重新启用。

### Q3：火绒 / 360 / Defender 报警？

本项目源码完全开放、可读、可重新构建。
若杀毒软件误报，请使用 `verify.ps1` 输出 + `src/` 源码自行确认安全性，然后将 `%LOCALAPPDATA%\GongwenAssistant\` 加入白名单。

更详细的「自证清白」说明见 `docs/原理分析.md`。

### Q4：安装失败提示路径名含「公文助手」字符乱码？

PowerShell 控制台默认 GBK 编码，中文路径在子进程下可能 lose-data。
本项目所有安装路径都是 ASCII 英文目录 (`GongwenAssistant`)，不会触发该问题。
如果您把 release zip 解压到了带中文名的目录，请用纯英文路径重新解压。

### Q5：和 Office 同时存在时会冲突吗？

不会。本项目只注册到 HKCU\Software\Microsoft\Office\Word\Addins\，Office Word 也会扫描这里。
但 WPS 和 Office Word 同时能加载，两边都会出现「公文助手」标签。
如果只想在 WPS 中用，可以在 Office Word 的「加载项」里手动禁用。

---

## 六、卸载

```powershell
PowerShell -ExecutionPolicy Bypass -File D:\GongwenAssistant\installer\uninstall.ps1
```

执行后：

- HKCU 下所有相关注册表项被清理；
- `%LOCALAPPDATA%\GongwenAssistant\` 整个删除。

如果想保留模板和用户配置数据，加 `-KeepUserData`：

```powershell
PowerShell -ExecutionPolicy Bypass -File .\uninstall.ps1 -KeepUserData
```

会保留 `conf\gwgs.sdf` 等用户数据文件，但仍然解除 COM 注册（即 WPS 下次启动不再加载插件）。

---

## 七、完全干净的状态

如果不放心残留，卸载后可手动复查：

```powershell
# 1. HKCU 应当为空
reg query 'HKCU\Software\Classes\CLSID\{2CD4E522-63D0-3A6B-9842-587DB1AB7FBF}' 2>&1
reg query 'HKCU\Software\Classes\Local_Wps_Vsto.MyAddin' 2>&1
reg query 'HKCU\Software\Microsoft\Office\Word\Addins\Local_Wps_Vsto.MyAddin' 2>&1
reg query 'HKCU\Software\Kingsoft\Office\WPS\AddinsWL' /v 'Local_Wps_Vsto.MyAddin' 2>&1

# 2. 安装目录应当被删
Test-Path "$env:LOCALAPPDATA\GongwenAssistant"
```

四条 `reg query` 应该全部返回「找不到」，`Test-Path` 应该返回 `False`。
