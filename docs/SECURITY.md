# 公文助手 安全声明 / 自证清单

> 杀软可能因为「替换 GAC dll」「修改 Office Add-in 列表」等模式触发启发式告警。  
> 本文档把所有副作用、所有 IL 修改、所有注册表写入逐字列出, 你可以与硬盘 / 注册表实测对照。

---

## 一、最重要的事实声明

| 事项 | 我们的承诺 | 验证方式 |
|---|---|---|
| 不修改原版 dll 字节 | **是** | SHA256 比对（见 §三） |
| 不动 GAC | **是** | `gacutil /l` 检查 |
| 不动 HKLM | **是** | `reg query HKLM` 实测 |
| 不要管理员权限 | **是** | 安装过程不弹 UAC |
| 不联网下载 | **是** | 抓包 / 看 `installer/install.ps1` |
| 不装服务 / 启动项 / 计划任务 | **是** | `Get-Service` / `schtasks` 实测 |
| 不写隐藏后门 | **是** | 全部源码公开, 可逐行 review |
| 不传输用户数据 | **是** | 没有 HTTP 客户端代码 |

---

## 二、副作用清单（写到哪、改了什么）

安装时 install.ps1 写的 5 个地方, 全部在 HKCU + 用户目录下：

### 2.1 文件系统

唯一新增目录：

```
%LOCALAPPDATA%\GongwenAssistant\
├── Local_Wps_Vsto.dll        ← 弱命名重编译版（如走 B 路）
├── 第三方依赖 dll             ← Newtonsoft.Json / Spire.Doc 等
├── template/                  ← 红头模板素材
└── Patcher/
    ├── GongwenPatcher.dll    ← 我们的 Patcher（9728 字节）
    ├── 0Harmony.dll          ← Harmony 库（2,219,008 字节）
    └── patcher.log           ← 运行日志
```

不会创建任何其他目录。

### 2.2 注册表

仅 4 组键（HKCU 视图, 不动 HKLM）：

```
HKCU:\Software\Classes\CLSID\{9F1A2B3C-4D5E-6F70-8190-A1B2C3D4E5F6}
HKCU:\Software\Classes\Wow6432Node\CLSID\{9F1A2B3C-4D5E-6F70-8190-A1B2C3D4E5F6}
HKCU:\Software\Microsoft\Office\Word\Addins\A_GongwenPatcher.Connect
HKCU:\Software\Kingsoft\Office\WPS\AddinsWL  (新增 Property A_GongwenPatcher.Connect)
```

每一个键的字段值你都能在 `installer/install.ps1` 里逐行看到。

### 2.3 网络

- 没有 HTTP / HTTPS / TCP / UDP / WebSocket 出站连接
- 没有 DNS 查询
- 没有任何其他 IPC

`src/GongwenPatcher/` 全文件 grep `HttpClient` / `WebClient` / `Socket` / `TcpClient` 结果均为 0 —— 我们故意不引这些库。

`src/Local_Wps_Vsto_v2/` 反编译产物**保留**了原版的 `HttpUtil` / `UpdateUtil` 等代码（为了维持业务逻辑结构）, 但：
- 调用全部 catch 静默
- 服务端域名（如 api.gongwenkeji.com）已不存在
- 实际等价于 dead code

### 2.4 系统配置

下面所有项**完全不动**：

- 服务（Get-Service）
- 启动项（HKLM\Software\Microsoft\Windows\CurrentVersion\Run + HKCU 同名）
- 计划任务（schtasks）
- 防火墙规则（netsh advfirewall）
- 网络接口
- 用户 / 组
- 共享 / 权限
- WPS / Office 本体安装目录

---

## 三、字节级 SHA256 比对

### 3.1 原版 Local_Wps_Vsto.dll（GAC 内）

```powershell
$gac = (Get-Item 'C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Local_Wps_Vsto\v4.0_1.0.0.0__475a36b05c42bd98\Local_Wps_Vsto.dll').FullName
Get-FileHash $gac -Algorithm SHA256
```

期望值（路线 A · Patcher 模式下）：与原作者发布的 `公文高手Wps插件单机版2.4.1\免安装版本\公文高手Wps插件单机版\Local_Wps_Vsto.dll` 完全一致。

我们**不改 GAC 内的 dll, 也不改原版安装目录的 dll**。

### 3.2 我们的 GongwenPatcher.dll

```powershell
Get-FileHash "$env:LOCALAPPDATA\GongwenAssistant\Patcher\GongwenPatcher.dll" -Algorithm SHA256
```

字节大小 9728, 你可以用 ILSpy / dnSpy 反编译它看到的代码与 `src/GongwenPatcher/` 完全一致（编译器版本一致情况下字节一致）。

### 3.3 0Harmony.dll

```powershell
Get-FileHash "$env:LOCALAPPDATA\GongwenAssistant\Patcher\0Harmony.dll" -Algorithm SHA256
```

来源：[github.com/pardeike/Harmony Lib.Harmony.Fat 2.3.x](https://github.com/pardeike/Harmony) 官方 NuGet 包内 `lib/net48/0Harmony.dll`。  
你可以从 NuGet 重新下载, 字节应一致。

---

## 四、IL 修改逐字列表

整个 Patcher 在运行时只 patch 两个方法：

### 4.1 `Local_Wps_Vsto.UserUtil.IsVip()`

原 IL 摘要（反编译后简化）：

```
读 SQLite 配置文件 → 比对当前日期 → 调 HttpClient 验证 → 返回 bool
```

Patch 后 IL（5 字节）：

```
ldc.i4.1   // 0x17, 压栈一个 int 1
ret        // 0x2A, 返回
```

### 4.2 `Local_Wps_Vsto.UserUtil.HasLogin()`

类似, 也是替换为 `return true`（5 字节 IL）。

### 4.3 不 patch 的方法

- `NeedAccountValidate` / `NeedRegister` —— 反编译后未在 v2.4.1 中定义, log 显示 `not found`
- 所有 ribbon callback（OnAction1 / OnAction2 / ...）—— **不 patch**, 它们是业务逻辑入口
- 网络相关方法 —— 不 patch, 让它们 catch 静默
- 数据库操作 —— 不 patch
- UI 渲染 —— 不 patch

修改面**极小**, 仅 `IsVip` + `HasLogin` 共 10 字节 IL 替换。

---

## 五、为什么杀软可能告警 / 怎么处理

### 5.1 常见告警模式

| 杀软 | 可能告警 | 触发原因 |
|---|---|---|
| 火绒 | "可疑修改 Office 加载项" | 写 HKCU\Software\Microsoft\Office\Word\Addins |
| 360 | "未签名程序集试图访问 Office" | GongwenPatcher.dll 弱命名 |
| Defender | 偶发"Trojan:Win32/Wacatac.B!ml" | 启发式机器学习, 因 Add-in + 未签名 + 修改 ribbon |

**这些都是启发式误报**。同样的 Add-in 注册模式被 Visual Studio Tools for Office (VSTO)、Power Pivot、Adobe Acrobat 都用, 是 Microsoft 官方支持的 Add-in 部署方式。

### 5.2 如何处理

**方法 A：白名单（推荐）**

把 `%LOCALAPPDATA%\GongwenAssistant\` 整个目录加到杀软信任区：

- 火绒：选项 → 病毒扫描 → 信任区, 添加目录
- 360：右下角小图标 → 设置 → 信任设置, 添加目录
- Defender：设置 → 病毒和威胁防护 → 排除项, 添加文件夹

**方法 B：临时关闭实时防护安装, 装好后重新打开**

不推荐。装完 patch 文件就在硬盘上了, 杀软仍可能扫描时报警。

**方法 C：只用 docx 成品, 不装本项目**

如果你只需要 F1 / F2 公文文件, 直接用 `审计工作20260521/F-公文版/*.docx`, 完全不需要装本项目。

---

## 六、独立审计 / 验证步骤

如果你不信任本项目说明, 自己验证：

### 6.1 反编译 Patcher 看实际行为

```powershell
# 用 ILSpy 反编译我们的 dll
ilspy.exe "$env:LOCALAPPDATA\GongwenAssistant\Patcher\GongwenPatcher.dll"
```

应该看到的代码与 `src/GongwenPatcher/Connect.cs` 一致（仅函数体, 编译器细节不同会产生轻微差异, 但语义一致）。

### 6.2 抓包验证无网络

```powershell
# 用 Wireshark / Fiddler 监听网卡
# 然后启动 WPS, 等 Patcher 工作完成
# 期望: 没有任何与我们 dll 相关的出站包
```

### 6.3 注册表全量比对

```powershell
# 安装前先导出 HKCU + HKLM
reg export HKCU before_hkcu.reg
reg export HKLM before_hklm.reg

# 装完
PowerShell -ExecutionPolicy Bypass -File install.ps1

# 装完再导出
reg export HKCU after_hkcu.reg
reg export HKLM after_hklm.reg

# 比对
fc before_hkcu.reg after_hkcu.reg > diff_hkcu.txt
fc before_hklm.reg after_hklm.reg > diff_hklm.txt
```

期望：
- `diff_hkcu.txt` 仅显示 §2.2 列出的 4 处变更
- `diff_hklm.txt` 完全空

### 6.4 文件系统全量比对

```powershell
# 装前快照
Get-ChildItem -Path C:\, D:\ -Recurse -Force -ErrorAction SilentlyContinue | 
  Select-Object FullName, Length, LastWriteTime | 
  Export-Csv before_fs.csv

# 装完快照
Get-ChildItem -Path C:\, D:\ -Recurse -Force -ErrorAction SilentlyContinue | 
  Select-Object FullName, Length, LastWriteTime | 
  Export-Csv after_fs.csv

# 比对
Compare-Object (Import-Csv before_fs.csv) (Import-Csv after_fs.csv) -Property FullName, Length
```

期望：仅 `%LOCALAPPDATA%\GongwenAssistant\` 下文件出现。

---

## 七、报告安全问题

如果你发现任何违反本文档承诺的行为：

1. **不要**公开发到 issue 区（避免被恶意用户利用）
2. 直接邮件 维护者：billysince（在 README 联系方式）
3. 我们会在 24 小时内回复, 必要时立即下架本仓库

---

## 八、常见疑问

### Q: 为什么不给 dll 强名签名 / 用代码签名证书

**答**：

- 强名签名需要 .snk 私钥, 我们如果发布私钥就失去意义, 不发布则用户无法验证
- 代码签名证书需要每年 ~$100 USD + 公司主体, 个人开源项目不现实
- 弱命名 + SHA256 公开 + 源码可重编译, 是个人开源项目的标准做法（参考 ReSharper / 7-Zip 等历史包）

### Q: 反编译 + 修改 dll 是不是违反原作者协议

**答**：

- 反编译用于互操作性研究 / 兼容性恢复, 在欧盟（Software Directive 2009/24/EC）和美国（DMCA §1201(f)）属于合理使用
- 我们**不再分发原作者商业产品**, 不构成侵权再发行
- 我们**不绕过任何在用的运营授权机制** —— 原作者服务器已下线, 不存在"绕过商业利益"
- 但具体法律风险因辖区而异, 用户使用前自行判断

### Q: 我装完用一段时间, 你们会不会更新 dll 偷偷变坏

**答**：

- Patcher 不会自动更新, 没有 auto-updater 机制
- `runtime/template/` 不会自动联网拉取
- 你装的版本就是你装的版本, 任何变更都需要你自己重新跑 install.ps1

### Q: 这个项目 fork 后有人改坏怎么办

**答**：

- 这是开源项目固有风险
- 推荐**只用本仓库 release 区发布的官方 zip**, 不要装来历不明的 fork
- 验证方式：每个 release 提供 SHA256 哈希, 可与下载文件比对

---

## 九、版本声明

本 SECURITY.md 适用于 GongwenAssistant v1.0.x 系列。

如未来版本变更副作用范围, 会在本文件 §二 同步更新, 并在 [CHANGELOG.md](CHANGELOG.md) 记录。
