# 公文助手 FAQ

> 高频问题 + 直接答案。每个问题都给到「短答 → 长答 → 引用源 / 验证步骤」。  
> 如果你的问题不在这里, 请提 Issue, 我们会补上。

---

## 一、装机相关

### Q1：装完启动 WPS 看不到「公文助手单机版2.4.1」tab

**短答**：99% 是 WPS 同时跑了多个进程, 或上一次 install 失败有残留。

**长答**：

按以下顺序排查：

1. 先确认 patcher.log 在写：
   ```powershell
   Get-Content "$env:LOCALAPPDATA\GongwenAssistant\patcher.log" -Tail 5
   ```
   - 有 `Patched IsVip` → patcher 在 work, 跳到 step 3
   - 没有 / 文件不存在 → step 2
   
2. 跑 verify.ps1, 看哪些 PASS / FAIL：
   ```powershell
   PowerShell -ExecutionPolicy Bypass -File D:\GongwenAssistant\installer\verify.ps1
   ```
   - 期望 `PASS=15 FAIL=0`
   - 任何 FAIL → 卸载重装：先 `uninstall.ps1`, 再 `install.ps1`

3. patcher 在 work 但 tab 没显示, 通常是 WPS 多进程问题：
   ```powershell
   Get-Process wps | Select-Object Id, MainWindowTitle, StartTime
   ```
   - 如果有多个 wps 进程, **手动**用窗口右上角 × 关掉所有 WPS（不要用 Stop-Process）
   - 等 30 秒让所有进程退出
   - 重启 WPS

4. 仍然不行 → 看 [踩坑全集.md #19 #22](踩坑全集.md)

**验证源**：[QUICKSTART.md §4](QUICKSTART.md), [INSTALL.md §故障排查](INSTALL.md)

---

### Q2：tab 显示, 但点按钮（红头模板 / 范文搜索 / 设为A4）没反应

**短答**：WPS 12.1+ 内核屏蔽了 COM Add-in 的 onAction 回调, 这是 WPS 的安全策略, **我们绕不过**, 不是本项目的 bug。

**长答**：

WPS 12.0.1.17xx 版本起, 出于"防恶意 Add-in"考虑, 把 COM Add-in 的运行时 ribbon callback 链路（`onAction` / `getEnabled` / `getLabel` / `getVisible` / `getImage`）屏蔽了。表现是：

- ribbon tab 仍能显示（因为 GetCustomUI 是初始化期一次性调用, 没被砍）
- VIP 字样仍能显示（IsVip 是业务代码内部调用, 没被砍）
- 点按钮**完全没反应** —— 因为 onAction 回调链路被 WPS 内核拦截

**怎么验证**：

1. 加调试日志到 ribbon callback 函数, 看是否被调用 → 不被调用
2. 用 WPS 旧版本（11.x / 12.0 早期）能调用
3. 同一份 dll 装到 MS Office Word 上 onAction 仍工作（说明不是 dll 问题）

**有没有解决办法**：

| 方案 | 状态 |
|---|---|
| 走 JS Add-in 路线（B 路） | 撞墙, 见 [踩坑全集.md #26](踩坑全集.md) |
| 走 wpsoa 官方应用商店发布 | 需要软件著作权 + 审核, 个人无法走通 |
| 降级 WPS 到 12.0 之前 | 用户体验损失大, 不推荐 |
| 迁到 MS Office Word | 需要购买 Office, 部分用户不可行 |
| 等 WPS 内核未来开放 | 没人能担保时间 |

**当前推荐**：

- 如果你只是想要具体 docx → 用 `审计工作20260521/F-公文版/` 已交付成品
- 如果你想自己写公文 → 用 markdown + Pandoc + python-docx 链路（见 `审计工作20260521/script/`）
- 如果你必须用 ribbon 按钮 → 切换到 MS Office Word

**验证源**：[踩坑全集.md #25](踩坑全集.md), 完整定位过程见 [工程复盘.md §第 X 段](工程复盘.md)

---

### Q3：WPS 启动比以前慢了

**短答**：本项目正常情况下增加约 200～400 ms 启动时间（patch 两个方法的 IL）, 不应该让人体感"明显变慢"。如果明显变慢, 通常是装了多次重复 Add-in。

**长答**：

按以下排查：

1. 看 patcher.log 的时间戳：
   ```powershell
   Get-Content "$env:LOCALAPPDATA\GongwenAssistant\patcher.log" -Tail 20
   ```
   - 从 `OnConnection start` 到 `Patched IsVip` 通常 < 500ms
   - > 1000ms → step 2

2. 检查是否有重复 Add-in 注册：
   ```powershell
   Get-ChildItem 'HKCU:\Software\Microsoft\Office\Word\Addins' | Select-Object PSChildName
   ```
   - 期望只看到 `Local_Wps_Vsto.MyAddin` + `A_GongwenPatcher.Connect`
   - 如有其他 Add-in（特别是同名带数字后缀）→ 历史残留, 用 uninstall.ps1 清

3. 检查 AddinsCL：
   ```powershell
   Get-ItemProperty 'HKCU:\Software\Kingsoft\Office\WPS\AddinsCL' -ErrorAction SilentlyContinue
   ```
   - 如果有累积条目, WPS 会反复尝试加载失败的 Add-in, 导致变慢
   - 用 uninstall.ps1 + install.ps1 清

---

### Q4：装完之后火绒 / 360 / Defender 报警

**短答**：误报, 把 `%LOCALAPPDATA%\GongwenAssistant\` 整个目录加白名单。

**长答**：见 [SECURITY.md](SECURITY.md) 完整自证 + 字节级 SHA256 比对。

简短版：

- 我们的 dll **不修改原版强名 dll** —— 原版 SHA256 不变
- 我们 patch 只在内存里, 硬盘字节零修改
- 我们写的注册表项明文在 `installer/install.ps1`, 任何人可逐行 review
- 没有联网代码 / 隐藏服务 / C&C / 后门

如果你的杀软仍报警, 选其一：
- 把 `%LOCALAPPDATA%\GongwenAssistant\` 加白名单（推荐）
- 不要用本项目, 直接用 `审计工作20260521/F-公文版/` 已交付成品

---

### Q5：能不能装到 MS Office Word（不是 WPS）

**短答**：理论上能, 因为 COM Add-in 链路在 Word 上完整, 但**未测试**。

**长答**：

- 原版 `Local_Wps_Vsto.dll` 设计目标就是 Word + WPS 双栈, 加载时先尝试 Word 再尝试 WPS
- 我们的 Patcher 也注册到 `HKCU\Software\Microsoft\Office\Word\Addins`, 路径相同
- onAction 在 MS Office Word 上**应该**能工作（Word 没砍 ribbon callback）
- 但我们没有 MS Office 测试环境, 不敢保证, 用了请反馈

**愿意测试的人欢迎在 Issue 区告知, 我们补充测试报告。**

---

## 二、原理 / 安全相关

### Q6：你们到底改了我电脑哪些东西

**短答**：仅 HKCU 注册表 4 组键 + `%LOCALAPPDATA%\GongwenAssistant\` 一个文件夹。逐字明细见 [INSTALL.md](INSTALL.md)。

**长答**：

| 类别 | 路径 | 用途 |
|---|---|---|
| 文件 | `%LOCALAPPDATA%\GongwenAssistant\` | dll、模板、依赖 |
| 注册表 | `HKCU\Software\Classes\CLSID\{9F1A...}` | Patcher 的 COM 类注册（64 位视图） |
| 注册表 | `HKCU\Software\Classes\Wow6432Node\CLSID\{9F1A...}` | Patcher 的 COM 类注册（32 位视图, 给 WPS 用） |
| 注册表 | `HKCU\Software\Microsoft\Office\Word\Addins\A_GongwenPatcher.Connect` | Word 端 Add-in 列表 |
| 注册表 | `HKCU\Software\Kingsoft\Office\WPS\AddinsWL` | WPS 白名单 |

**不动的地方**：

- HKLM 整个分支
- GAC（全局 .NET 缓存）
- 任何系统目录
- 任何启动项 / 服务 / 计划任务
- 网络配置
- WPS 本体安装目录

卸载时上面所有改动 100% 撤回。

---

### Q7：Harmony 是什么？为什么不直接改原 dll

**短答**：Harmony 在内存里临时改方法实现, 不动硬盘文件。直接改 dll 字节会让强名失效 + 触发杀软。

**长答**：见 [术语表.md - Harmony](术语表.md) + [Patcher方案.md](Patcher方案.md)

简短版对比：

| 方案 | dll 字节 | SHA256 | PKT | 杀软 | 卸载 |
|---|---|---|---|---|---|
| 直接改 dll | 改 | 变 | 失效 | 高几率告警 | 难 |
| Harmony 内存 hook | **不变** | **不变** | **不变** | **极低** | WPS 关掉就回原样 |

我们选 Harmony 路线就是为了这五个属性。

---

### Q8：服务器下线了, 范文 / 红头模板还能用吗

**短答**：本地模板能用（`runtime/template/`）, 联网功能（在线检索范文 / 用户登录 / 模板更新）全部下线 → catch 静默失败。

**长答**：

原作者「公文助手」团队的服务器（如 `api.gongwenkeji.com`）已不再响应。受影响的功能：

- 在线范文搜索 → 接口 404, 返回空列表
- 用户登录 / 注册 → 接口 404, 客户端 catch 静默
- 在线红头模板更新 → 接口 404, 不影响本地模板使用
- 升级检查 → 接口 404, 不影响本地使用

**未受影响的功能**（即使在能点 ribbon 按钮的环境下）：

- 本地红头模板（在 `runtime/template/`）
- 本地范文（在 `runtime/template/范文/`）
- 本地素材库（在 `runtime/template/素材/`）
- 本地排版功能（设为 A4 / 各级标题 / 文末日期）
- 本地写作提词器（基于本地词库）

---

## 三、协议 / 法律相关

### Q9：这个项目合法吗？我用了会不会被告

**短答**：本项目仅用于个人学习与已购买用户的自服务用途, 不分发原作者收费内容, 不再联网调原作者服务器。

**长答**：

我们做的事：
- 反编译原 dll, 给已购买用户提供「服务器下线后仍可用」的本地使用路径
- 不再分发原版安装程序 `公文助手Wps插件单机版2.4.1.exe`
- 不再使用原作者商业域名 / 接口

不要做的事（用户责任）：
- 不要把本项目商业化再分发
- 不要把本项目用于未购买原版的场景下假冒原版商业销售
- 如果原作者团队恢复运营要求下架, 本仓库会立即下架

我们的法律处境：
- 反编译用于互操作性研究 / 已购买用户兼容性恢复, 在多数司法辖区属于合理使用
- 但具体法律风险因辖区而异, 用户自行判断

---

### Q10：能给我商业授权吗

**短答**：不能。本项目源码 MIT, 但反编译产物受原作者著作权保护, 我们没有原作者授权, 也无权再授权他人。

**长答**：

- 本仓库自有的脚本 / 文档（installer/, docs/, README.md）以 MIT 协议释放
- `src/Local_Wps_Vsto_v2/` 内反编译产物**不**以 MIT 释放, 仅作个人学习与已购买用户兼容性恢复用途
- 任何商业使用请直接联系原作者团队（如能联系上）

---

## 四、开发 / 贡献相关

### Q11：怎么自己重编译 dll

**短答**：见 [README.md §六](../README.md), 一行 msbuild 命令。

**长答**：

```powershell
# 不需要装 Visual Studio
$proj = 'D:\GongwenAssistant\src\Local_Wps_Vsto_v2\Local_Wps_Vsto.csproj'
$msb  = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe'
& $msb $proj /t:Build /p:Configuration=Release
# 产物在 src\Local_Wps_Vsto_v2\bin\Release\Local_Wps_Vsto.dll
# 字节应与发布版完全一致（除非你改了源码）
```

字节比对：

```powershell
$expect = (Get-FileHash 'D:\GongwenAssistant\runtime\Local_Wps_Vsto.dll').Hash
$actual = (Get-FileHash 'D:\GongwenAssistant\src\Local_Wps_Vsto_v2\bin\Release\Local_Wps_Vsto.dll').Hash
if ($expect -eq $actual) { "字节一致" } else { "字节不一致 expect=$expect actual=$actual" }
```

---

### Q12：怎么提 Issue / PR

见 [CONTRIBUTING.md](CONTRIBUTING.md)。

---

### Q13：能不能加个 onAction 不响应的临时绕过

**短答**：我们试过 4 种, 全部撞墙。详见 [踩坑全集.md #25](踩坑全集.md)。

**长答**（4 种已尝试方案）：

1. **WPS Word 自动化兼容层** —— WPS 没暴露 onAction 反射调用接口
2. **直接 PostMessage 模拟点击** —— WPS ribbon 是 Qt 自绘, 不响应标准 Windows 消息
3. **JS Add-in 路线** —— 个人版 dev 模式被砍, [踩坑全集.md #26](踩坑全集.md)
4. **官方应用商店发布** —— 需要软著审核 + 公司主体, 个人不可行

如果你有第 5 种思路, 强烈欢迎在 Issue 区分享。

---

## 五、不在 FAQ 里的问题

直接提 Issue, 我们会补到这里。
