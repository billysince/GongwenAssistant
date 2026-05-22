# 公文助手 5 分钟极速上手

> 这份文档假设你**完全不懂代码**, 只会"双击 / 右键 / 复制粘贴"。  
> 全程 5～10 分钟, 不需要安装任何编译器、不需要 GitHub 账号、不需要管理员权限。

---

## 你需要准备

1. **Windows 10 或 11** —— 在「设置 > 系统 > 关于」可以看到版本号
2. **WPS Office** —— 任何版本, 桌面上能找到 WPS 图标即可
3. **本仓库 release zip** —— 文件名形如 `GongwenAssistant-1.0.x-win.zip`
4. **5 分钟空闲时间**

---

## 第 1 步：解压

把 zip 解压到任意你看得到的目录, 推荐 `D:\GongwenAssistant\`。

解压后里面应该长这样：

```
D:\GongwenAssistant\
├── README.md          ← 项目说明
├── runtime\           ← 运行时文件（dll、模板、依赖）
├── installer\         ← 安装脚本
│   ├── install.ps1    ← 用这个
│   ├── uninstall.ps1  ← 卸载用这个
│   └── verify.ps1     ← 验证用这个
├── docs\              ← 文档（包括本文件）
└── LICENSE
```

---

## 第 2 步：跑安装脚本

### 方法 A：右键运行（最简单）

1. 在 `installer\` 目录里找到 `install.ps1`
2. 在文件上**右键**, 选「使用 PowerShell 运行」

如果 Windows 弹「策略禁止运行脚本」, 用方法 B。

### 方法 B：PowerShell 命令行

1. 按 **Win 键**, 输入 **PowerShell**, 打开
2. 复制下面这行, **路径换成你自己的解压路径**, 然后粘贴回车：

```powershell
PowerShell -ExecutionPolicy Bypass -File D:\GongwenAssistant\installer\install.ps1
```

### 脚本会做什么（透明声明）

每一步都有屏幕输出, 你能看到它在做什么：

```
[INFO] 检查 PowerShell 版本... OK
[INFO] 检查 .NET Framework 4.x... OK
[INFO] 检查 WPS 是否在跑... 没在跑
[STEP] 复制 runtime\ 到 %LOCALAPPDATA%\GongwenAssistant\
[STEP] 写注册表 HKCU\Software\Classes\CLSID\{...}
[STEP] 写注册表 HKCU\Software\Microsoft\Office\Word\Addins\Local_Wps_Vsto.MyAddin
[STEP] 写注册表 HKCU\Software\Kingsoft\Office\WPS\AddinsWL
[OK]   完成
```

**不会做的事**（你应该警惕的事我们不做）：

- 不会改 `HKLM`（整机级注册表）
- 不会改 GAC（系统全局 .NET 缓存）
- 不会装服务 / 启动项 / 计划任务
- 不会联网下载任何东西
- 不会要求管理员密码（即 UAC 弹窗不应该出现）

如果中途弹 UAC 要管理员权限, **立即取消, 来 Issue 区报告**。

---

## 第 3 步：启动 WPS, 看到 tab

1. 打开 WPS Office, 新建或打开任何 docx 文档
2. 看顶部菜单栏, 应该看到「**公文助手**」
3. 点进去, 看到「终身VIP - **已激活**」字样

期望看到的样子：

![expected_state](../dist/wps_now.png)

如果看到了 → **安装成功**, 跳到第 5 步。  
如果没看到 → 看第 4 步排查。

---

## 第 4 步：没看到 tab？排查表

### 4.1 先看 patcher.log 是否在写

在 PowerShell 里执行：

```powershell
Get-Content "$env:LOCALAPPDATA\GongwenAssistant\patcher.log" -Tail 10
```

| 你看到 | 说明 |
|---|---|
| `Patched IsVip` `Patched HasLogin` | patcher 在 work, 但 ribbon 渲染有问题 → 4.2 |
| `Local_Wps_Vsto detected: ... PublicKeyToken=null` | dll 弱命名版被加载了, 但 IsVip patch 失败 → 4.3 |
| 文件不存在 | patcher 没被加载 → 4.4 |

### 4.2 patcher 在 work 但 tab 没显示

1. 检查 WPS 是不是同时跑了多个进程：
   ```powershell
   Get-Process wps | Select-Object Id, MainWindowTitle
   ```
2. 如果有多个, **手动**关掉所有 WPS 窗口（不要用 `Stop-Process -Force`, 用窗口右上角 X）
3. 重新启动 WPS, 等 5 秒
4. 切到「公文助手」tab

### 4.3 IsVip patch 失败

打开 [踩坑全集.md](踩坑全集.md) #19 / #22 节, 通常是 GAC 残留 / Wow6432Node 不全。

或直接：

```powershell
# 卸载干净后重装
PowerShell -ExecutionPolicy Bypass -File D:\GongwenAssistant\installer\uninstall.ps1
PowerShell -ExecutionPolicy Bypass -File D:\GongwenAssistant\installer\install.ps1
```

### 4.4 patcher 没加载

```powershell
PowerShell -ExecutionPolicy Bypass -File D:\GongwenAssistant\installer\verify.ps1
```

应该输出 `结果: PASS=15 FAIL=0`。如果有 FAIL, 看输出哪一项失败, 对照 [INSTALL.md](INSTALL.md) §故障排查 找解决办法。

---

## 第 5 步：装好之后能干什么 / 不能干什么

### 能（已验证）

- 启动 WPS 时顶部有 tab —— **OK**
- VIP 字样显示「已激活」—— **OK**
- ribbon 上看到红头模板 / 范文搜索 / 排版按钮 —— **OK**
- 用 markdown + Pandoc 转 docx —— **OK**（不需要本插件）
- 用 python-docx 后处理排版 —— **OK**（不需要本插件）

### 不能（已知天花板）

- **点 ribbon 业务按钮没反应** —— WPS 12.1+ 内核屏蔽 COM Add-in onAction 回调, 不是本项目 bug
  - 详见 [踩坑全集.md #25](踩坑全集.md)
  - 替代方案见下文
- 模板搜索 / 范文搜索联网检索 —— 原作者服务器下线, 接口 404
- macOS / Linux —— COM Add-in 是 Windows 专属

### 替代方案（如果你只是想要公文输出）

| 你想要 | 用什么 | 在哪 |
|---|---|---|
| 一份具体公文 docx（F1 / F2） | 直接拿成品 | `审计工作20260521/F-公文版/` |
| 自己写 markdown 转 docx | Pandoc + python-docx | `审计工作20260521/script/` |
| 红头模板素材 | 仍在本插件里 | `runtime/template/` |

---

## 第 6 步：怎么卸载干净

```powershell
PowerShell -ExecutionPolicy Bypass -File D:\GongwenAssistant\installer\uninstall.ps1
```

会清理：

- `%LOCALAPPDATA%\GongwenAssistant\` 整个文件夹
- HKCU 下所有本项目相关注册表项

不会动：

- 你的任何文档
- WPS 本体
- 系统其他注册表

---

## 第 7 步：常见问题（看完就走的人）

**Q: 为什么按按钮没反应？**  
A: WPS 12.1+ 屏蔽了 COM Add-in 的回调, 这是 WPS 的安全策略, 我们绕不过。详见 [FAQ.md](FAQ.md) Q2。

**Q: 杀软（火绒 / 360）报警怎么办？**  
A: 把 `%LOCALAPPDATA%\GongwenAssistant\` 整个目录加白名单。详细分析见 [SECURITY.md](SECURITY.md)。

**Q: 我能信任这个项目吗？**  
A: 所有源码在 `src/` 公开, 所有副作用在 `installer/*.ps1` 明文, 你可以逐字 review。详见 [SECURITY.md](SECURITY.md) §自证清单。

**Q: 装完用不了我能恢复原版吗？**  
A: 能。原版 `Local_Wps_Vsto.dll` 在 GAC, 我们没动。卸载本项目后原版仍可用（只是依然受 VIP 拦截）。

**Q: 有没有视频教程？**  
A: 没有, 因为我们认为 5 分钟文字 + 截图就够了。如果你看完仍卡住, 直接提 Issue。

---

## 第 8 步：还想了解更多？

- 看原理 → [原理分析.md](原理分析.md)
- 看注册表细节 → [INSTALL.md](INSTALL.md)
- 看试错全过程 → [工程复盘.md](工程复盘.md) + [踩坑全集.md](踩坑全集.md)
- 看名词解释 → [术语表.md](术语表.md)
- 看怎么贡献 → [CONTRIBUTING.md](CONTRIBUTING.md)
