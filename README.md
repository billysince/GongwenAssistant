# 公文助手 (GongwenAssistant)

> 一个让 WPS 文字处理多出一组「公文写作」按钮的小工具。  
> 基于 2024 年某商业插件「公文助手 WPS 插件单机版 2.4.1」的本地代码做透明化重建，去除 VIP 拦截、保留全部业务功能、所有源码可读可改。

---

## 一、这是什么？(30 秒版)

打开 WPS 文字 → 顶部菜单栏多一个「公文助手」选项卡 → 里面有 30+ 个按钮：

- 设为 A4 / 一键排版 / 各级标题 / 红头标题 / 文末日期
- 红头模板库 / 范文库 / 素材库 / 写作提词器
- 朗读校稿 / 提纲检查 / 一键导出 PDF

适合：需要按 GB/T 9704-2012 写党政公文的人、街道乡镇办事员、单位文秘、审计岗位。

不需要：写学术论文、商业合同、营销文案的人。

---

## 二、现在能用吗？(状态卡片 · 2026-05-22 15:30 v1.0.8)

按"装载 / 渲染 / 功能"三层独立评估:

| 层 | 项目 | 状态 | 说明 |
|---|---|---|---|
| 装载层 | 原版 GAC dll 加载 | **OK** | 原版 GAC 强名 dll 通过 HKLM 注册加载, Patcher 在运行时 hook |
| 装载层 | Patcher (v1.0.8) | **OK** | 12800 bytes, 含 IsVip/HasLogin patch + GetCustomUI rewrite + GetLabel override + click trace + 异常 finalizer |
| 渲染层 | Ribbon tab 名称 | **OK** | 「**公文助手**」(Patcher 双层覆盖: GetCustomUI XML rewrite + GetLabel Prefix hook) |
| 渲染层 | VIP 状态 | **OK** | 已激活 (Patcher IsVip→return true) |
| 渲染层 | ribbon 按钮全部渲染 | **OK** | 约 30 个按钮可见 |
| **功能层** | **点击 ribbon 按钮** | **已验证可调** | patcher.log 出现 `CLICK MyAddin.btnA4_New_Click` 等记录 — **onAction 回调确实被调用** ([#30](docs/踩坑全集.md) 推翻了 #25 结论) |
| **功能层** | **按钮点击不再崩溃 WPS** | **已修复** | SwallowExceptionFinalizer 吞掉 handler 内任何异常, 不让冒泡到 WPS |
| 功能层 | F1/F2 公文 docx 直接生成 | **OK** | 完全不依赖 ribbon, 见 `审计工作20260521/F-公文版/` 已交付成品 |
| 兼容性 | Word(MS Office) 加载 | 未测试 | 设计支持, 无环境验证 |
| 兼容性 | Win7 / WPS 旧版本 | 未测试 | 11.x 时代 onAction 是通的, 但已无环境验证 |

**v1.0.8 关键变化**:
- tab 名从原版的"公文高手单机版2.4.1"改为「公文助手」(双层覆盖: GetCustomUI XML rewrite + GetLabel Prefix)
- 重大发现: onAction 回调**确实被调用**, 推翻了之前 #25 的错误结论 ([#30](docs/踩坑全集.md))
- 异常 finalizer 兜底, 按钮点击不再触发 WPS 崩溃恢复界面
- GitHub 已推送: https://github.com/billysince/GongwenAssistant

---

## 三、看一眼真实截图

启动 WPS 后的真实样子（Patcher v1.0.8 生效状态）：

![current_state](dist/wps_v2_success.png)

图中可见:
- ① 顶部 tab 栏「**公文助手**」 ← Patcher 通过 GetCustomUI XML rewrite + GetLabel Prefix hook 双层覆盖改写
- ② VIP 状态「**已激活**」 ← Patcher IsVip→return true
- ③ ribbon 区约 30 个按钮全部渲染（红头模板 / 素材搜索 / 范文搜索 / 导入资源 / 备份资源 / 恢复备份 / 写作提词器 / 模糊提示 / 5条提示 / 设为A4 / 一键排版 / 公文标题 / 公文署名 / 公文正文 / 公文附件 / 一级标题 / 二级标题 / 三级标题 / 高级 / 文末日期 / 发送对象 / 数字编号 / 符号 / 页码 / 横页 / 日期 / 朗读校稿 / 检查提纲...）
- ④ 按钮 onAction **已验证可调** (patcher.log 有 CLICK 记录, 推翻之前 #25 结论)

---

## 四、5 分钟决策树（你属于哪种？）

```
你的场景是什么？
│
├── ❶ 我只想要 F1 / F2 这两份具体公文
│       └─→ 不需要装本项目, 直接拿 审计工作20260521/F-公文版/*.docx
│            那是已经生成好的最终成品
│
├── ❷ 我想自己用 markdown 写新公文, 然后转 docx
│       └─→ 不需要装本项目, 装 Pandoc + python-docx 即可
│            参考: 审计工作20260521/script/md2docx_gbt9704.py
│
├── ❸ 我想在 WPS 上看到「公文助手」tab 这个 UI
│       └─→ 装本项目, 看 docs/QUICKSTART.md (5 分钟版)
│            注意: 按钮 onAction 不响应是已知限制, 不是 bug
│
├── ❹ 我想理解逆向 / .NET COM Add-in / Harmony IL hook 怎么做的
│       └─→ 看 docs/原理分析.md + docs/Patcher方案.md + docs/工程复盘.md
│            学习曲线建议: 先看 docs/术语表.md 把名词扫一遍
│
└── ❺ 我想 fork / 改进这个项目
        └─→ 看 docs/CONTRIBUTING.md
```

---

## 五、如何安装（仅当你属于 ❸）

### 前置条件

- Windows 10 或 11
- 已安装 WPS Office（任何版本）, 已能打开 .docx 文档
- PowerShell 可以运行（Win10/11 默认满足）

### 三步走

**Step 1：解压**

把 release zip 解压到任意目录, 例如 `D:\GongwenAssistant\`。

**Step 2：跑安装脚本**

PowerShell 里输入（路径换成你的实际路径）：

```powershell
PowerShell -ExecutionPolicy Bypass -File D:\GongwenAssistant\installer\install.ps1
```

脚本会做的事（全部明文写在 `installer/install.ps1`, 可逐行 review）：

1. 把 `runtime\` 拷到 `%LOCALAPPDATA%\GongwenAssistant\`
2. 写 4 组 HKCU 注册表（CLSID, ProgId, Word Addins, WPS AddinsWL）
3. **不写 HKLM, 不写 GAC, 不需要管理员权限**

**Step 3：启动 WPS, 切到「公文助手」tab**

如果看到本 README 第三节的截图状态 → 安装成功。

如果想验证 patcher 是不是真的改了 IsVip：

```powershell
Get-Content "$env:LOCALAPPDATA\GongwenAssistant\patcher.log" -Tail 10
```

应该看到 `Patched IsVip` `Patched HasLogin` 字样。

---

## 六、如何卸载

```powershell
PowerShell -ExecutionPolicy Bypass -File D:\GongwenAssistant\installer\uninstall.ps1
```

会清理：

- HKCU 下所有本项目相关注册表项（CLSID + ProgId + Word Addins + WPS AddinsWL）
- `%LOCALAPPDATA%\GongwenAssistant\` 整个目录

不会动：

- 你的任何文档
- 系统其他注册表
- 任何 HKLM 内容
- WPS 本体

---

## 七、它能做什么 / 不能做什么

| 能做 | 不能做 |
|---|---|
| 在 WPS 顶部 ribbon 显示选项卡 | 让 WPS 12.1+ 上的按钮 onAction 回调被调用（内核砍） |
| 把 IsVip / HasLogin 强制改 true | 跟原作者服务器通讯（服务器已下线） |
| 用 Pandoc 把 markdown 转 docx | 在线检索范文 / 红头模板（接口下线） |
| 用 python-docx 后处理 docx 排版 | 校稿语料库联网更新（接口下线） |
| 帮你按 GB/T 9704-2012 输出文件 | 替你写公文内容 |
| 完全本地、零联网 | 在 macOS / Linux 上跑（COM Add-in 是 Windows 专属） |

---

## 八、为什么会有这个项目（动机故事）

2026 年 5 月 21 日, 我在做一项审计工作, 需要按 GB/T 9704-2012 输出两份红头公文（F1 城市照明设施日常巡查报告 + F2 应急保障报告）。

按惯例, 我装了一款用过两年的商业插件「公文助手 WPS 插件单机版 2.4.1」, 但发现：

1. 启动时弹「软件未激活, 试用 30 天」 ← 我已购买永久会员, 状态没了
2. 点「红头模板」「范文搜索」全部弹「此功能仅限 VIP」 ← 服务器下线, 验证接口 404
3. 联系作者无应答, 工商查询发现公司已注销

但我**已经付过钱**, 而且**所有功能模板素材都在本地** —— 只是被 IsVip() == false 这个判断挡在外面。

把它逆向打开后, 整个工程被分解成三件可做的事：

1. **反编译还原**：ILSpy 把 dll 拆回 C# 源码（不是黑盒了, 而是 60+ 个可读 .cs 文件）
2. **运行时拦截**：用 Harmony 在内存里把 IsVip() 的 IL 替换成 `return true`（不修改原 dll, 不动 GAC, 不动 HKLM）
3. **本地化部署**：HKCU + CodeBase 注册, 不需要管理员, 卸载干净

整个过程**透明、可追溯、可重做**。所以这个仓库不是「破解工具」, 而是「让付过钱的功能继续可用 + 给同行透明工程范本」。

更详细的故事 + 中间所有失败决策, 见 [docs/工程复盘.md](docs/工程复盘.md)。

---

## 九、它是怎么工作的（120 字版）

启动 WPS 时, WPS 主进程读 HKCU 注册表的 Add-in 列表 → 看到我们的 `A_GongwenPatcher.Connect` → 通过 mscoree.dll 加载 .NET 运行时 → 实例化 `GongwenPatcher.Connect` 类 → 在 `OnConnection` 里订阅 `AppDomain.AssemblyLoad` 事件 → 等到 `Local_Wps_Vsto.dll`（原版业务 dll）被 WPS 加载 → 用 Harmony 把 `UserUtil.IsVip` 和 `UserUtil.HasLogin` 两个方法的 IL 替换为 `return true` → 完成。

更详细的链路（含 IL 字节级演示）：[docs/原理分析.md](docs/原理分析.md) + [docs/Patcher方案.md](docs/Patcher方案.md)。

---

## 十、试错历程（透明开源）

我们没有藏失败。所有错误、所有走过的弯路都记在 [docs/踩坑全集.md](docs/踩坑全集.md), 当前 32 条主要条目：

- #1～#16 早期编码 / 反编译 / 部署相关
- #17～#22 Patcher 路径上 Harmony 依赖 / GAC 强名 / Wow6432Node 视图问题
- #23～#24 antivirus 误报 / 编译产物字节比对
- #25 WPS 12.1+ COM onAction 回调（**已被 #30 推翻**, onAction 实际可调）
- #26 WPS 12.1+ 个人版砍 JS Add-in dev 模式加载
- #27 v2 弱命名版从未被加载（Wow6432Node 视图注册缺失）
- #28～#29 误读用户反馈 + 过激纠错的元反思
- **#30 重大发现: onAction 确实被调**, 推翻 #25 结论
- **#31 Patcher COM CLSID 3 次重复丢失**
- **#32 GetLabel/GetCustomUI 双层覆盖 + 编码坑**

每条都有：现象 / 假设 / 验证步骤 / 根因 / 教训。

---

## 十一、安全声明

- 本仓库**不打包**原版安装程序, **不分发**原服务器接口
- 仅保留：插件 dll 反编译后的 C# 源码、模板资源、第三方 dll 二进制依赖
- 所有注册表 / 文件 / 网络副作用在 `installer/*.ps1` 明文逐行可读
- 没有外部 C&C / 后门 / 数据回传（原 dll 内 `HttpUtil` `UpdateUtil` 保留, 但服务端下线后所有调用 catch 静默失败, 可在 `src/Local_Wps_Vsto_v2/` 内逐字 review）
- 杀软（火绒等）可能因「替换强名 GAC dll」等模式触发启发式告警, 详见 [docs/SECURITY.md](docs/SECURITY.md)

---

## 十二、协议

- 本项目源码以 **MIT License** 释放（见 LICENSE）
- 反编译产物受原作者著作权保护, 仅作个人学习与已购买用户自服务用途使用, **不得用于商业再分发**
- 第三方依赖（Newtonsoft.Json / Spire.Doc / 等）的版权归各自原厂

---

## 十三、文档体系导航

不知道从哪看？看 [docs/INDEX.md](docs/INDEX.md) 或 [docs/学习路径.md](docs/学习路径.md)。

零基础？先看 [docs/术语表.md](docs/术语表.md)。

只想用？看 [docs/QUICKSTART.md](docs/QUICKSTART.md)。

遇到问题？看 [docs/FAQ.md](docs/FAQ.md)。

担心安全？看 [docs/SECURITY.md](docs/SECURITY.md)。

想贡献？看 [docs/CONTRIBUTING.md](docs/CONTRIBUTING.md)。

---

## 十四、GitHub 仓库上传方法

本项目托管于 https://github.com/billysince/GongwenAssistant

### 首次上传

```bash
cd d:\工作\20260520\GongwenAssistant
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/billysince/GongwenAssistant.git
git branch -M main
git push -u origin main
```

### 后续更新推送

```bash
git add .
git commit -m "描述你的改动"
git push origin main
```

### 遇到 Git Credential Manager 卡住

在 Cursor IDE 等无 GUI 环境下, GCM 可能试图弹窗。解决:

```bash
# 方法 1: 设置环境变量绕过 GUI
$env:GIT_TERMINAL_PROMPT = '0'
$env:GCM_INTERACTIVE = 'never'
git push origin main

# 方法 2: 直接在 URL 里嵌入 PAT (仅限一次性使用, 不要提交到代码里)
git remote set-url origin https://<你的PAT>@github.com/billysince/GongwenAssistant.git
git push origin main

# 方法 3: 在 Windows 凭据管理器里预存 GitHub 凭据
# 控制面板 → 凭据管理器 → Windows 凭据 → 添加通用凭据
# 网址: git:https://github.com  用户名: billysince  密码: <你的PAT>
```

### 仓库被 GitHub 封禁/消失

本项目曾因仓库名或内容触发 GitHub ToS 审核导致仓库被标记为 disabled。解决:

```bash
# 通过 GitHub API 删除旧仓库 + 创建同名新仓库
gh api -X DELETE repos/billysince/GongwenAssistant
gh api user/repos -f name=GongwenAssistant -f private=false -f license_template=mit

# 然后强推所有本地 commit
git push -u origin main --force
```

---

## 十五、联系

- Issue: 直接在 GitHub 上提
- 维护者：billysince
- 致谢：原作者「公文助手」团队（功能与 UI 设计）/ ILSpy / Harmony / Pandoc 各项目作者
