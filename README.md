# 公文助手 (GongwenAssistant)

> 一个让 WPS 文字处理多出一组「公文写作」按钮的小工具。  
> 基于 2024 年某商业插件「公文助手 WPS 插件单机版 2.4.1」的本地代码做透明化重建，去除 VIP 拦截、保留全部业务功能、所有源码可读可改。

---

## 一、这是什么？(30 秒版)

打开 WPS 文字 → 顶部菜单栏多一个「公文助手单机版2.4.1」选项卡 → 里面有 30+ 个按钮：

- 设为 A4 / 一键排版 / 各级标题 / 红头标题 / 文末日期
- 红头模板库 / 范文库 / 素材库 / 写作提词器
- 朗读校稿 / 提纲检查 / 一键导出 PDF

适合：需要按 GB/T 9704-2012 写党政公文的人、街道乡镇办事员、单位文秘、审计岗位。

不需要：写学术论文、商业合同、营销文案的人。

---

## 二、现在能用吗？(状态卡片 · 2026-05-22 11:50 v1.0.6)

按"装载 / 渲染 / 功能"三层独立评估:

| 层 | 项目 | 状态 | 说明 |
|---|---|---|---|
| 装载层 | v2 弱命名版加载 | **OK** | 通过 HKCU 双视图 + 剥离 HKLM 36 个原版 CLSID, .NET CLR Fusion 加载 PKT=null 的 v2 dll. 见 [踩坑全集.md #27](docs/踩坑全集.md) |
| 装载层 | Patcher (v1.0.6) | **OK** | 11264 bytes, 含 IsVip/HasLogin patch + click trace + 异常 finalizer. 已部署 LOCALAPPDATA Patcher\ |
| 渲染层 | Ribbon tab 显示 | **OK** | 「**公文助手 1.0.0**」(v2 ribbon XML 内嵌, 不再是原版的"公文助手单机版2.4.1") |
| 渲染层 | 终身VIP 字样 | **OK** | v2 源码级 IsVip→return true (v2 编译时已写死, Patcher 是双保险) |
| 渲染层 | ribbon 按钮全部渲染 | **OK** | 约 30 个按钮可见 (见 dist/wps_v2_success.png) |
| **功能层** | **点击 ribbon 按钮** | **受限** | WPS 12.1+ 内核屏蔽 COM Add-in onAction 回调 ([#25](docs/踩坑全集.md)). v1.0.6 加 click trace 后将给出**硬证据**: 看 patcher.log 是否出现 `CLICK <Type>.<Method>` 行 |
| **功能层** | **按钮点击不再崩溃 WPS** | **应已修复** | v1.0.6 finalizer 吞掉 handler 内任何异常, 不让冒泡到 WPS 触发崩溃恢复界面 ([#28](docs/踩坑全集.md)). 用户重启 WPS 验证 |
| 功能层 | F1/F2 公文 docx 直接生成 | **OK** | 完全不依赖 ribbon, 见 `审计工作20260521/F-公文版/` 已交付成品 |
| 兼容性 | Word(MS Office) 加载 | 未测试 | 设计支持, 无环境验证 |
| 兼容性 | Win7 / WPS 旧版本 | 未测试 | 11.x 时代 onAction 是通的, 但已无环境验证 |

**v1.0.6 关键变化**:
- 撤销 v1.0.5 的"还原原版"操作, 让 v2 ribbon 重新生效 (这就是用户记忆里"那次成功的"状态)
- patcher 加 click trace + 异常 finalizer, 解决"按钮一点就崩溃恢复界面"问题
- 用户工作流: ribbon 美观可看 + IsVip 解锁 + 点击不再崩 + 实际写公文走 docx 路径不依赖 ribbon

**实事求是结论**:
- **能看 ribbon、不能用 ribbon 触发业务功能** (除非 patcher.log 未来真的看到 CLICK 行)
- **写公文请走 docx 路径**: `审计工作20260521/F-公文版/F1*.docx F2*.docx` 已成品交付; 后续公文用 `审计工作20260521/script/md2docx_gbt9704.py` 把 markdown 转 docx
- **本项目对 0 基础用户的价值**: 学习 .NET COM Add-in 逆向 / Harmony IL hook / WPS 12.1+ 安全策略演进的活案例, 见 [docs/INDEX.md](docs/INDEX.md)

---

## 三、看一眼真实截图

启动 WPS 后的真实样子（v2 ribbon 生效状态）：

![current_state](dist/wps_v2_success.png)

图中可见:
- ① 顶部 tab 栏「**公文助手 1.0.0**」 ← v2 重编译版的 ribbon XML
- ② 左下大图标「公文助手」上方红字「**终身VIP**」 ← v2 源码级 IsVip→return true
- ③ ribbon 区约 30 个按钮全部渲染（红头模板 / 素材搜索 / 范文搜索 / 导入资源 / 备份资源 / 恢复备份 / 写作提词器 / 模糊提示 / 5条提示 / 设为A4 / 一键排版 / 公文标题 / 公文署名 / 公文正文 / 公文附件 / 一级标题 / 二级标题 / 三级标题 / 高级 / 文末日期 / 发送对象 / 数字编号 / 符号 / 页码 / 横页 / 日期 / 朗读校稿 / 检查提纲...）

**这张图证明的边界**: 装载层(v2 dll 被加载) + 渲染层(tab + VIP 字样 + 30 按钮) 三个事实是真的. **它不能证明这些按钮可以被点击触发功能**, 那是另一码事 — 详见状态卡片"功能层"行与 [踩坑全集.md #25 #28 #29](docs/踩坑全集.md).

**如果你装完后看到 tab 名是「公文助手单机版2.4.1」**: 那是原版 GAC 强名 dll 被加载（HKLM\Wow6432Node 注册压过 HKCU 我们的）。**这个状态本身没坏**, ribbon 也照样渲染、Patcher 照样把 IsVip patch 成 true, 唯一区别是 tab 名是原版的字符串. 是否要剥离 HKLM 让 v2 生效, 取决于你的偏好 — 见 [踩坑全集.md #27](docs/踩坑全集.md), 或者跑 `bin/_reapply_v2.ps1` 一键执行（需 UAC elevate）.

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

**Step 3：启动 WPS, 切到「公文助手单机版2.4.1」tab**

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

我们没有藏失败。所有错误、所有走过的弯路都记在 [docs/踩坑全集.md](docs/踩坑全集.md), 当前 26 条主要条目：

- #1～#16 早期编码 / 反编译 / 部署相关
- #17～#22 Patcher 路径上 Harmony 依赖 / GAC 强名 / Wow6432Node 视图问题
- #23～#24 antivirus 误报 / 编译产物字节比对
- **#25 WPS 12.1+ 内核砍 COM Add-in onAction 回调** ← 项目天花板
- **#26 WPS 12.1+ 个人版砍 JS Add-in dev 模式加载** ← B 路也撞墙

每条都有：现象 / 假设 / 验证步骤 / 根因 / 教训。

---

## 十一、安全声明

- 本仓库**不打包**原版安装程序 `公文助手Wps插件单机版2.4.1.exe`, **不分发**原服务器接口
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

## 十四、联系

- Issue: 直接在 GitHub 上提
- 维护者：billysince
- 致谢：原作者「公文助手」团队（功能与 UI 设计）/ ILSpy / Harmony / Pandoc 各项目作者
