# 公文助手 文档体系 INDEX

> 本仓库 docs/ 目录共 15 份 markdown 文档, 按受众与目的整理如下。  
> 不知道看哪份？请先看 [学习路径.md](学习路径.md)（按时间预算给四档阅读路线）。

---

## 一、按受众分类

### 给完全 0 基础的用户

| 文件 | 用途 | 阅读时间 |
|---|---|---|
| [../README.md](../README.md) | 项目门面, 一句话讲清是什么 | 5 分钟 |
| [QUICKSTART.md](QUICKSTART.md) | 不懂代码也能跟着做完安装 | 10 分钟 |
| [FAQ.md](FAQ.md) | "为什么按按钮没反应""为什么杀软报警" | 边看边查 |
| [术语表.md](术语表.md) | COM Add-in / IL / Harmony / GAC / CLSID 大白话 | 20 分钟 |
| [学习路径.md](学习路径.md) | 5 分钟 / 30 分钟 / 2 小时 / 半天 四档路线 | 5 分钟选路 |

### 给警惕用户（杀软报警 / 怀疑安全）

| 文件 | 用途 |
|---|---|
| [SECURITY.md](SECURITY.md) | 自证不是病毒, 副作用清单, 杀软配置 |
| [INSTALL.md](INSTALL.md) | 注册表项 / 文件路径 / 加载链路 全明文 |

### 给想理解原理的工程师

| 文件 | 用途 |
|---|---|
| [原理分析.md](原理分析.md) | WPS COM Add-in 加载链路解构, mscoree.dll → AppDomain → Local_Wps_Vsto |
| [Patcher方案.md](Patcher方案.md) | Harmony IL hook 的方案设计、为什么走这条路 |
| [ARCHITECTURE.md](ARCHITECTURE.md) | 整个仓库目录结构 / 模块依赖 / 关键决策点 |

### 给想 fork / 改进的开发者

| 文件 | 用途 |
|---|---|
| [CONTRIBUTING.md](CONTRIBUTING.md) | issue / PR 流程, 怎么读源码 |
| [工程复盘.md](工程复盘.md) | 决策点 + 失败重试 完整时间线 |
| [踩坑全集.md](踩坑全集.md) | 26 条具体踩坑, 包含现象/假设/验证/根因/教训 |

### 给运维 / 部署人员

| 文件 | 用途 |
|---|---|
| [部署演练指南.md](部署演练指南.md) | 在多种环境上的部署 dry-run 步骤 |
| [CHANGELOG.md](CHANGELOG.md) | 版本变更记录, 每个版本解决了什么、踩了什么 |

---

## 二、按问题分类

| 你的问题 | 看哪份 |
|---|---|
| 这是干嘛的？ | [README.md](../README.md) §一 |
| 现在能用吗？ | [README.md](../README.md) §二（状态卡片） |
| 为什么我看不到 ribbon tab？ | [FAQ.md](FAQ.md) Q1 + [踩坑全集.md](踩坑全集.md) #25 |
| 为什么按钮按了没反应？ | [FAQ.md](FAQ.md) Q2 + [踩坑全集.md](踩坑全集.md) #25 |
| 为什么杀软报警？ | [SECURITY.md](SECURITY.md) |
| 我能在 macOS 上用吗？ | [README.md](../README.md) §七 |
| 这玩意是不是后门？ | [SECURITY.md](SECURITY.md) §三 + [README.md](../README.md) §十一 |
| 怎么卸载干净？ | [README.md](../README.md) §六 + [INSTALL.md](INSTALL.md) §卸载 |
| Harmony 是什么？ | [术语表.md](术语表.md) 词条 Harmony |
| 我怎么改源码？ | [CONTRIBUTING.md](CONTRIBUTING.md) |
| 你们走过哪些弯路？ | [踩坑全集.md](踩坑全集.md) + [工程复盘.md](工程复盘.md) |
| 改完怎么验证有没有破坏？ | [INSTALL.md](INSTALL.md) §verify.ps1 |

---

## 三、按 docs/ 目录全文清单

```
docs/
├── INDEX.md              ← 本文件
├── README.md (在仓库根)  ← 项目门面
├── 学习路径.md            ← 0 基础导航
├── 术语表.md              ← 名词大白话
├── QUICKSTART.md         ← 5 分钟上手
├── FAQ.md                ← 常见问题
├── SECURITY.md           ← 安全声明
├── CONTRIBUTING.md       ← 贡献指南
├── INSTALL.md            ← 详细安装手册
├── ARCHITECTURE.md       ← 系统设计
├── 原理分析.md            ← COM Add-in 加载链路
├── Patcher方案.md         ← Harmony IL hook 方案
├── 工程复盘.md            ← 决策点时间线
├── 踩坑全集.md            ← 26 条踩坑
├── 部署演练指南.md        ← 部署 dry-run
└── CHANGELOG.md          ← 版本记录
```

---

## 四、文档维护原则

1. **透明开源** — 所有失败、所有走错的弯路、所有当前未解决的限制全部如实写, 不藏拙
2. **0 基础友好** — 第一次出现的术语必须在 [术语表.md](术语表.md) 有大白话解释
3. **可追溯** — 每个技术结论都附验证步骤, 任何人复制粘贴能复现
4. **稳定优先** — 已交付的内容不删, 修订用 CHANGELOG 增量追加
5. **Markdown 标准** — UTF-8 BOM 编码, 中英文混排空格用半角, 二级标题以上必须有锚点

---

## 五、当前已完成的事 vs 未完成的事

### 已完成

- 反编译重建 → 60+ 文件 C# 可读源码
- 弱命名重编译 → 1,408,512 字节 dll
- HKCU 部署链路 → 不动 HKLM / GAC, 无需管理员
- Harmony IL 注入 → IsVip / HasLogin patch 成功（log 可证）
- VIP 字样改写 → 「终身VIP - 已激活」可见
- ribbon tab 显示 → 「公文高手单机版2.4.1」可见
- F1 / F2 公文 docx 直接交付 → 见 `审计工作20260521/F-公文版/`

### 未完成（已知天花板）

- ribbon 业务按钮 onAction 在 WPS 12.1.0.17xx+ 个人版无法响应
  - 根因：WPS 内核安全策略屏蔽 COM Add-in 的运行时回调
  - 详见 [踩坑全集.md #25](踩坑全集.md)
  - 这是 WPS 内核行为, 不是本项目可控范围
- JS Add-in 路线（B 路）在 WPS 12.1.0.16910+ 个人版被 dev 模式加载砍
  - 详见 [踩坑全集.md #26](踩坑全集.md)
- macOS / Linux 不支持（COM Add-in 是 Windows 专属技术栈）

### 替代方案

- 如果你只想要具体公文 docx → `审计工作20260521/F-公文版/` 已有 F1 / F2 成品
- 如果你想自己写 markdown 转 docx → 用 Pandoc + python-docx
- 如果你需要 ribbon 上点按钮触发功能 → 等 WPS 内核未来开放, 或迁移到 MS Office Word（COM Add-in 链路在 Word 上仍然完整）
