# 公文助手 贡献指南

> 这个项目欢迎所有人参与, 但请按下面流程来, 让我们的协作高效一点。

---

## 一、参与方式（按门槛由低到高）

| 方式 | 门槛 | 操作 |
|---|---|---|
| 提 Issue | 0 | GitHub 上点 New Issue |
| 改文档 | 低 | 在线编辑 + Pull Request |
| 报告 bug + 复现步骤 | 低 | Issue 用 bug 模板 |
| 提交 PR 改源码 | 中 | 见 §三 |
| 重写一个模块 | 高 | 先开 Issue 讨论方向 |
| Fork 另起项目 | 自由 | MIT 协议允许 |

---

## 二、提 Issue 的最佳实践

### 2.1 Issue 模板

我们没有强制模板, 但好 issue 包含这 6 项：

```markdown
## 现象（一句话）
打开 WPS 后, 公文助手 tab 看不到。

## 环境
- Windows 11 23H2 22631.4112
- WPS Office 12.1.0.21482（个人版）
- GongwenAssistant v1.0.3
- 安装方式：路线 A · Patcher

## 复现步骤
1. 解压 release zip 到 D:\GongwenAssistant\
2. 跑 PowerShell -ExecutionPolicy Bypass -File installer\install.ps1
3. 启动 WPS, 打开任意 docx
4. 顶部菜单栏看不到「公文助手」

## 期望行为
顶部菜单栏出现「公文助手」tab

## 实际行为
菜单栏只有「文件 / 开始 / 插入 / 页面 / 引用 / 审阅 / 视图」, 没有公文助手 tab

## 已查证据
- patcher.log 文件存在, 但只有 OnConnection start 记录, 没有 Patched IsVip
- verify.ps1 输出 PASS=14 FAIL=1, 失败项: "Local_Wps_Vsto detected"
- 杀软无告警

## 我尝试过的
- 卸载重装一次
- 重启电脑后再启动 WPS
- 关闭杀软重试
```

### 2.2 不要这样开 issue

| 错误示范 | 为什么错 | 怎么改 |
|---|---|---|
| 「不能用」 | 信息量为 0, 我们无法定位 | 用上面模板 |
| 「我是 0 基础, 你要负责教我」 | 文档已尽量友好, 完全 0 基础不属于本项目支持目标 | 先看 [QUICKSTART.md](QUICKSTART.md) + [FAQ.md](FAQ.md), 仍不会再问 |
| 「破解版还要钱？」 | 本项目从不收费, 但**不是破解版**, 是合规的反编译产物 | 看 [README.md §十一 §十二](../README.md) |
| 「商业授权报价」 | 我们没有原作者授权, 无法再授权他人 | 直接联系原作者团队 |

---

## 三、提交 PR 的流程

### 3.1 改之前先讨论

**改源码 / 加大模块前, 先开 Issue 描述思路**, 等维护者回复同意后再动手。否则可能：

- 你做完发现方向不对, 白干
- 跟其他人撞车, 重复劳动
- 不符合项目 scope（比如想加 Linux 支持, 但 COM Add-in 是 Windows 专属）

**改文档 / 修小 bug, 不需要预先讨论**, 直接 PR。

### 3.2 fork → branch → PR

```bash
# 1. fork 本仓库到你账号
# 2. clone 你的 fork
git clone https://github.com/YOUR_NAME/GongwenAssistant.git
cd GongwenAssistant

# 3. 创建分支（不要直接改 main）
git checkout -b fix/ribbon-callback-edge-case

# 4. 改你的代码
# ...

# 5. 提交（参考 §四 提交规范）
git add .
git commit -m "fix(patcher): handle null assembly name in OnAssemblyLoad"

# 6. push 到你的 fork
git push origin fix/ribbon-callback-edge-case

# 7. 在 GitHub 上开 PR
```

### 3.3 PR 描述模板

```markdown
## 改动概述
（一句话说清楚做了什么）

## 关联 Issue
Closes #123

## 改动详情
1. ...
2. ...

## 测试
- [ ] 跑了 verify.ps1, 输出 PASS=15 FAIL=0
- [ ] 装到 WPS XX 版本能看到 tab
- [ ] patcher.log 有 Patched IsVip 行

## 副作用
（有没有改文件结构 / 注册表项 / 配置）
```

---

## 四、提交规范

我们用 [Conventional Commits](https://www.conventionalcommits.org/) 风格：

```
<type>(<scope>): <subject>

[optional body]
[optional footer]
```

**type**：

| type | 用途 |
|---|---|
| `feat` | 新功能 |
| `fix` | 修 bug |
| `docs` | 改文档 |
| `refactor` | 重构（不改行为） |
| `test` | 改测试 |
| `chore` | 工程杂事（构建 / CI / 依赖） |
| `revert` | 回退某个 commit |

**scope**（本项目常用）：

- `patcher` — Harmony Patcher 相关
- `installer` — 安装脚本
- `runtime` — 业务 dll
- `decompile` — 反编译流程
- `docs` — 文档
- `verify` — 自检脚本

**示例**：

```
feat(patcher): add NeedAccountValidate to patch list
fix(installer): preserve other AddinsWL entries when adding ours
docs(faq): add Q14 about wps multi-process issue
refactor(runtime): extract template path to const
chore(deps): bump 0Harmony to 2.3.5
```

---

## 五、源码组织

```
GongwenAssistant/
├── src/
│   ├── Local_Wps_Vsto_v2/        ← 反编译产物 + 弱命名重编译
│   │   ├── Local_Wps_Vsto.csproj
│   │   ├── Local_Wps_Vsto/        ← 业务代码（60+ 文件）
│   │   ├── Properties/
│   │   ├── Word/ Excel/ Office/   ← PIA 嵌入类型
│   │   └── *.resx                 ← 资源
│   └── GongwenPatcher/             ← Harmony Patcher 自研代码（200 行内）
│       ├── GongwenPatcher.csproj
│       └── Connect.cs              ← 唯一文件
├── runtime/                        ← 发布运行时
├── installer/                      ← 安装 / 卸载 / 自检
├── tools/                          ← 反编译器 + 驱动脚本
└── docs/                           ← 文档
```

### 5.1 `src/GongwenPatcher/Connect.cs`

最小入口, 200 行内。改这里之前先看 [Patcher方案.md](Patcher方案.md) 了解设计意图。

关键 API：

- `OnConnection` — Add-in 启动入口
- `AppDomain.AssemblyLoad` 订阅 — 等 Local_Wps_Vsto 被加载
- `Harmony.Patch(IsVip, prefix=ReturnTrue)` — IL 替换
- `OnStartupComplete` — 清理 AddinsCL 防误禁用

### 5.2 `src/Local_Wps_Vsto_v2/`

反编译产物, **不建议主动修改**, 因为：

- 改了就不是"忠实于原版"了, 失去溯源
- 如果原作者将来恢复运营, 会冲突
- 维护成本极高（60+ 文件）

如果一定要改, 推荐用 Harmony 在 GongwenPatcher.dll 里 hook, 而不是改 Local_Wps_Vsto 源码。

---

## 六、文档贡献规范

### 6.1 编码

- 全部 markdown 用 **UTF-8 with BOM**（Cursor / VSCode 默认 UTF-8 不带 BOM, 写完后请用 PowerShell 转换, 看 [#16](踩坑全集.md)）
- 中英文混排空格用半角
- 代码块明确指定语言（如 ` ```powershell` ` ```csharp` ）

### 6.2 风格

- **0 基础友好** — 第一次出现的术语必须解释或链接到 [术语表.md](术语表.md)
- **可追溯** — 每个技术结论附验证步骤
- **如实** — 失败也写出来, 不藏拙

### 6.3 校验

```powershell
# 跑全部 markdown 文件 BOM 校验
Get-ChildItem -Path 'docs\', 'README.md' -Recurse -Include *.md | ForEach-Object {
    $bytes = [System.IO.File]::ReadAllBytes($_.FullName)
    if ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        "OK  $($_.FullName)"
    } else {
        "NO BOM $($_.FullName)"
    }
}
```

期望全 OK。

---

## 七、行为准则（Code of Conduct）

我们采用简单版：

1. 友好, 不人身攻击
2. 遵守事实, 不传播谣言
3. 尊重他人时间, 提 issue 前先 check 老 issue
4. 不发病毒 / 后门 / 无关广告
5. 协议合规

违反任一条 → 警告 → 拉黑。

---

## 八、维护者

- 维护者：**billysince**
- Issue 响应预期：72 小时（业余项目）
- PR 合并预期：1 周（含测试 + review）

---

## 九、致谢

历史贡献者：

- billysince（架构 + 全部初版代码 + 文档）
- ILSpy 团队（反编译器）
- pardeike（Harmony 框架）
- Pandoc 团队（md ? docx 转换）
- 「公文助手」原作者团队（功能 + UI 设计 + 模板素材）

未来贡献者会列在这里。
