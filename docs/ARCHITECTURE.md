# 系统架构 (ARCHITECTURE.md)

本文档说明公文助手在运行时如何被 WPS Office 加载，以及代码层的核心模块划分。

---

## 一、运行时全景

```
+-----------------+        +----------------------------------+
|   wps.exe       |        | %LOCALAPPDATA%\GongwenAssistant\ |
|  (WPS 主进程)   |        |   Local_Wps_Vsto.dll  (弱命名)   |
+--------+--------+        |   Newtonsoft.Json.dll            |
         |                 |   Spire.Doc.dll  ...             |
         | (1) WPS 启动时  +----------------------------------+
         |    扫描 HKCU                ^
         |    Software\Microsoft\Office|
         |    \Word\Addins\            |
         |    枚举所有 ProgId          |
         v                             |
+--------+--------+                    |
|  COM Activation |                    |
|   (mscoree.dll) |                    |
+--------+--------+                    |
         |                             |
         | (2) ProgId -> CLSID         |
         |     HKCU\Software\Classes\  |
         |     CLSID\{2CD4E522-...}\   |
         |     InprocServer32          |
         |       CodeBase = file:///...|
         v                             |
+--------+--------+                    |
|  CLR Hosting    |                    |
|  (.NET 4.x)     +--------------------+
|  ApplicationBase|  (3) CLR 通过 CodeBase
|  Probing        |      从本地路径加载弱命名 dll
+--------+--------+
         |
         | (4) Activator.CreateInstance(Local_Wps_Vsto.MyAddin)
         v
+--------+----------+
| Local_Wps_Vsto.dll|
|  MyAddin 类实例   |
|  - IDTExtensibility2.OnConnection   ← WPS 注入 Application 对象
|  - IRibbonExtensibility.GetCustomUI ← 返回 Resource1.MyRibbon (XML)
|  - GetLabel / GetImage / GetVisible ← WPS 回调
|  - btn*_Click 事件处理              ← 用户点击按钮
+-------------------+
```

---

## 二、关键链路与对应注册表

| 步骤 | WPS / CLR 行为 | 依赖的注册表 / 文件 |
|------|----------------|---------------------|
| 1. 启动扫描 | WPS 读取 HKCU Word\Addins 与 AddinsWL | `HKCU\Software\Microsoft\Office\Word\Addins\Local_Wps_Vsto.MyAddin` <br> `HKCU\Software\Kingsoft\Office\WPS\AddinsWL` |
| 2. ProgId → CLSID | COM 把 ProgId 字符串解析为 CLSID GUID | `HKCU\Software\Classes\Local_Wps_Vsto.MyAddin\CLSID` |
| 3. CLSID → 加载方式 | CLSID 对应 `mscoree.dll` 表示是 .NET COM Add-in | `HKCU\Software\Classes\CLSID\{2CD4E522-...}\InprocServer32` |
| 4. CLR 定位 dll | CLR 用 `Assembly` 字段做 binding，遇到 `PublicKeyToken=null` + `CodeBase=file:///...` 时直接走 CodeBase | 同上键的 `Assembly` / `CodeBase` 字段 |
| 5. 加载弱命名 dll | CLR 加载本地 dll，由于无强名，**不会进入 GAC 优先路径** | `%LOCALAPPDATA%\GongwenAssistant\Local_Wps_Vsto.dll` |
| 6. 实例化 AddIn | CLR 调用 `Local_Wps_Vsto.MyAddin` 默认构造函数 | dll 内 `MyAddin.cs` |

> **为什么强调「弱命名 + CodeBase」？** 原版 dll 是强名签名 (PublicKeyToken=475a36b05c42bd98)。强名 + 已注册到 GAC 时，CLR 会优先从 GAC 加载，覆盖 ApplicationBase 与 CodeBase。所以**本项目编译产物必须是弱命名**，并且**安装时必须先确保 GAC 中没有同名条目**（`gacutil /u Local_Wps_Vsto`）。

---

## 三、源码模块划分

```
src/Local_Wps_Vsto_v2/
├── Local_Wps_Vsto/                          # 业务代码 (namespace Local_Wps_Vsto)
│   ├── MyAddin.cs                ?── 入口类: 实现 IDTExtensibility2 + IRibbonExtensibility
│   ├── UserUtil.cs               ?── ★ 用户态 / VIP 校验, IsVip/HasLogin 已 patch 为 ret true
│   ├── CommonConfig.cs           ?── 全局常量 (产品名/版本号/路径/价格/提示文案)
│   ├── ConfigFile.cs / ConfigLoad.cs / ConfigSet.cs   ← 用户偏好读写
│   │
│   ├── HttpUtil.cs               ?── 原版网络通讯封装, 服务端已下线, 所有调用 catch 静默
│   ├── UpdateUtil.cs             ?── 原版自动更新, 服务端已下线
│   ├── WebMessageUtil.cs         ?── 原版站内消息推送, 服务端已下线
│   │
│   ├── SQLiteHelper.cs           ?── SQL CE 本地数据库 (conf/gwgs.sdf 中存模板索引/历史)
│   ├── FileCacheUtil.cs / MemCacheUtil.cs     ← 缓存层
│   │
│   ├── DocWpsUtil.cs / WpsUtil.cs / WpsHelper.cs       ← Word/WPS Range/Selection 操作工具
│   ├── CorrectUtil.cs            ?── 校稿核心 (朗读 + 提纲检查)
│   ├── RtfUtil.cs                ?── RTF 资源解析
│   ├── ImproveUtil.cs            ?── 写作提词器
│   │
│   ├── Uc_*.cs / *Tips.cs / MesBoxShow.cs / OnDoing.cs ← WinForms UI (用户控件)
│   │                                                      合并了原 *.Designer.cs
│   ├── FanWenItem.cs / SucaiItem.cs / ResultItem.cs ...  ← 数据 POCO
│   └── MachineCode.cs / SecretUtil.cs / FingerPrint.cs   ← 原版机器码 / 加解密 (仍存在，但 IsVip 已直通)
│
├── Word/  Excel/  Office/  AddInDesignerObjects/    ?── 嵌入式 PIA 接口 (TypeIdentifier)
│   └── *.cs                      ?── 反编译自原 dll 的 metadata, 由 ILSpy 输出
│
├── Properties/AssemblyInfo.cs    ?── AssemblyTitle/Product 已改为 GongwenAssistant
└── *.resx                        ?── Ribbon XML / 字符串 / 二进制资源
    ├── Local_Wps_Vsto.Resource1.resx        ← MyRibbon XML (顶部选项卡 UI 定义)
    ├── Local_Wps_Vsto.Properties.Reso.resx  ← 各类 RTF / 图标 / 公文模板
    ├── Local_Wps_Vsto.MesBoxShow.resx
    └── Local_Wps_Vsto.WriterTips.resx
```

---

## 四、Ribbon XML 与回调约定

资源 `Local_Wps_Vsto.Resource1.resx` 中名为 `MyRibbon` 的字符串 = Office Custom UI XML。其顶层结构：

```xml
<customUI xmlns="..." onLoad="MyAddInInitialize">
  <ribbon>
    <tabs>
      <tab id="gwgs" insertBeforeMso="TabHome" getLabel="GetLabel">
        <group id="gUser"        label="公文助手">
        <group id="group3"       label="素材/范文">
        <group id="gWps"         label="排版">
        <group id="group1"       label="插入">
        <group id="groupCorrect" label="校稿">
        <group id="groupOther"   label="其他">
      </tab>
    </tabs>
  </ribbon>
</customUI>
```

每个 `<button onAction="btnXxx_Click">` 都映射到 `MyAddin.btnXxx_Click(IRibbonControl)` 方法。`getLabel="GetLabel"` 表示按钮文字由代码回调 `MyAddin.GetLabel(control)` 动态返回。`getImage="GetRibbonImage"` 同理。

> 原版用 `<group id="gUser" label="尚未激活">` 在加载时显示「尚未激活」红色提示。本项目将该 label 改为「公文助手」，配合 `IsVip()` 已恒返回 `true`，运行时 `GetLabel(btnUserMsg)` 返回的 `strLbUserMsg` 也已被改为「公文助手」。

---

## 五、第三方依赖说明

| 依赖 | 用途 | 版权 |
|------|------|------|
| `Newtonsoft.Json.dll` | JSON 解析 (本地配置 / 搜索结果) | MIT |
| `Spire.Doc.dll` | Word 文档操作 (导出 PDF 等) | 商业, e-iceblue 公司 |
| `Spire.Pdf.dll` | PDF 操作 | 商业, e-iceblue 公司 |
| `Spire.License.dll` | Spire 系列许可证 | 商业 |
| `System.Data.SqlServerCe.dll` + `amd64/x86 native` | SQL CE 4.0 本地数据库 | Microsoft, 已停止维护 |
| `ICSharpCode.SharpZipLib.dll` | zip / gzip | MIT |
| `Microsoft.mshtml.dll` | mshtml COM 互操作 (原版 WebBrowser 用) | Microsoft 自带 |
| `RibbonControlsLibrary.dll` | WPF Ribbon 控件库 (部分对话框用) | Microsoft 引用程序集 |
| `EntityFramework.dll` | ORM (历史代码引用, 现已无主动调用) | Apache 2.0 |

第三方商业组件以**二进制依赖**的形式存在于 `runtime/`，本项目**不重新分发其源码**，不影响其许可证状态。如果您要把本项目用于商业场景，请自行确认 Spire 系列产品的授权。

---

## 六、与原版的差异清单

| 维度 | 原版 v2.4.1 | 公文助手 v1.0.0 |
|------|-------------|-----------------|
| `Local_Wps_Vsto.dll` 大小 | 1,408,000 bytes | 1,408,512 bytes (新编译) |
| 强名签名 | PublicKeyToken=475a36b05c42bd98 | 无 (弱命名) |
| 安装 / 注册 | HKLM CLSID + 写入 GAC, 需要管理员 | 仅 HKCU, 无需管理员 |
| `UserUtil.IsVip()` | 读 SQL CE 中激活码 → 校验签名 | 直接 `return true` |
| `UserUtil.HasLogin()` | 读 SQL CE 用户表 → 判断是否登录 | 直接 `return true` |
| Ribbon 默认 label | "尚未激活" 红色 | "公文助手" |
| `GetLabel(btnUserMsg)` 返回 | "软件未激活" / "已激活" | 恒为 "公文助手" |
| 产品元数据 | AssemblyTitle="Wps_Vsto" | AssemblyTitle="GongwenAssistant" |
| 在线服务 | 调用 `http://www.6dgww.com/*` 等 | 同代码保留, 但服务端已下线, 全部失败静默 |

---

## 七、构建管线

```
原 dll (1,408,000 bytes, 强名)
    │
    │ ILSpy 7.2.1 + ICSharpCode.Decompiler
    │ tools/decompile.ps1
    ▼
src/Local_Wps_Vsto_v2/  (156 个 .cs 文件 + 4 个 .resx)
    │
    │ 手工修复 6 处反编译瑕疵
    │ (CS1737 / CS0596 见 docs/工程复盘.md)
    │
    │ 改名 / 文案 / AssemblyInfo
    ▼
src/Local_Wps_Vsto_v2/Local_Wps_Vsto.csproj  (旧 msbuild 4.0 格式)
    │
    │ C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe
    │ /p:Configuration=Release
    ▼
src/Local_Wps_Vsto_v2/bin/Release/Local_Wps_Vsto.dll  (1,408,512 bytes, 弱命名)
    │
    │ Copy-Item → runtime/
    ▼
runtime/Local_Wps_Vsto.dll  + 全部第三方依赖 + 模板资源
    │
    │ installer/install.ps1 (HKCU 注册 + 复制到 %LOCALAPPDATA%)
    ▼
WPS 重启 → 顶部 Ribbon 看见「公文助手 1.0.0」选项卡, 全部功能可用
```
