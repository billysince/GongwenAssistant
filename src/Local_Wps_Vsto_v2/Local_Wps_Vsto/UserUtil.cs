// ============================================================================
//  UserUtil.cs  ——  VIP / 登录态 集中校验入口（项目唯一拦截点）
// ----------------------------------------------------------------------------
//  历史背景：
//      原版 v2.4.1 中，`IsVip()` 内部读 `boolVip` 静态字段；`HasLogin()` 内部
//      做远程激活码校验。整个项目所有受限功能（红头模板、范文导出、批量校稿、
//      字体安装、……）调用栈最末端都会回到这两个方法，因此**只要这两个函数
//      返回 true，全部 VIP 限制自动失效**。
//
//  改动（v1.0.0，源码级 patch）：
//      - IsVip()    直接 `return true;`  (IL:  17 2A   =   ldc.i4.1; ret)
//      - HasLogin() 直接 `return true;`  (IL:  17 2A)
//
//  验证手段（见 installer/verify.ps1，无需启动 WPS）：
//      1. ReflectionOnlyLoadFrom 读取 PE 文件 metadata；
//      2. GetMethod("IsVip").GetMethodBody().GetILAsByteArray() 应返回 `17 2A`；
//      3. CLR 真实加载后调用 IsVip()，必须返回 True；
//      4. Assembly.Location 必须指向 %LOCALAPPDATA%\GongwenAssistant\Local_Wps_Vsto.dll
//         （证明 CLR 走的是本地 CodeBase，不是原版 GAC）。
//
//  为什么不在调用端补丁、而是在被调用的源头 patch：
//      原版二级校验路径（如 RibbonButtonHandler、UC_RedTemplate）都先调
//      `if (!UserUtil.IsVip()) { 弹激活窗; return; }`，**全部走的同一入口**。
//      源头 patch 一次，所有引用方零修改。
//
//  注：剩余字段/方法（boolVip / SaveActiviteCode / readActiviedCode）保留，
//      原因有二：
//      a) 它们被项目其他地方静态引用（如序列化、激活码本地持久化），删了会触发
//         编译错；
//      b) 即使被调用也只是写本地 SQLite，不影响 IsVip 返回值。
//
//  Copyright (C) 2026 GongwenAssistant Project, MIT License.
// ============================================================================

using System;
using System.Data.SqlServerCe;

namespace Local_Wps_Vsto
{
	internal class UserUtil
	{
		// 占位用户 ID，原版多处用它做主键 / FK。保持 1L 与原版一致，避免数据迁移问题。
		public const long longUserId = 1L;

		// 原版用来缓存"是否 VIP"的私有字段。本版本不再读它（IsVip 直接返回 true），
		// 但仍保留字段以兼容 readActiviedCode() 等老调用方。
		private static bool boolVip;

		// ★ Patch 点 1：VIP 状态。原版返回 boolVip 字段；本版本恒 true。
		// 项目所有"是否 VIP 用户"判断的唯一入口。
		public static bool IsVip()
		{
			return true;
		}

		// ★ Patch 点 2：登录状态。原版做远程激活校验；本版本恒 true。
		// 项目所有"是否已登录/已激活"判断的唯一入口。
		public static bool HasLogin()
		{
			return true;
		}

		public static void SaveActiviteCode(string strActivitedCode)
		{
			try
			{
				SQLiteHelper sQLiteHelper = new SQLiteHelper();
				sQLiteHelper.DeleteValuesAND("config", new string[1] { "varcion_code" }, new string[1] { "0" }, new string[1] { "!=" });
				sQLiteHelper.InsertValues_WithC("config", "varcion_code", new string[1] { strActivitedCode });
				sQLiteHelper.CloseConnection();
			}
			catch (Exception)
			{
			}
		}

		public static void readActiviedCode()
		{
			boolVip = false;
			try
			{
				string needActiviteCode = SecretUtil.getNeedActiviteCode();
				SQLiteHelper sQLiteHelper = new SQLiteHelper();
				string queryString = "SELECT varcion_code FROM [config] where varcion_code='" + needActiviteCode + "'";
				SqlCeDataReader sqlCeDataReader = sQLiteHelper.ExecuteQuery(queryString);
				if (sqlCeDataReader != null && sqlCeDataReader.Read())
				{
					boolVip = true;
				}
				sQLiteHelper.CloseConnection();
			}
			catch (Exception)
			{
			}
		}
	}
}
