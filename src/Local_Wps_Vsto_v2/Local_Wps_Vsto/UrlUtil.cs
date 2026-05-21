namespace Local_Wps_Vsto
{
	internal class UrlUtil
	{
		private static string strBaseDomain_Base = "http://www.6dgww.com/";

		private static string strBaseDomain = "http://www.6dgww.com/gwbl/";

		public static string strImageUrl_GZH = "http://cache.gongwengaoshou.com/image/gzh_1.jpg?x-oss-process=image/resize,w_500";

		public static string strImageUrl_Pay_Ali = "http://cache.gongwengaoshou.com/image/pay_ali_n.jpg?x-oss-process=image/resize,w_500";

		public static string strImageUrl_Pay_Wx = "http://cache.gongwengaoshou.com/image/pay_wx_n.jpg?x-oss-process=image/resize,w_500";

		public static string strGetUpdateVersionUrl = "http://www.6dgww.com/gwbl/get.version.php";

		public static string strGetUpdateVersionUrl_Vsto()
		{
			return string.Concat(strBaseDomain + "get.version.wps.php?", "version=", CommonConfig.strVersionCode);
		}

		public static string strGetUpdateVersionUrl_Wps()
		{
			return string.Concat(strBaseDomain + "get.version.wps.php?", "version=", CommonConfig.strVersionCode);
		}

		public static string strLoginUrl(string strEmail, string strPwd)
		{
			return string.Concat(string.Concat(string.Concat(strBaseDomain + "login.20200823.php?", "email=", strEmail), "&pwd=", strPwd), "&mac=", SecretUtil.GetMachineCode());
		}

		public static string strImageUrl_Pay_Ali_XOR()
		{
			return string.Concat(string.Concat(string.Concat(strBaseDomain + "get_img.xor.php?", "user_id=", 1L.ToString()), "&pay_type=alipay"), "&mac=", SecretUtil.GetMachineCode());
		}

		public static string strImageUrl_Pay_Ali_XOR_Forever()
		{
			return string.Concat(string.Concat(string.Concat(strBaseDomain + "get_img.xor.forever.php?", "user_id=", 1L.ToString()), "&pay_type=alipay"), "&mac=", SecretUtil.GetMachineCode());
		}

		public static string strImageUrl_Pay_Wx_XOR()
		{
			return string.Concat(string.Concat(string.Concat(strBaseDomain + "get_img.xor.php?", "user_id=", 1L.ToString()), "&pay_type=native"), "&mac=", SecretUtil.GetMachineCode());
		}

		public static string strImageUrl_Pay_Wx_XOR_Forever()
		{
			return string.Concat(string.Concat(string.Concat(strBaseDomain + "get_img.xor.forever.php?", "user_id=", 1L.ToString()), "&pay_type=native"), "&mac=", SecretUtil.GetMachineCode());
		}

		public static string strRegUrl(string strEmail, string strPwd)
		{
			return string.Concat(string.Concat(string.Concat(strBaseDomain + "reg.php?", "email=", strEmail), "&pwd=", strPwd), "&mac=", SecretUtil.GetMachineCode());
		}

		public static string strGetSuCaiUrl(string strSortName, string strSearchWord, string strTagName, int intStartPosition = 0)
		{
			return string.Concat(string.Concat(string.Concat(string.Concat(string.Concat(string.Concat(strBaseDomain + "getsucai.php?", "search=", strSearchWord), "&user_id=", 1L.ToString()), "&sort=", strSortName), "&tag=", strTagName), "&start=", intStartPosition.ToString()), "&mac=", SecretUtil.GetMachineCode());
		}

		public static string strGetWordCorrectUrl()
		{
			return strBaseDomain + "correct.get.php?";
		}

		public static string strGetSuCaiUrl_Super(string strSearchWord, string strSortName, string strTagName, int maxCount = 10, int gwSort = 0)
		{
			return string.Concat(string.Concat(string.Concat(string.Concat(strBaseDomain + "getsucai.super.php?", "search=", strSearchWord), "&sort_id=", gwSort.ToString()), "&maxcount=", maxCount.ToString()), "&mac=", SecretUtil.GetMachineCode());
		}

		public static string strGetTiaoZhanUrl(string strSearchWord)
		{
			return string.Concat(string.Concat(strBaseDomain + "gettiaozhan.php?", "search=", strSearchWord), "&mac=", SecretUtil.GetMachineCode());
		}

		public static string strUploadOneSucai()
		{
			return strBaseDomain + "upload.onesucai.php?";
		}

		public static string strUploadOneArticle()
		{
			return strBaseDomain + "upload.onearticle.php?";
		}

		public static string strGetSucaiPackage(string strUserId)
		{
			return string.Concat(string.Concat(strBaseDomain + "get.sucai.package.php?", "user_id=", strUserId), "&mac=", SecretUtil.GetMachineCode());
		}

		public static string strGetSucaiPackage(long longUserId)
		{
			return string.Concat(string.Concat(strBaseDomain + "get.sucai.package.php?", "user_id=", longUserId.ToString()), "&mac=", SecretUtil.GetMachineCode());
		}

		public static string strGetSucaiPackage()
		{
			return string.Concat(string.Concat(strBaseDomain + "get.sucai.package.php?", "user_id=", 1L.ToString()), "&mac=", SecretUtil.GetMachineCode());
		}

		public static string strGetFanWenPackage()
		{
			return string.Concat(string.Concat(strBaseDomain + "get.fanwen.package.php?", "user_id=", 1L.ToString()), "&mac=", SecretUtil.GetMachineCode());
		}

		public static string strGetSucaiPackageDetail(string strPackageId, long longUserId)
		{
			return string.Concat(string.Concat(strBaseDomain + "get.sucai.package.detail.php?package_id=" + strPackageId, "&user_id=", longUserId.ToString()), "&mac=", SecretUtil.GetMachineCode());
		}

		public static string strGetFanWenPackageDetail(string strPackageId, long longUserId)
		{
			return string.Concat(string.Concat(strBaseDomain + "get.fanwen.package.detail.php?package_id=" + strPackageId, "&user_id=", longUserId.ToString()), "&mac=", SecretUtil.GetMachineCode());
		}

		public static string strGetSucaiPackageDetail(long longPackageId, long longUserId)
		{
			return string.Concat(string.Concat(strBaseDomain + "get.sucai.package.php?package_id=" + longPackageId, "&user_id=", longUserId.ToString()), "&mac=", SecretUtil.GetMachineCode());
		}

		public static string strQiuGaoUrl()
		{
			return "" + "http://www.6dgww.com/qiugao.php?code=" + SecretUtil.GetMachineCode();
		}

		public static string strPptUrl()
		{
			return "" + "http://www.6dgww.com/ppt.php?code=" + SecretUtil.GetMachineCode();
		}

		public static string strHomeUrl()
		{
			return "" + "http://www.6dgww.com/home.php?code=" + SecretUtil.GetMachineCode();
		}

		public static string strGetImrpoveUrl()
		{
			return string.Concat(string.Concat(strBaseDomain + "get.user.improve.new.php?", "&user_id=", 1L.ToString()), "&mac=", SecretUtil.GetMachineCode());
		}

		public static string strCheckCodeUrl(string strCode)
		{
			return string.Concat(string.Concat(string.Concat(strBaseDomain + "improve.code.php?", "&user_id=", 1L.ToString()), "&mac=", SecretUtil.GetMachineCode()), "&code=", strCode);
		}

		public static string strOpenImproveUrl()
		{
			return string.Concat(string.Concat(strBaseDomain_Base + "help.improve.gwgs.php?", "&user_id=", 1L.ToString()), "&mac=", SecretUtil.GetMachineCode());
		}

		public static string strGetMsgUrl()
		{
			return string.Concat(string.Concat(strBaseDomain + "get.msg.php?", "&user_id=", 1L.ToString()), "&mac=", SecretUtil.GetMachineCode());
		}

		public static string strUnLockPcUrl()
		{
			return "http://www.6dgww.com/" + "home.unlock.php?code=" + SecretUtil.GetMachineCode();
		}

		public static string strGoToQqUrl()
		{
			return strBaseDomain + "goto.qq.php?code=" + SecretUtil.GetMachineCode();
		}

		public static string strGoToHelpUrl()
		{
			return "http://www.6dgww.com/" + "help/v1000/index.vsto.php?code=" + SecretUtil.GetMachineCode();
		}

		public static string strUploadBeifenSucaiUrl()
		{
			return strBaseDomain + "upload.beifen.sucai.php?";
		}

		public static string strGetBeifenSucaiUrl()
		{
			return string.Concat(string.Concat(strBaseDomain + "get.beifen.sucai.php?", "&user_id=", 1L.ToString()), "&mac=", SecretUtil.GetMachineCode());
		}

		public static string strGetBeifenSucaiDetailUrl(string strBeifenId)
		{
			return string.Concat(string.Concat(strBaseDomain + "get.beifen.sucai.detail.php?", "&beifenid=", strBeifenId), "&mac=", SecretUtil.GetMachineCode());
		}

		public static string strGetFanWen_LikeUrl(string strSearch)
		{
			return string.Concat(string.Concat(string.Concat(strBaseDomain + "get.fanwen.like.php?", "&search=", strSearch), "&user_id=", 1L.ToString()), "&mac=", SecretUtil.GetMachineCode());
		}

		public static string strGetFanWen_Like_Rtf_Url(string strSearch)
		{
			return string.Concat(string.Concat(string.Concat(strBaseDomain + "get.fanwen.like.rtf.php?", "&search=", strSearch), "&user_id=", 1L.ToString()), "&mac=", SecretUtil.GetMachineCode());
		}

		public static string strGetFanWen_WebUrl(string strSearch)
		{
			return string.Concat(string.Concat(string.Concat(strBaseDomain + "get.fanwen.web.php?", "&search=", strSearch), "&user_id=", 1L.ToString()), "&mac=", SecretUtil.GetMachineCode());
		}

		public static string strGetFanWen_Web_Rtf_Url(string strSearch)
		{
			return string.Concat(string.Concat(string.Concat(strBaseDomain + "get.fanwen.web.rtf.php?", "&search=", strSearch), "&user_id=", 1L.ToString()), "&mac=", SecretUtil.GetMachineCode());
		}

		public static string strGetFanWen_Web_Txt_Url(string strSearch)
		{
			return string.Concat(string.Concat(string.Concat(strBaseDomain + "get.fanwen.web.v2.1.php?", "&search=", strSearch), "&user_id=", 1L.ToString()), "&mac=", SecretUtil.GetMachineCode());
		}

		public static string strGetFanWen_DetailUrl(string strFwId)
		{
			return string.Concat(string.Concat("http://www.6dgww.com/" + "fanwen.php?", "fwid=", strFwId), "&mac=", SecretUtil.GetMachineCode());
		}

		public static string strGetFanWen_ContentUrl(string strFwId)
		{
			return string.Concat(string.Concat(string.Concat(strBaseDomain + "get.fanwen.content.php?", "fwid=", strFwId), "&user_id=", 1L.ToString()), "&mac=", SecretUtil.GetMachineCode());
		}
	}
}
