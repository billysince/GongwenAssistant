using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Word;

namespace Local_Wps_Vsto
{
	internal class WpsHelper
	{
		public static int MAX_TITLE_LENGTH = 100;

		public static bool Is_BiaoTi_One(string strTitle)
		{
			if (strTitle.Length > MAX_TITLE_LENGTH)
			{
				return false;
			}
			bool result = false;
			try
			{
				string[] array = new string[109]
				{
					"一、", "二、", "三、", "四、", "五、", "六、", "七、", "八、", "九、", "十、",
					"十一、", "十二、", "十三、", "十四、", "十五、", "十六、", "十七、", "十八、", "十九、", "二十、",
					"二十一、", "二十二、", "二十三、", "二十四、", "二十五、", "二十六、", "二十七、", "二十八、", "二十九、", "三十、",
					"三十一、", "三十二、", "三十三、", "三十四、", "三十五、", "三十六、", "三十七、", "三十八、", "三十九、", "三十、",
					"三十一、", "三十二、", "三十三、", "三十四、", "三十五、", "三十六、", "三十七、", "三十八、", "三十九、", "四十、",
					"四十一、", "四十二、", "四十三、", "四十四、", "四十五、", "四十六、", "四十七、", "四十八、", "四十九、", "五十、",
					"五十一、", "五十二、", "五十三、", "五十四、", "五十五、", "五十六、", "五十七、", "五十八、", "五十九、", "六十、",
					"六十一、", "六十二、", "六十三、", "六十四、", "六十五、", "六十六、", "六十七、", "六十八、", "六十九、", "七十、",
					"七十一、", "七十二、", "七十三、", "七十四、", "七十五、", "七十六、", "七十七、", "七十八、", "七十九、", "八十、",
					"八十一、", "八十二、", "八十三、", "八十四、", "八十五、", "八十六、", "八十七、", "八十八、", "八十九、", "九十、",
					"九十一、", "九十二、", "九十三、", "九十四、", "九十五、", "九十六、", "九十七、", "九十八、", "九十九、"
				};
				int num = array.Length;
				for (int i = 0; i < num; i++)
				{
					if (strTitle.IndexOf(array[i]) == 0)
					{
						result = true;
						return result;
					}
				}
				return result;
			}
			catch (Exception)
			{
				return result;
			}
		}

		public static bool Is_BiaoTi_Two(string strTitle)
		{
			if (strTitle.Length > MAX_TITLE_LENGTH)
			{
				return false;
			}
			bool result = false;
			try
			{
				strTitle = strTitle.Replace("(", "（");
				strTitle = strTitle.Replace(")", "）");
				strTitle = strTitle.Replace("（", "（");
				strTitle = strTitle.Replace("）", "）");
				string[] array = new string[109]
				{
					"（一）", "（二）", "（三）", "（四）", "（五）", "（六）", "（七）", "（八）", "（九）", "（十）",
					"（十一）", "（十二）", "（十三）", "（十四）", "（十五）", "（十六）", "（十七）", "（十八）", "（十九）", "（二十）",
					"（二十一）", "（二十二）", "（二十三）", "（二十四）", "（二十五）", "（二十六）", "（二十七）", "（二十八）", "（二十九）", "（三十）",
					"（三十一）", "（三十二）", "（三十三）", "（三十四）", "（三十五）", "（三十六）", "（三十七）", "（三十八）", "（三十九）", "（三十）",
					"（三十一）", "（三十二）", "（三十三）", "（三十四）", "（三十五）", "（三十六）", "（三十七）", "（三十八）", "（三十九）", "（四十）",
					"（四十一）", "（四十二）", "（四十三）", "（四十四）", "（四十五）", "（四十六）", "（四十七）", "（四十八）", "（四十九）", "（五十）",
					"（五十一）", "（五十二）", "（五十三）", "（五十四）", "（五十五）", "（五十六）", "（五十七）", "（五十八）", "（五十九）", "（六十）",
					"（六十一）", "（六十二）", "（六十三）", "（六十四）", "（六十五）", "（六十六）", "（六十七）", "（六十八）", "（六十九）", "（七十）",
					"（七十一）", "（七十二）", "（七十三）", "（七十四）", "（七十五）", "（七十六）", "（七十七）", "（七十八）", "（七十九）", "（八十）",
					"（八十一）", "（八十二）", "（八十三）", "（八十四）", "（八十五）", "（八十六）", "（八十七）", "（八十八）", "（八十九）", "（九十）",
					"（九十一）", "（九十二）", "（九十三）", "（九十四）", "（九十五）", "（九十六）", "（九十七）", "（九十八）", "（九十九）"
				};
				int num = array.Length;
				for (int i = 0; i < num; i++)
				{
					if (strTitle.IndexOf(array[i]) == 0)
					{
						result = true;
						break;
					}
				}
				string[] array2 = new string[10] { "㈠", "㈡", "㈢", "㈣", "㈤", "㈥", "㈦", "㈧", "㈨", "㈩" };
				int num2 = array2.Length;
				for (int j = 0; j < num2; j++)
				{
					if (strTitle.IndexOf(array2[j]) == 0)
					{
						result = true;
						return result;
					}
				}
				return result;
			}
			catch (Exception)
			{
				return result;
			}
		}

		public static bool Is_BiaoTi_Three(string strTitle)
		{
			if (strTitle.Length > MAX_TITLE_LENGTH)
			{
				return false;
			}
			bool flag = false;
			try
			{
				strTitle = strTitle.Replace("(", "（");
				strTitle = strTitle.Replace(")", "）");
				strTitle = strTitle.Replace("（", "（");
				strTitle = strTitle.Replace("）", "）");
				strTitle = strTitle.ToLower();
				strTitle = strTitle.Replace("．", ".");
				string[] array = new string[109]
				{
					"一是", "二是", "三是", "四是", "五是", "六是", "七是", "八是", "九是", "十是",
					"十一是", "十二是", "十三是", "十四是", "十五是", "十六是", "十七是", "十八是", "十九是", "二十是",
					"二十一是", "二十二是", "二十三是", "二十四是", "二十五是", "二十六是", "二十七是", "二十八是", "二十九是", "三十是",
					"三十一是", "三十二是", "三十三是", "三十四是", "三十五是", "三十六是", "三十七是", "三十八是", "三十九是", "三十是",
					"三十一是", "三十二是", "三十三是", "三十四是", "三十五是", "三十六是", "三十七是", "三十八是", "三十九是", "四十是",
					"四十一是", "四十二是", "四十三是", "四十四是", "四十五是", "四十六是", "四十七是", "四十八是", "四十九是", "五十是",
					"五十一是", "五十二是", "五十三是", "五十四是", "五十五是", "五十六是", "五十七是", "五十八是", "五十九是", "六十是",
					"六十一是", "六十二是", "六十三是", "六十四是", "六十五是", "六十六是", "六十七是", "六十八是", "六十九是", "七十是",
					"七十一是", "七十二是", "七十三是", "七十四是", "七十五是", "七十六是", "七十七是", "七十八是", "七十九是", "八十是",
					"八十一是", "八十二是", "八十三是", "八十四是", "八十五是", "八十六是", "八十七是", "八十八是", "八十九是", "九十是",
					"九十一是", "九十二是", "九十三是", "九十四是", "九十五是", "九十六是", "九十七是", "九十八是", "九十九是"
				};
				int num = array.Length;
				for (int i = 0; i < num; i++)
				{
					if (strTitle.IndexOf(array[i]) == 0)
					{
						flag = true;
						break;
					}
				}
				if (!flag)
				{
					string[] array2 = new string[109]
					{
						"一要", "二要", "三要", "四要", "五要", "六要", "七要", "八要", "九要", "十要",
						"十一要", "十二要", "十三要", "十四要", "十五要", "十六要", "十七要", "十八要", "十九要", "二十要",
						"二十一要", "二十二要", "二十三要", "二十四要", "二十五要", "二十六要", "二十七要", "二十八要", "二十九要", "三十要",
						"三十一要", "三十二要", "三十三要", "三十四要", "三十五要", "三十六要", "三十七要", "三十八要", "三十九要", "三十要",
						"三十一要", "三十二要", "三十三要", "三十四要", "三十五要", "三十六要", "三十七要", "三十八要", "三十九要", "四十要",
						"四十一要", "四十二要", "四十三要", "四十四要", "四十五要", "四十六要", "四十七要", "四十八要", "四十九要", "五十要",
						"五十一要", "五十二要", "五十三要", "五十四要", "五十五要", "五十六要", "五十七要", "五十八要", "五十九要", "六十要",
						"六十一要", "六十二要", "六十三要", "六十四要", "六十五要", "六十六要", "六十七要", "六十八要", "六十九要", "七十要",
						"七十一要", "七十二要", "七十三要", "七十四要", "七十五要", "七十六要", "七十七要", "七十八要", "七十九要", "八十要",
						"八十一要", "八十二要", "八十三要", "八十四要", "八十五要", "八十六要", "八十七要", "八十八要", "八十九要", "九十要",
						"九十一要", "九十二要", "九十三要", "九十四要", "九十五要", "九十六要", "九十七要", "九十八要", "九十九要"
					};
					int num2 = array2.Length;
					for (int j = 0; j < num2; j++)
					{
						if (strTitle.IndexOf(array2[j]) == 0)
						{
							flag = true;
							break;
						}
					}
				}
				if (!flag)
				{
					string[] array3 = new string[100]
					{
						"1.", "2.", "3.", "4.", "5.", "6.", "7.", "8.", "9.", "10.",
						"11.", "12.", "13.", "14.", "15.", "16.", "17.", "18.", "19.", "20.",
						"21.", "22.", "23.", "24.", "25.", "26.", "27.", "28.", "29.", "30.",
						"31.", "32.", "33.", "34.", "35.", "36.", "37.", "38.", "39.", "40.",
						"41.", "42.", "43.", "44.", "45.", "46.", "47.", "48.", "49.", "50.",
						"51.", "52.", "53.", "54.", "55.", "56.", "57.", "58.", "59.", "60.",
						"61.", "62.", "63.", "64.", "65.", "66.", "67.", "68.", "69.", "70.",
						"71.", "72.", "73.", "74.", "75.", "76.", "77.", "78.", "79.", "80.",
						"81.", "82.", "83.", "84.", "85.", "86.", "87.", "88.", "89.", "90.",
						"91.", "92.", "93.", "94.", "95.", "96.", "97.", "98.", "99.", "100."
					};
					int num3 = array3.Length;
					for (int k = 0; k < num3; k++)
					{
						if (strTitle.IndexOf(array3[k]) == 0)
						{
							flag = true;
							break;
						}
					}
				}
				if (!flag)
				{
					string[] array4 = new string[100]
					{
						"1．", "2．", "3．", "4．", "5．", "6．", "7．", "8．", "9．", "10．",
						"11．", "12．", "13．", "14．", "15．", "16．", "17．", "18．", "19．", "20．",
						"21．", "22．", "23．", "24．", "25．", "26．", "27．", "28．", "29．", "30．",
						"31．", "32．", "33．", "34．", "35．", "36．", "37．", "38．", "39．", "40．",
						"41．", "42．", "43．", "44．", "45．", "46．", "47．", "48．", "49．", "50．",
						"51．", "52．", "53．", "54．", "55．", "56．", "57．", "58．", "59．", "60．",
						"61．", "62．", "63．", "64．", "65．", "66．", "67．", "68．", "69．", "70．",
						"71．", "72．", "73．", "74．", "75．", "76．", "77．", "78．", "79．", "80．",
						"81．", "82．", "83．", "84．", "85．", "86．", "87．", "88．", "89．", "90．",
						"91．", "92．", "93．", "94．", "95．", "96．", "97．", "98．", "99．", "100．"
					};
					int num4 = array4.Length;
					for (int l = 0; l < num4; l++)
					{
						if (strTitle.IndexOf(array4[l]) == 0)
						{
							flag = true;
							break;
						}
					}
				}
				if (!flag)
				{
					string[] array5 = new string[100]
					{
						"1、", "2、", "3、", "4、", "5、", "6、", "7、", "8、", "9、", "10、",
						"11、", "12、", "13、", "14、", "15、", "16、", "17、", "18、", "19、", "20、",
						"21、", "22、", "23、", "24、", "25、", "26、", "27、", "28、", "29、", "30、",
						"31、", "32、", "33、", "34、", "35、", "36、", "37、", "38、", "39、", "40、",
						"41、", "42、", "43、", "44、", "45、", "46、", "47、", "48、", "49、", "50、",
						"51、", "52、", "53、", "54、", "55、", "56、", "57、", "58、", "59、", "60、",
						"61、", "62、", "63、", "64、", "65、", "66、", "67、", "68、", "69、", "70、",
						"71、", "72、", "73、", "74、", "75、", "76、", "77、", "78、", "79、", "80、",
						"81、", "82、", "83、", "84、", "85、", "86、", "87、", "88、", "89、", "90、",
						"91、", "92、", "93、", "94、", "95、", "96、", "97、", "98、", "99、", "100、"
					};
					int num5 = array5.Length;
					for (int m = 0; m < num5; m++)
					{
						if (strTitle.IndexOf(array5[m]) == 0)
						{
							flag = true;
							break;
						}
					}
				}
				if (!flag)
				{
					string[] array6 = new string[12]
					{
						"Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ", "Ⅴ", "Ⅵ", "Ⅶ", "Ⅷ", "Ⅸ", "Ⅹ",
						"Ⅺ", "Ⅻ"
					};
					int num6 = array6.Length;
					for (int n = 0; n < num6; n++)
					{
						if (strTitle.IndexOf(array6[n]) == 0)
						{
							flag = true;
							break;
						}
					}
				}
				if (!flag)
				{
					string[] array7 = new string[100]
					{
						"(1)", "(2)", "(3)", "(4)", "(5)", "(6)", "(7)", "(8)", "(9)", "(10)",
						"(11)", "(12)", "(13)", "(14)", "(15)", "(16)", "(17)", "(18)", "(19)", "(20)",
						"(21)", "(22)", "(23)", "(24)", "(25)", "(26)", "(27)", "(28)", "(29)", "(30)",
						"(31)", "(32)", "(33)", "(34)", "(35)", "(36)", "(37)", "(38)", "(39)", "(40)",
						"(41)", "(42)", "(43)", "(44)", "(45)", "(46)", "(47)", "(48)", "(49)", "(50)",
						"(51)", "(52)", "(53)", "(54)", "(55)", "(56)", "(57)", "(58)", "(59)", "(60)",
						"(61)", "(62)", "(63)", "(64)", "(65)", "(66)", "(67)", "(68)", "(69)", "(70)",
						"(71)", "(72)", "(73)", "(74)", "(75)", "(76)", "(77)", "(78)", "(79)", "(80)",
						"(81)", "(82)", "(83)", "(84)", "(85)", "(86)", "(87)", "(88)", "(89)", "(90)",
						"(91)", "(92)", "(93)", "(94)", "(95)", "(96)", "(97)", "(98)", "(99)", "(100）"
					};
					int num7 = array7.Length;
					for (int num8 = 0; num8 < num7; num8++)
					{
						if (strTitle.IndexOf(array7[num8]) == 0)
						{
							flag = true;
							break;
						}
					}
				}
				if (!flag)
				{
					string[] array8 = new string[100]
					{
						"（1）", "（2）", "（3）", "（4）", "（5）", "（6）", "（7）", "（8）", "（9）", "（10）",
						"（11）", "（12）", "（13）", "（14）", "（15）", "（16）", "（17）", "（18）", "（19）", "（20）",
						"（21）", "（22）", "（23）", "（24）", "（25）", "（26）", "（27）", "（28）", "（29）", "（30）",
						"（31）", "（32）", "（33）", "（34）", "（35）", "（36）", "（37）", "（38）", "（39）", "（40）",
						"（41）", "（42）", "（43）", "（44）", "（45）", "（46）", "（47）", "（48）", "（49）", "（50）",
						"（51）", "（52）", "（53）", "（54）", "（55）", "（56）", "（57）", "（58）", "（59）", "（60）",
						"（61）", "（62）", "（63）", "（64）", "（65）", "（66）", "（67）", "（68）", "（69）", "（70）",
						"（71）", "（72）", "（73）", "（74）", "（75）", "（76）", "（77）", "（78）", "（79）", "（80）",
						"（81）", "（82）", "（83）", "（84）", "（85）", "（86）", "（87）", "（88）", "（89）", "（90）",
						"（91）", "（92）", "（93）", "（94）", "（95）", "（96）", "（97）", "（98）", "（99）", "（100）"
					};
					int num9 = array8.Length;
					for (int num10 = 0; num10 < num9; num10++)
					{
						if (strTitle.IndexOf(array8[num10]) == 0)
						{
							flag = true;
							break;
						}
					}
				}
				if (!flag)
				{
					string[] array9 = new string[10] { "①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩" };
					int num11 = array9.Length;
					for (int num12 = 0; num12 < num11; num12++)
					{
						if (strTitle.IndexOf(array9[num12]) == 0)
						{
							flag = true;
							break;
						}
					}
				}
				if (!flag)
				{
					string[] array10 = new string[20]
					{
						"⑴", "⑵", "⑶", "⑷", "⑸", "⑹", "⑺", "⑻", "⑼", "⑽",
						"⑾", "⑿", "⒀", "⒁", "⒂", "⒃", "⒄", "⒅", "⒆", "⒇"
					};
					int num13 = array10.Length;
					for (int num14 = 0; num14 < num13; num14++)
					{
						if (strTitle.IndexOf(array10[num14]) == 0)
						{
							flag = true;
							break;
						}
					}
				}
				if (!flag)
				{
					string[] array11 = new string[10] { "⒈", "⒉", "⒊", "⒋", "⒌", "⒍", "⒎", "⒏", "⒐", "⒑" };
					int num15 = array11.Length;
					for (int num16 = 0; num16 < num15; num16++)
					{
						if (strTitle.IndexOf(array11[num16]) == 0)
						{
							flag = true;
							break;
						}
					}
				}
				if (!flag)
				{
					string[] array12 = new string[26]
					{
						"A.", "B.", "C.", "D.", "E.", "F.", "G.", "H.", "I.", "J.",
						"K.", "L.", "M.", "N.", "O.", "P.", "Q.", "R.", "S.", "T.",
						"U.", "V.", "W.", "X.", "Y.", "Z."
					};
					int num17 = array12.Length;
					for (int num18 = 0; num18 < num17; num18++)
					{
						if (strTitle.IndexOf(array12[num18]) == 0)
						{
							flag = true;
							break;
						}
					}
				}
				if (!flag)
				{
					string[] array13 = new string[26]
					{
						"A．", "B．", "C．", "D．", "E．", "F．", "G．", "H．", "I．", "J．",
						"K．", "L．", "M．", "N．", "O．", "P．", "Q．", "R．", "S．", "T．",
						"U．", "V．", "W．", "X．", "Y．", "Z．"
					};
					int num19 = array13.Length;
					for (int num20 = 0; num20 < num19; num20++)
					{
						if (strTitle.IndexOf(array13[num20]) == 0)
						{
							flag = true;
							break;
						}
					}
				}
				if (!flag)
				{
					string[] array14 = new string[26]
					{
						"A、", "B、", "C、", "D、", "E、", "F、", "G、", "H、", "I、", "J、",
						"K、", "L、", "M、", "N、", "O、", "P、", "Q、", "R、", "S、", "T、",
						"U、", "V、", "W、", "X、", "Y、", "Z、"
					};
					int num21 = array14.Length;
					for (int num22 = 0; num22 < num21; num22++)
					{
						if (strTitle.IndexOf(array14[num22]) == 0)
						{
							flag = true;
							return flag;
						}
					}
					return flag;
				}
				return flag;
			}
			catch (Exception)
			{
				return flag;
			}
		}

		public static bool Is_BiaoTi_TopTitle(string strTitle, int nowPos)
		{
			bool result = false;
			try
			{
				if (nowPos <= 3)
				{
					if (strTitle.Length < 30)
					{
						string[] array = new string[20]
						{
							"讲话", "方案", "办法", "情况", "总结", "报告", "汇报", "通知", "函", "请示",
							"意见", "关于", "印发", "计划", "心得", "体会", "交流", "发言", "范文", "申请"
						};
						for (int i = 0; i < array.Length; i++)
						{
							if (strTitle.IndexOf(array[i]) > 0)
							{
								return true;
							}
						}
						return result;
					}
					return result;
				}
				return result;
			}
			catch (Exception)
			{
				return result;
			}
		}

		public static bool Is_BiaoTi_Time(string strTitle, int nowPos)
		{
			bool result = false;
			try
			{
				if (nowPos == 1)
				{
					if (strTitle.Length < 20)
					{
						string[] array = new string[7] { "发", "文", "【", "[", "〔", "〖", "号" };
						for (int i = 0; i < array.Length; i++)
						{
							if (strTitle.IndexOf(array[i]) > 0)
							{
								return true;
							}
						}
					}
					if (strTitle.IndexOf("年") > 0)
					{
						if (strTitle.IndexOf("月") > 0)
						{
							if (strTitle.IndexOf("日") > 0)
							{
								return true;
							}
							return result;
						}
						return result;
					}
					return result;
				}
				return result;
			}
			catch (Exception)
			{
				return result;
			}
		}

		public static bool Is_BiaoTi_People(string strTitle, int nowPos)
		{
			bool result = false;
			try
			{
				if (nowPos <= 6)
				{
					if (strTitle.Length < 80)
					{
						if (strTitle.IndexOf("各") >= 0 || strTitle.IndexOf("们") > 0)
						{
							return true;
						}
						return result;
					}
					return result;
				}
				return result;
			}
			catch (Exception)
			{
				return result;
			}
		}

		public static bool Is_FuJian(string strTitle)
		{
			bool result = false;
			if (strTitle.IndexOf("附件：") >= 0 || strTitle.IndexOf("附件:") >= 0)
			{
				return true;
			}
			return result;
		}

		public static string Pre_Exe_Article(string strArticle)
		{
			try
			{
				strArticle = strArticle.Replace("０", "0");
				strArticle = strArticle.Replace("１", "1");
				strArticle = strArticle.Replace("２", "2");
				strArticle = strArticle.Replace("３", "3");
				strArticle = strArticle.Replace("４", "4");
				strArticle = strArticle.Replace("５", "5");
				strArticle = strArticle.Replace("６", "6");
				strArticle = strArticle.Replace("７", "7");
				strArticle = strArticle.Replace("８", "8");
				strArticle = strArticle.Replace("９", "9");
				strArticle = strArticle.Replace("．", ".");
				strArticle = strArticle.Replace("。", "。");
				strArticle = strArticle.Replace("，", "，");
				strArticle = strArticle.Replace(",", "，");
				strArticle = strArticle.Replace(";", "；");
				strArticle = strArticle.Replace("、", "、");
				strArticle = strArticle.Replace("?", "？");
				strArticle = strArticle.Replace(":", "：");
				strArticle = strArticle.Replace("!", "！");
				string[] array = new string[109]
				{
					"（一）", "（二）", "（三）", "（四）", "（五）", "（六）", "（七）", "（八）", "（九）", "（十）",
					"（十一）", "（十二）", "（十三）", "（十四）", "（十五）", "（十六）", "（十七）", "（十八）", "（十九）", "（二十）",
					"（二十一）", "（二十二）", "（二十三）", "（二十四）", "（二十五）", "（二十六）", "（二十七）", "（二十八）", "（二十九）", "（三十）",
					"（三十一）", "（三十二）", "（三十三）", "（三十四）", "（三十五）", "（三十六）", "（三十七）", "（三十八）", "（三十九）", "（三十）",
					"（三十一）", "（三十二）", "（三十三）", "（三十四）", "（三十五）", "（三十六）", "（三十七）", "（三十八）", "（三十九）", "（四十）",
					"（四十一）", "（四十二）", "（四十三）", "（四十四）", "（四十五）", "（四十六）", "（四十七）", "（四十八）", "（四十九）", "（五十）",
					"（五十一）", "（五十二）", "（五十三）", "（五十四）", "（五十五）", "（五十六）", "（五十七）", "（五十八）", "（五十九）", "（六十）",
					"（六十一）", "（六十二）", "（六十三）", "（六十四）", "（六十五）", "（六十六）", "（六十七）", "（六十八）", "（六十九）", "（七十）",
					"（七十一）", "（七十二）", "（七十三）", "（七十四）", "（七十五）", "（七十六）", "（七十七）", "（七十八）", "（七十九）", "（八十）",
					"（八十一）", "（八十二）", "（八十三）", "（八十四）", "（八十五）", "（八十六）", "（八十七）", "（八十八）", "（八十九）", "（九十）",
					"（九十一）", "（九十二）", "（九十三）", "（九十四）", "（九十五）", "（九十六）", "（九十七）", "（九十八）", "（九十九）"
				};
				int num = array.Length;
				for (int i = 0; i < num; i++)
				{
					strArticle = strArticle.Replace(array[i], "\r\n" + array[i]);
				}
				return strArticle;
			}
			catch (Exception)
			{
				return strArticle;
			}
		}

		public static void ExeWordFile_FromTxt(string strArticleTxt, string strWordFileName)
		{
			try
			{
				WpsUtil wpsUtil = new WpsUtil();
				wpsUtil.StartDocument();
				string[] source = Pre_Exe_Article(strArticleTxt).Split('\r', '\n');
				source = source.Where((string s) => !string.IsNullOrEmpty(s)).ToArray();
				for (int i = 0; i < source.Length; i++)
				{
					string text = source[i];
					text = text.Trim();
					text = text.Replace("\t", "");
					if (text.IndexOf("。") > 0)
					{
						char[] separator = new char[1] { '。' };
						string[] source2 = text.Split(separator);
						source2 = source2.Where((string s) => !string.IsNullOrEmpty(s)).ToArray();
						wpsUtil.Start_P();
						for (int j = 0; j < source2.Length; j++)
						{
							string text2 = source2[j];
							text2 = text2.Trim();
							if (text2.Length > 0)
							{
								if (Is_BiaoTi_One(text2))
								{
									wpsUtil.Add_One_DuanNei(text2);
								}
								else if (Is_BiaoTi_Two(text2))
								{
									wpsUtil.Add_Two_DuanNei(text2);
								}
								else if (Is_BiaoTi_Three(text2))
								{
									wpsUtil.Add_Three_DuanNei(text2);
								}
								else
								{
									wpsUtil.Add_NeiRong_DuanNei(text2);
								}
								wpsUtil.Add_NeiRong_DuanNei("。");
							}
						}
						wpsUtil.Close_P();
					}
					else if (text.IndexOf("：") > 0)
					{
						if (Is_BiaoTi_People(text, i))
						{
							wpsUtil.Add_People(text);
						}
						else if (text.IndexOf("附件") == 0)
						{
							wpsUtil.Start_P();
							wpsUtil.Add_NeiRong_DuanNei("");
							wpsUtil.Close_P();
							wpsUtil.Start_P();
							wpsUtil.Add_NeiRong_DuanNei(text);
							wpsUtil.Close_P();
						}
					}
					else if (Is_BiaoTi_TopTitle(text, i))
					{
						wpsUtil.Add_Top(text);
						if (source.Length >= 1 && !Is_BiaoTi_Time(source[1], 1))
						{
							wpsUtil.Add_Time("");
						}
					}
					else if (Is_BiaoTi_Time(text, i))
					{
						wpsUtil.Add_Time(text);
						wpsUtil.Add_Time("");
					}
					else if (Is_BiaoTi_One(text))
					{
						wpsUtil.Add_One(text);
					}
					else if (Is_BiaoTi_Two(text))
					{
						wpsUtil.Add_Two(text);
					}
					else if (Is_BiaoTi_Three(text))
					{
						wpsUtil.Add_Three(text);
					}
					else if (i == source.Length - 1 && text.Length < 11 && text.IndexOf("年") > 0 && text.IndexOf("月") > 0 && text.IndexOf("日") > 0)
					{
						wpsUtil.Start_P();
						wpsUtil.Add_NeiRong_DuanNei("");
						wpsUtil.Close_P();
						wpsUtil.Start_P();
						wpsUtil.Add_NeiRong_DuanNei("");
						wpsUtil.Close_P();
						wpsUtil.Add_Last_Time(text);
					}
					else
					{
						wpsUtil.Start_P();
						wpsUtil.Add_NeiRong_DuanNei(text);
						wpsUtil.Close_P();
					}
				}
				wpsUtil.CloseDocument();
				string zipFilePath = CommonConfig.strBaseFolder + "\\template\\model.zip";
				string text3 = CommonConfig.strBaseFolder + "\\cache\\model";
				ZipHelper.UnZipFile(zipFilePath, text3);
				File.WriteAllText(text3 + "\\word\\document.xml", wpsUtil.getContent());
				ZipHelper.PackFiles(strWordFileName, text3);
				ZipHelper.DeleteDir(text3);
			}
			catch (Exception)
			{
			}
		}

		public static void ExeNewWord_FromRtfString(string strRtf, string strWordFileName)
		{
			try
			{
				File.Copy(CommonConfig.strBaseFolder + "\\template\\A4_1.docx", strWordFileName, true);
				string text = string.Concat(str2: ((long)new Random().Next(10000000, 99999999)).ToString(), str0: CommonConfig.strBaseFolder, str1: "\\cache\\temp_", str3: ".rtf");
				File.WriteAllText(text, strRtf);
				Word.Application application = (Word.Application)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("000209FF-0000-0000-C000-000000000046")));
				Document document = (Document)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("00020906-0000-0000-C000-000000000046")));
				Word.Application application2 = (Word.Application)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("000209FF-0000-0000-C000-000000000046")));
				Document document2 = (Document)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("00020906-0000-0000-C000-000000000046")));
				object FileName = strWordFileName;
				object FileName2 = text;
				object ReadOnly = false;
				object Visible = true;
				object ConfirmConversions = Missing.Value;
				document2 = application2.Documents.Open(ref FileName2, ref ConfirmConversions, ref ReadOnly, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref Visible, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions);
				ReadOnly = false;
				document = application.Documents.Open(ref FileName, ref ConfirmConversions, ref ReadOnly, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref Visible, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions);
				document2.ActiveWindow.Selection.WholeStory();
				document2.ActiveWindow.Selection.Copy();
				document.ActiveWindow.Selection.WholeStory();
				document.ActiveWindow.Selection.Paste();
				document2.Close(ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions);
				document2 = null;
				application2.Quit(ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions);
				application2 = null;
				foreach (Paragraph paragraph in document.Paragraphs)
				{
					if (paragraph.FirstLineIndent > 0f)
					{
						paragraph.CharacterUnitFirstLineIndent = 2f;
					}
					paragraph.Format.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
					paragraph.Format.LineSpacing = 28.95f;
					paragraph.Format.WidowControl = 0;
					if (paragraph.Alignment == WdParagraphAlignment.wdAlignParagraphCenter)
					{
						paragraph.CharacterUnitFirstLineIndent = 0f;
					}
					else if (paragraph.Alignment == WdParagraphAlignment.wdAlignParagraphRight)
					{
						paragraph.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
						paragraph.CharacterUnitFirstLineIndent = 0f;
					}
					else
					{
						paragraph.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
					}
				}
				document.Save();
				document.Close(ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions);
				document = null;
				application.Quit(ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions);
				application = null;
			}
			catch (Exception)
			{
			}
		}

		public static void ExeNewWord_FromRtfString_New(string strRtf, string strWordFileName)
		{
			try
			{
				object Template = CommonConfig.strBaseFolder + "\\template\\A4_1.docx";
				Word.Application application = (Word.Application)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("000209FF-0000-0000-C000-000000000046")));
				Document document = (Document)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("00020906-0000-0000-C000-000000000046")));
				Missing value = Missing.Value;
				Clipboard.SetData(DataFormats.Rtf, strRtf);
				Documents documents = application.Documents;
				object NewTemplate = Type.Missing;
				object DocumentType = Type.Missing;
				object Visible = Type.Missing;
				documents.Add(ref Template, ref NewTemplate, ref DocumentType, ref Visible);
				application.ActiveWindow.Visible = true;
				application.ActiveWindow.Selection.PasteAndFormat(WdRecoveryType.wdFormatOriginalFormatting);
			}
			catch (Exception)
			{
			}
		}

		public static void ExeNewWord_FromOldWord(string strOldFileName, string strWordFileName)
		{
			try
			{
				object Template = CommonConfig.strBaseFolder + "\\template\\A4_1.docx";
				Word.Application application = (Word.Application)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("000209FF-0000-0000-C000-000000000046")));
				Document document = (Document)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("00020906-0000-0000-C000-000000000046")));
				Word.Application application2 = (Word.Application)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("000209FF-0000-0000-C000-000000000046")));
				Document document2 = (Document)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("00020906-0000-0000-C000-000000000046")));
				object FileName = strWordFileName;
				object FileName2 = strOldFileName;
				object ReadOnly = false;
				object Visible = true;
				object ConfirmConversions = Missing.Value;
				document2 = application2.Documents.Open(ref FileName2, ref ConfirmConversions, ref ReadOnly, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref Visible, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions);
				ReadOnly = false;
				Documents documents = application.Documents;
				object NewTemplate = Type.Missing;
				object DocumentType = Type.Missing;
				object Visible2 = Type.Missing;
				document = documents.Add(ref Template, ref NewTemplate, ref DocumentType, ref Visible2);
				document2.ActiveWindow.Selection.WholeStory();
				document2.ActiveWindow.Selection.Copy();
				document.ActiveWindow.Selection.WholeStory();
				document.ActiveWindow.Selection.Paste();
				document2.Close(ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions);
				document2 = null;
				application2.Quit(ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions);
				application2 = null;
				foreach (Paragraph paragraph in document.Paragraphs)
				{
					if (paragraph.FirstLineIndent > 0f)
					{
						paragraph.CharacterUnitFirstLineIndent = 2f;
					}
					paragraph.Format.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
					paragraph.Format.LineSpacing = 28.95f;
					paragraph.Format.WidowControl = 0;
					if (paragraph.Alignment == WdParagraphAlignment.wdAlignParagraphCenter)
					{
						paragraph.CharacterUnitFirstLineIndent = 0f;
					}
					else if (paragraph.Alignment == WdParagraphAlignment.wdAlignParagraphRight)
					{
						paragraph.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
						paragraph.CharacterUnitFirstLineIndent = 0f;
					}
					else
					{
						paragraph.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
					}
				}
				document.SaveAs(ref FileName, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions);
				document.Close(ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions);
				document = null;
				application.Quit(ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions);
				application = null;
			}
			catch (Exception)
			{
			}
		}

		private static void resetParagraph(Paragraph paragraph)
		{
			try
			{
				paragraph.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
				paragraph.LineSpacing = 28.95f;
				paragraph.WidowControl = 0;
				paragraph.LeftIndent = 0f;
				paragraph.RightIndent = 0f;
				paragraph.Range.Font.Size = 16f;
				paragraph.Range.Underline = WdUnderline.wdUnderlineNone;
				paragraph.Range.HighlightColorIndex = WdColorIndex.wdAuto;
				paragraph.Range.Font.Bold = 0;
				paragraph.Range.Font.Italic = 0;
				paragraph.LineUnitBefore = 0f;
				paragraph.LineUnitAfter = 0f;
				paragraph.DisableLineHeightGrid = 0;
				paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
				paragraph.Range.Font.Color = WdColor.wdColorBlack;
				paragraph.CharacterUnitFirstLineIndent = 2f;
			}
			catch (Exception)
			{
			}
		}

		public static void setBiaoTi_TopTitle(Paragraph paragraph)
		{
			resetParagraph(paragraph);
			paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			paragraph.Range.Font.Size = 22f;
			paragraph.Range.Font.Name = "方正小标宋简体";
			paragraph.Range.Font.Bold = 0;
			paragraph.CharacterUnitFirstLineIndent = 0f;
			paragraph.FirstLineIndent = 0f;
		}

		public static void setBiaoTi_Content(Paragraph paragraph)
		{
			resetParagraph(paragraph);
			paragraph.CharacterUnitFirstLineIndent = 2f;
			paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
			paragraph.Range.Font.Size = 16f;
			paragraph.Range.Font.Name = "仿宋_GB2312";
			paragraph.Range.Font.Bold = 0;
		}

		public static void setBiaoTi_People(Paragraph paragraph)
		{
			resetParagraph(paragraph);
			paragraph.CharacterUnitFirstLineIndent = 0f;
			paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
			paragraph.Range.Font.Size = 16f;
			paragraph.Range.Font.Name = "仿宋_GB2312";
			paragraph.Range.Font.Bold = 0;
			paragraph.CharacterUnitFirstLineIndent = 0f;
			paragraph.FirstLineIndent = 0f;
		}

		public static void setBiaoTi_TopAuthor(Paragraph paragraph)
		{
			resetParagraph(paragraph);
			paragraph.CharacterUnitFirstLineIndent = 0f;
			paragraph.FirstLineIndent = 0f;
			paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			paragraph.Range.Font.Size = 16f;
			paragraph.Range.Font.Name = "楷体_GB2312";
			paragraph.Range.Font.Bold = 0;
		}

		public static void setBiaoTi_LastAuthor(Paragraph paragraph)
		{
			resetParagraph(paragraph);
			paragraph.CharacterUnitFirstLineIndent = 0f;
			paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
			paragraph.Range.Font.Size = 16f;
			paragraph.Format.CharacterUnitRightIndent = 4f;
			paragraph.Range.Font.Name = "仿宋_GB2312";
			paragraph.Range.Font.Bold = 0;
		}

		public static void setBiaoTi_One_Paragraph(Paragraph paragraph)
		{
			resetParagraph(paragraph);
			paragraph.CharacterUnitFirstLineIndent = 2f;
			paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
			paragraph.Range.Font.Size = 16f;
			paragraph.Range.Font.Name = "黑体";
			paragraph.Range.Font.Bold = 0;
		}

		public static void setBiaoTi_Two_Paragraph(Paragraph paragraph)
		{
			resetParagraph(paragraph);
			paragraph.CharacterUnitFirstLineIndent = 2f;
			paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
			paragraph.Range.Font.Size = 16f;
			paragraph.Range.Font.Name = "楷体_GB2312";
			paragraph.Range.Font.Bold = 0;
		}

		public static void setBiaoTi_Three_Paragraph(Paragraph paragraph)
		{
			resetParagraph(paragraph);
			paragraph.CharacterUnitFirstLineIndent = 2f;
			paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
			paragraph.Range.Font.Size = 16f;
			paragraph.Range.Bold = 1;
			paragraph.Range.Font.Name = "仿宋_GB2312";
		}

		private static void resetRange(Range range)
		{
			try
			{
				range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
				range.ParagraphFormat.LineSpacing = 28.95f;
				range.ParagraphFormat.WidowControl = 0;
				range.ParagraphFormat.CharacterUnitFirstLineIndent = 2f;
				range.ParagraphFormat.LeftIndent = 0f;
				range.ParagraphFormat.RightIndent = 0f;
				range.Font.Size = 16f;
				range.Underline = WdUnderline.wdUnderlineNone;
				range.HighlightColorIndex = WdColorIndex.wdAuto;
				range.Font.Bold = 0;
				range.Font.Italic = 0;
				range.ParagraphFormat.LineUnitBefore = 0f;
				range.ParagraphFormat.LineUnitAfter = 0f;
				range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
				range.Font.Color = WdColor.wdColorBlack;
			}
			catch (Exception)
			{
			}
		}

		public static void setBiaoTi_One_JuZhi(Range range)
		{
			resetRange(range);
			range.Font.Name = "黑体";
			range.Bold = 0;
		}

		public static void setBiaoTi_Two_JuZhi(Range range)
		{
			resetRange(range);
			range.Font.Name = "楷体_GB2312";
			range.Bold = 0;
		}

		public static void setBiaoTi_Three_JuZhi(Range range)
		{
			resetRange(range);
			range.Font.Name = "仿宋_GB2312";
			range.Bold = 1;
		}

		public static void setBiaoTi_Content_JuZhi(Range range)
		{
			resetRange(range);
			range.Font.Name = "仿宋_GB2312";
			range.Bold = 0;
		}

		public static void setBiaoTi_TopTitle_JuZhi(Range range)
		{
			resetRange(range);
			range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			range.Font.Size = 22f;
			range.Font.Name = "方正小标宋简体";
			range.Font.Bold = 0;
			range.ParagraphFormat.CharacterUnitFirstLineIndent = 0f;
			range.ParagraphFormat.FirstLineIndent = 0f;
		}

		public static void setBiaoTi_TopAuthor_JuZhi(Range range)
		{
			resetRange(range);
			range.ParagraphFormat.CharacterUnitFirstLineIndent = 0f;
			range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			range.Font.Size = 16f;
			range.Font.Name = "楷体_GB2312";
			range.Font.Bold = 0;
			range.ParagraphFormat.CharacterUnitFirstLineIndent = 0f;
			range.ParagraphFormat.FirstLineIndent = 0f;
		}

		public static void setBiaoTi_People_JuZhi(Range range)
		{
			resetRange(range);
			range.ParagraphFormat.CharacterUnitFirstLineIndent = 0f;
			range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
			range.Font.Size = 16f;
			range.Font.Name = "仿宋_GB2312";
			range.Font.Bold = 0;
			range.ParagraphFormat.CharacterUnitFirstLineIndent = 0f;
			range.ParagraphFormat.FirstLineIndent = 0f;
		}

		public static void setBiaoTi_FuJian_JuZhi(Range range)
		{
			resetRange(range);
			range.ParagraphFormat.CharacterUnitFirstLineIndent = 2f;
			range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
			range.Font.Size = 16f;
			range.Font.Name = "仿宋_GB2312";
			range.Font.Bold = 0;
		}

		public static void setBiaoTi_LastAuthor_JuZhi(Range range)
		{
			resetRange(range);
			range.ParagraphFormat.CharacterUnitFirstLineIndent = 0f;
			range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
			range.Font.Size = 16f;
			range.ParagraphFormat.CharacterUnitRightIndent = 4f;
			range.Font.Name = "仿宋_GB2312";
			range.Font.Bold = 0;
		}
	}
}
