using System;
using System.Collections.Generic;

namespace Local_Wps_Vsto
{
	internal class MemCacheUtil
	{
		private static List<string> _memCache = new List<string>();

		private static Dictionary<string, string> memCache = new Dictionary<string, string>();

		public static void setCache(string strKeyWord, string strCache)
		{
			if (memCache.ContainsKey(strKeyWord))
			{
				memCache.Remove(strKeyWord);
			}
			memCache.Add(strKeyWord, strCache);
		}

		public static string getCache(string strKeyWord)
		{
			string value = "";
			if (memCache.ContainsKey(strKeyWord))
			{
				memCache.TryGetValue(strKeyWord, out value);
			}
			return value;
		}

		private void test()
		{
			Dictionary<int, string> dictionary = new Dictionary<int, string>();
			dictionary.Add(1, "a");
			dictionary.Add(2, "b");
			dictionary.Add(3, "c");
			dictionary.Remove(3);
			if (!dictionary.ContainsKey(4))
			{
				dictionary.Add(4, "d");
			}
			Console.WriteLine("元素个数：{0}", dictionary.Count);
			if (dictionary.ContainsKey(1))
			{
				Console.WriteLine("key:{0},value:{1}", "1", dictionary[1]);
				Console.WriteLine(dictionary[1]);
			}
			foreach (KeyValuePair<int, string> item in dictionary)
			{
				Console.WriteLine("key={0},value={1}", item.Key, item.Value);
			}
			foreach (int key in dictionary.Keys)
			{
				Console.WriteLine("key={0}", key);
			}
			foreach (string value2 in dictionary.Values)
			{
				Console.WriteLine("value：{0}", value2);
			}
			string value = string.Empty;
			if (dictionary.TryGetValue(5, out value))
			{
				Console.WriteLine("查找结果：{0}", value);
			}
			else
			{
				Console.WriteLine("查找失败");
			}
			Console.ReadKey();
		}
	}
}
