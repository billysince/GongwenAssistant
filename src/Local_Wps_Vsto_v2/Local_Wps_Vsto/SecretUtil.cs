using System;
using System.Globalization;

namespace Local_Wps_Vsto
{
	internal class SecretUtil
	{
		private static char[] Source_String = new char[33]
		{
			'F', 'D', 'G', '9', 'M', 'W', 'J', '4', '2', 'T',
			'U', 'X', '3', 'K', 'L', 'E', 'P', 'V', '6', '7',
			'Y', 'Z', 'B', '1', 'N', '8', 'H', 'Q', 'A', '5',
			'C', 'R', 'S'
		};

		public static long longPianYi = 1212121212L;

		public static string GetMachineCode()
		{
			return encodeToSecret(GetMachineCode_Long());
		}

		public static long GetMachineCode_Long()
		{
			string machineCodeString = MachineCode.GetMachineCodeString();
			string s = machineCodeString.Substring(0, machineCodeString.Length / 2);
			string s2 = machineCodeString.Substring(machineCodeString.Length / 2);
			long num = long.Parse(s, NumberStyles.HexNumber);
			long num2 = long.Parse(s2, NumberStyles.HexNumber);
			return num + num2;
		}

		public static string encodeToSecret(long longNo)
		{
			string text = "";
			while (longNo > 0)
			{
				long num = longNo % 33;
				longNo = (longNo - num) / 33;
				text = Source_String[num] + text;
			}
			if (text.Length < 4)
			{
				text = text.PadRight(4, '0');
			}
			return text;
		}

		public static long decodeFromSecret(string code)
		{
			code = code.ToUpper();
			code = code.Replace("0", "");
			long num = code.Length;
			code = Reverse(code);
			long num2 = 0L;
			for (int i = 0; i < num; i++)
			{
				char charCode = code[i];
				num2 += getPosChar(charCode) * long.Parse(Math.Pow(33.0, i).ToString());
			}
			return num2;
		}

		private static int getPosChar(char charCode)
		{
			int num = 0;
			for (num = 0; num < Source_String.Length; num++)
			{
				if (charCode == Source_String[num])
				{
					return num;
				}
			}
			return num;
		}

		private static string Reverse(string str)
		{
			char[] array = str.ToCharArray();
			Array.Reverse(array);
			return new string(array);
		}

		public static string getNeedActiviteCode()
		{
			return encodeToSecret(GetMachineCode_Long() + longPianYi);
		}
	}
}
