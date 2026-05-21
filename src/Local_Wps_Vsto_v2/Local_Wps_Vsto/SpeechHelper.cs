using System.Speech.Synthesis;

namespace Local_Wps_Vsto
{
	internal class SpeechHelper
	{
		private static SpeechSynthesizer voice = new SpeechSynthesizer();

		public SpeechHelper(int rate, int volume)
		{
			voice.Rate = rate;
			voice.Volume = volume;
		}

		public static void SetRate(int rate)
		{
			voice.Rate = rate;
		}

		public static void SetVolume(int volume)
		{
			voice.Volume = volume;
		}

		public static void ReadContextAsync(string context)
		{
			voice.SpeakAsync(context);
		}

		public static void CancelSpeechAsync()
		{
			voice.SpeakAsyncCancelAll();
		}

		public static void ReadContext(string context)
		{
			voice.Speak(context);
		}

		public static void CancelSpeechc()
		{
			voice.SpeakAsyncCancelAll();
		}
	}
}
