using System;
using System.IO;

namespace mailco.alutoys
{
	public static partial class Program
	{
		public static void LogMessage(string message)
		{
			string logFilePath = "log.txt";
			using (StreamWriter logFileWriter = new StreamWriter(logFilePath, append: true))
			{
				logFileWriter.WriteLine($"{DateTime.Now}: {message}");
				logFileWriter.Flush();
				logFileWriter.Close();
			}
		}
		private struct Result
		{
			public bool isOk;
			public int code;
			public string message;
			public string exception;
		}
	}
}