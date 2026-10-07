using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace mailco.alutoys
{
	public static partial class Program
	{
		static async Task DoBackupProcess()
		{
			Console.WriteLine("Loremipsum");
			Console.ReadKey();
			Console.WriteLine("Starting async operation...");
			await PerformAsyncOperation(DoCallback);
			Console.WriteLine("Async operation completed.");
			Console.ReadKey();
		}
		static async Task PerformAsyncOperation(Func<Task> callback)
		{
			// Simulate async work
			await Task.Delay(2000);

			// Execute the callback
			await callback();
		}
		static async Task DoCallback()
		{
			Console.WriteLine("Async callback executed!");
		}
	}
}