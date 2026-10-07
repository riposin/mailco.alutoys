using System;
using System.Runtime.Remoting.Messaging;
using System.Threading.Tasks;

namespace mailco.alutoys
{
	public static partial class Program
	{
		static async Task Main(string[] args)
		{
			bool continueExecution = false;
			string message = "";

			LogMessage("Aplicación iniciada.");
			Console.WriteLine("Aplicación de copia de seguridad de correos, presione cualquier tecla para iniciar.");
			Console.ReadKey();
			continueExecution = DoOutlookProcess();
			if(continueExecution)
			{
				message = "Proceso de copia de seguridad iniciado.";
				LogMessage(message);
				Console.WriteLine(message);
				await DoBackupProcess();
			}
			else
			{
				message = "Proceso cancelado.";
				LogMessage(message);
				Console.WriteLine(message);
			}
			Console.ReadKey();
		}
	}
}
