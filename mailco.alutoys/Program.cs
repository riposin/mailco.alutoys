using System;
using System.Threading.Tasks;

namespace mailco.alutoys
{
	public static partial class Program
	{
		static async Task Main(string[] args)
		{
			LogMessage("Aplicación iniciada.");
			Console.WriteLine("Aplicación de copia de seguridad de correos, presione cualquier tecla para iniciar.");
			Console.ReadKey();
			DoOutlookProcess();
		}
	}
}
