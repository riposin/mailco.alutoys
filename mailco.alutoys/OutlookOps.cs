using Microsoft.Office.Interop.Outlook;
using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;

namespace mailco.alutoys
{
	public static partial class Program
	{
		private static void DoOutlookProcess()
		{
			string message = "";
			bool endProcess = false;
			Result res = new Result();

			while (!endProcess)
			{
				if (IsOutlookRunning())
				{
					message = "Outlook está en ejecución. Se procede a cerrar.";
					LogMessage(message);
					Console.WriteLine(message);
					res = CloseOutlook();

					LogMessage(res.message);
					Console.WriteLine(res.message);
					endProcess = res.isOk;
					if (!res.isOk)
					{
						message = "\nPresione la tecla X para cancelar la operación.";
						message += "\nSi desea reintentar, por favor cierre Outlook manualmente y presione cualquier otra tecla para continuar.";
						LogMessage("Fallo al cerrar Outlook: " + res.exception);
						Console.WriteLine(message);
						endProcess = Console.ReadKey().KeyChar.ToString().ToLower() == "x";
						LogMessage(endProcess ? "Usuario solicito cancelar." : "Usuario solicitó continuar.");
					}
					Console.Clear();
				}
				else
				{
					message = "No está en ejecución.";
					LogMessage(message);
					Console.WriteLine(message);
					endProcess = true;
				}
			}
			message = "Proceso completado, presione cualquier tecla para finalizar...";
			LogMessage(message);
			Console.WriteLine(message);
			Console.ReadKey();
		}
		private static bool IsOutlookRunning()
		{
			return Process.GetProcessesByName("OUTLOOK").Count() > 0;
		}

		private static Result CloseOutlook()
		{
			Result res = new Result();

			res.isOk = false;
			res.code = -1;
			res.message = "";
			res.exception = "";

			try
			{
				var app = Marshal.GetActiveObject("Outlook.Application") as _Application;
				if (app != null)
				{
					app.Quit();
					Marshal.ReleaseComObject(app);
					res.message = "Se ha cerrado correctamente.";
					res.isOk = true;
					res.code = 0;
				}
				else
				{
					res.message = "No se pudo obtener la instancia.";
				}
			}
			catch (COMException ex)
			{
				res.message = "No se pudo cerrar.";
				res.exception = ex.Message;
				res.code = -2;
			}

			return res;
		}
	}
}