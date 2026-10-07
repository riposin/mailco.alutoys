using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;

namespace mailco.alutoys
{
	internal class FileCopier
	{
		public static async Task CopyWithProgressAsync(string sourcePath, string destinationPath, Action<int> progressCallback, int bufferSize = 8192)
		{
			// 8192 = 8 KB buffer // 1024 * 1024 = 1MiB
			if (!File.Exists(sourcePath))
				throw new ArgumentException("Source file does not exist.");

			using (var sourceStream = new FileStream(
				sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize, useAsync: true))
			using (var destStream = new FileStream(
				destinationPath, FileMode.Create, FileAccess.Write, FileShare.None, bufferSize, useAsync: true))
			{
				long totalBytesCopied = 0;
				long fileSize = sourceStream.Length;
				byte[] buffer = new byte[bufferSize];
				int bytesRead;

				while ((bytesRead = await sourceStream.ReadAsync(buffer, 0, buffer.Length)) > 0)
				{
					await destStream.WriteAsync(buffer, 0, bytesRead);
					totalBytesCopied += bytesRead;

					// Calculate and report progress percentage
					int percentage = (int)((double)totalBytesCopied / fileSize * 100);
					progressCallback?.Invoke(percentage);
				}
			}
		}
	}
}
