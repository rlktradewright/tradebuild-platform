using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

using TradeWright.Utilities.DataStorage;
using TradeWright.Utilities.Logging;

using TWUtilities40;

namespace FileAutoReader
{

	internal sealed class Program
    {
		[STAThread]
		public static void Main(string[] args)
        {
			(new MainClass()).run(args);
        }

	}

	internal sealed class MainClass : ITimerExpiryListener
	{

		private static TradeWright.Utilities.Logging.FormattingLogger mLogger;

		private static FileSystemWatcher mFileSystemWatcher;

		private static readonly BlockingCollection<string> mFilenameQueue = new BlockingCollection<string>();

		private static CancellationTokenSource mCancellationTokenSource = new CancellationTokenSource();

		private static Task mDispatcher;

		private static readonly Dictionary<string, DateTime> fileNotifications = new Dictionary<string, DateTime>();

		private static readonly TimeSpan OneSecond = new TimeSpan(0, 0, 0, 0, 500);

		private static _TWUtilities TW;
		private static TWUtilities40._Console mCon;
		private static TWUtilities40._IntervalTimer mTimer;

		internal void run(string[] args)
		{
			try
			{
				Logging.DefaultLogLevel = LogLevel.Normal;

				TW = new TWUtilities();
				mCon = TW.GetConsole();
				TW.ApplicationGroupName = "TradeWright";
				TW.ApplicationName = nameof(FileAutoReader);
				DataStorage.ApplicationDataPathUser = TW.ApplicationSettingsFolder;
				
				mLogger = new TradeWright.Utilities.Logging.FormattingLogger("", nameof(FileAutoReader));

  				var logfileName = Logging.SetupDefaultLogging(synchronized: true);
				
				if (!checkArgs(args)) return;

				var mutexName = $"Global\\{args[2].Replace("\\", "$")}";
				if (Mutex.TryOpenExisting(mutexName, out Mutex mutex))
				{
					writeDiagnostic("This directory is already being monitored");
					return;
				}

				writeDiagnostic("FileAutoReader starting");

				using (mutex = new Mutex(false, mutexName))
				{
					mDispatcher = Task.Factory.StartNew(() =>
					{
						fileProcessingLoop(mFilenameQueue,
											args[2],
											mCancellationTokenSource.Token);
					});

					mFileSystemWatcher = startFileSystemWatcher(ensureDirectorySeparatorAtEnd(args[0]), args[1], mFilenameQueue);
					mTimer = writePeriodically();
					readInputFromConsole();

					mCancellationTokenSource.Cancel();
				}
			}
			catch (Exception ex)
			{
				writeDiagnostic($"RLK message: {ex.Message}");
				writeDiagnostic($"RLK message: {ex.StackTrace}");
			}
			finally
			{
				mTimer?.StopTimer();
				if (mCancellationTokenSource != null)
				{
					((IDisposable)mCancellationTokenSource).Dispose();
				}
			}
		}

		private static bool checkArgs(string[] args)
		{
			if (args.Length <= 2 || ((String.Compare(args[0], "/?") == 0) | (String.Compare(args[0], "-?") == 0)) || args.Length > 3)
			{
				writeToConsole("Usage: FileAutoReader ordersfolder filter [archivefolder]");
				return false;
			}
			args[2] = ensureDirectorySeparatorAtEnd(args[2]);
			return true;
		}

		private static FileSystemWatcher startFileSystemWatcher(string orderFilesFolderPath, string filter, BlockingCollection<string> filenameQueue)
		{
			writeDiagnostic($"Starting FileSystemWatcher: {orderFilesFolderPath}");
			FileSystemWatcher fileSystemWatcher = new FileSystemWatcher(orderFilesFolderPath, filter);
			fileSystemWatcher.NotifyFilter = NotifyFilters.FileName;
			fileSystemWatcher.InternalBufferSize = 65536;
			fileSystemWatcher.EnableRaisingEvents = true;
			fileSystemWatcher.Created += (s, e) =>
			{
				if (e.ChangeType == WatcherChangeTypes.Created)
				{
					var filename = e.FullPath;
					var notifyTime = DateTime.Now;
					DateTime prevNotifyTime;
					if (fileNotifications.TryGetValue(filename, out prevNotifyTime))
					{
						fileNotifications.Remove(filename);
						if ((notifyTime - prevNotifyTime) >= OneSecond)
						{
							writeDiagnostic($"File created: {filename}");
							filenameQueue.Add(filename);
						}
					}
					else
					{
						writeDiagnostic($"File created: {filename}");
						filenameQueue.Add(filename);
					}
					fileNotifications.Add(filename, notifyTime);
				}
			};
			return fileSystemWatcher;
		}

		private static async void fileProcessingLoop(BlockingCollection<string> filenameQueue, string archiveFolderPath, CancellationToken token)
		{
			try
			{
				writeDiagnostic("Starting file processing loop");
				int num = 0;
				while (!(filenameQueue.IsAddingCompleted | token.IsCancellationRequested))
				{
					string text = filenameQueue.Take(token);
					num = checked(num + 1);
					writeDiagnostic($"Take from queue file {num}: {text}");
					string archiveFilename = (archiveFolderPath.Length != 0) ? archiveFolderPath + Path.GetFileName(text) : "";
					if (!await outputFile(text, archiveFilename, num))
					{
						// we failed unexpectedly to output this file, so exit the program
						return;
					}

				}
			}
			catch (OperationCanceledException ex)
			{
				writeDiagnostic(ex.Message);
				writeDiagnostic(ex.StackTrace);
			}
			catch (ObjectDisposedException ex)
			{
				writeDiagnostic(ex.Message);
				writeDiagnostic(ex.StackTrace);
			}
			catch (InvalidOperationException ex)
			{
				writeDiagnostic(ex.Message);
				writeDiagnostic(ex.StackTrace);
			}
			catch (Exception ex)
			{
				writeDiagnostic(ex.Message);
				writeDiagnostic(ex.StackTrace);
			}
		}

		private static string ensureDirectorySeparatorAtEnd(string folderPath)
		{
			if (String.Compare(folderPath.Substring(folderPath.Length - 1), Convert.ToString(Path.DirectorySeparatorChar)) != 0)
			{
				folderPath += Path.DirectorySeparatorChar;
			}
			return folderPath;
		}

		private static async Task<bool> outputFile(string filename, string archiveFilename, int filenumber)
		{
			StreamReader sr;
			while (true)
			{
				try
				{
					sr = new StreamReader(new FileStream(filename, FileMode.Open, FileAccess.Read));
					break;
				}
				catch (FileNotFoundException)
				{
					// This should never happen: it probably means there is more than one instance
					// of this program running, and one of the others has already processed it
					writeDiagnostic($"outputFile: file not found for {filename}. Perhaps you have more than one instance of FileAutoReader running against this folder. Program will exit.");
					return false;
				}
				catch (IOException ex)
				{
					// probably means the file hasn't been closed by its creator yet - so try again
					writeDiagnostic($"outputFile: can't create StreamReader: {ex.Message}");
				}
				await Task.Delay(10);
			}
			writeDiagnostic($"Start of file {filenumber}: {filename}");
			while (!sr.EndOfStream)
			{
				writeToConsole(sr.ReadLine());
			}
			sr.Close();

			writeDiagnostic($"End of file {filenumber}: {filename}");
			if (archiveFilename.Length != 0)
			{
				try
				{
					if (File.Exists(archiveFilename)) File.Delete(archiveFilename);
					File.Move(filename, archiveFilename);
				}
				catch (IOException)
				{
					// the file has been deleted/moved by another instance
					writeDiagnostic($"outputFile: file already moved to archive {filename}. Perhaps you have more than one instance of FileAutoReader running against this folder. Program will exit.");
					return false;
				}
			}
			return true;
		}

		internal static void writeDiagnostic(string message)
		{
			mLogger.Log(message);
			writeToConsole(message, isDiagnostic: true);
		}

		private IntervalTimer writePeriodically()
		{
			writeDiagnostic("Starting periodic writer");
			var timer = TW.CreateIntervalTimer(10, ExpiryTimeUnits.ExpiryTimeUnitMilliseconds, 10);
			TW.LogMessage("Created IntervalTimer");
			timer.AddTimerExpiryListener(this);
			timer.StartTimer();
			return timer;
		}

		internal static void writeSingleLineToConsole(string message, bool isDiagnostic = false)
		{
			if (isDiagnostic)
			{
				string arg = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
				message = $"# {arg}  {message}";
			}
			mCon.WriteLine(message);
		}

		private static void writeToConsole(string message, bool isDiagnostic = false)
		{
			string[] array = message.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
			for (int i = 0; i < array.Length; i = checked(i + 1))
			{
				writeSingleLineToConsole(array[i], isDiagnostic);
			}
		}

		private static void readInputFromConsole()
		{
			writeDiagnostic("Starting reading from console");
			while (true)
			{
				var s = String.Empty;
				try
				{
					s = getInputLineAndWait(5);
					if (String.Compare(s, mCon.EofString) == 0) return;
					if (String.Compare(s.ToUpper(), "EXIT") == 0) return;
					if (s.Length != 0)
					{
						mCon.WriteLine(s);
					}
				}
				catch (COMException ex)
				{
					// ignore
					writeDiagnostic(ex.Message);
					writeDiagnostic(ex.StackTrace);
				}
			}
		}

		private static string getInputLineAndWait(int waitTimeMIllisecs = 5)
		{
			var lWaitUntilTime = DateTime.UtcNow + TimeSpan.FromMilliseconds(waitTimeMIllisecs);

			var s = mCon.ReadLine(">");
			do
			{
				// allow queued system messages to be handled
				TW.Wait(5);
			} while (DateTime.UtcNow < lWaitUntilTime);

			return s;
		}

		public void TimerExpired(ref TimerExpiredEventData ev)
		{
			try
			{
				writeSingleLineToConsole("");
			}
			catch (COMException ex)
			{
				writeDiagnostic(ex.Message);
				writeDiagnostic(ex.StackTrace);
			}
		}
	}

}
