using System;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading;

namespace VsInterceptor {
	internal static class Program {
		// ─────────────────────────────────────────────────────────────────────
		// CONFIGURATION
		// ─────────────────────────────────────────────────────────────────────

		/// <summary>
		/// Path to Visual Studio 2026 (18.x) command-line host.
		/// Adjust if you install a different edition or path.
		/// </summary>
		private const string DefaultDevenvPath =
			@"C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\devenv.com";

		/// <summary>
		/// Global mutex to ensure there is only one server instance.
		/// </summary>
		private const string MutexName = "Global\\VsInterceptorSingleton";

		/// <summary>
		/// Named pipe for client → server messages.
		/// </summary>
		private const string PipeName = "VsInterceptorPipe";

		private static readonly string LogDirectory =
			Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
				"VsInterceptor");

		private static readonly string LogFilePath =
			Path.Combine(LogDirectory, "VsInterceptor.log");

		private static bool _isServer;
		private static bool _running = true;

		// ─────────────────────────────────────────────────────────────────────
		// NATIVE: CONSOLE + ROT
		// ─────────────────────────────────────────────────────────────────────

		[DllImport("kernel32.dll", SetLastError = true)]
		private static extern bool AllocConsole();

		[DllImport("ole32.dll")]
		private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable runningObjectTable);

		[DllImport("ole32.dll")]
		private static extern int CreateBindCtx(int reserved, out IBindCtx bindContext);

		/// <summary>
		/// Ensures a console is attached for the server so you can see live logs.
		/// </summary>
		private static void EnsureConsole() {
			try {
				_ = Console.WindowHeight; // throws if no console attached
				LogText("[EnsureConsole] Existing console detected.");
			}
			catch {
				try {
					LogText("[EnsureConsole] No console attached. Calling AllocConsole()...");
					if (!AllocConsole()) {
						LogText("[EnsureConsole] AllocConsole() FAILED.");
					} else {
						LogText("[EnsureConsole] AllocConsole() succeeded.");
					}
				}
				catch (Exception exception) {
					LogText("[EnsureConsole] Exception during AllocConsole: " + exception);
				}
			}
		}

		// ─────────────────────────────────────────────────────────────────────
		// ENTRY POINT
		// ─────────────────────────────────────────────────────────────────────

		/// <summary>
		/// Entry point. Acts either as long-lived server (first instance)
		/// or as a short-lived client that sends requests to the server.
		/// </summary>
		[STAThread]
		private static int Main(string[] args) {
			Directory.CreateDirectory(LogDirectory);

			// Optional debug flag to force server mode: "--server"
			bool forceServerMode = false;
			string[] effectiveArgs = args;

			if (effectiveArgs.Length > 0 &&
				string.Equals(effectiveArgs[0], "--server", StringComparison.OrdinalIgnoreCase)) {
				forceServerMode = true;
				if (effectiveArgs.Length > 1) {
					effectiveArgs = effectiveArgs[1..];
				} else {
					effectiveArgs = Array.Empty<string>();
				}
			}

			using var mutex = new Mutex(initiallyOwned: true, name: MutexName, out bool isNewMutexOwner);
			_isServer = forceServerMode || isNewMutexOwner;

			if (_isServer) {
				// SERVER MODE
				EnsureConsole();
				LogHeader(forceServerMode ? "SERVER START (FORCED)" : "SERVER START", effectiveArgs);

				Console.Title = "VsInterceptor – Server";
				Console.WriteLine("VsInterceptor server started.");
				Console.WriteLine("This window will stay open and log all Unity open-file requests.");
				Console.WriteLine("Press Ctrl+C to exit.\n");

				Console.CancelKeyPress += (sender, eventArgs) => {
					eventArgs.Cancel = true;
					_running = false;
					LogText("[Main] Ctrl+C received. Requesting server shutdown.");
				};

				// If Unity invoked us as server with arguments, process once
				if (effectiveArgs.Length > 0) {
					LogText("[Main] Initial server invocation has args; processing once before pipe loop.");
					SafeProcessCommand(effectiveArgs);
				}

				PipeServerLoop();
				LogText("[Main] Server loop ended. Exiting server process.");
				return 0;
			}

			// CLIENT MODE
			LogHeader("CLIENT INVOCATION", effectiveArgs);
			return SendToServer(effectiveArgs);
		}

		// ─────────────────────────────────────────────────────────────────────
		// SERVER: PIPE LOOP
		// ─────────────────────────────────────────────────────────────────────

		/// <summary>
		/// Long-lived loop that accepts commands over a named pipe,
		/// parses them, and triggers Visual Studio actions.
		/// </summary>
		private static void PipeServerLoop() {
			LogText("[PipeServerLoop] Entering pipe server loop. PipeName=" + PipeName);

			while (_running) {
				try {
					LogText("[PipeServerLoop] Creating NamedPipeServerStream...");
					using var pipeServer = new NamedPipeServerStream(
						PipeName,
						PipeDirection.In,
						maxNumberOfServerInstances: 1,
						PipeTransmissionMode.Byte,
						PipeOptions.Asynchronous
					);

					LogText("[PipeServerLoop] Waiting for client connection...");
					IAsyncResult waitResult = pipeServer.BeginWaitForConnection(null, null);

					while (_running && !waitResult.IsCompleted) {
						Thread.Sleep(50);
					}

					if (!_running) {
						LogText("[PipeServerLoop] _running == false while waiting; breaking loop.");
						break;
					}

					pipeServer.EndWaitForConnection(waitResult);
					LogText("[PipeServerLoop] Client connected.");

					using var reader = new StreamReader(pipeServer, Encoding.UTF8);
					string? messageLine = reader.ReadLine();
					LogText("[PipeServerLoop] Raw line read from pipe: " + (messageLine ?? "<null>"));

					if (string.IsNullOrWhiteSpace(messageLine)) {
						LogText("[PipeServerLoop] Line was null/whitespace; continuing.");
						continue;
					}

					string[] commandArguments = messageLine.Split(new[] { '\u001F' }, StringSplitOptions.None);
					LogText("[PipeServerLoop] Split args from client: " + string.Join(" | ", commandArguments));

					SafeProcessCommand(commandArguments);
				}
				catch (Exception exception) {
					LogText("[PipeServerLoop] ERROR: " + exception);
					Thread.Sleep(500);
				}
			}

			Console.WriteLine("Server shutting down...");
			LogText("[PipeServerLoop] Exiting pipe server loop.");
		}

		// ─────────────────────────────────────────────────────────────────────
		// CLIENT: SEND ARGS TO SERVER
		// ─────────────────────────────────────────────────────────────────────

		/// <summary>
		/// Sends the command-line arguments to the long-lived server via pipe.
		/// </summary>
		private static int SendToServer(string[] args) {
			try {
				LogText("[SendToServer] Creating NamedPipeClientStream to '" + PipeName + "'...");
				using var client = new NamedPipeClientStream(
					".",
					PipeName,
					PipeDirection.Out
				);

				LogText("[SendToServer] Connecting to server with 1000ms timeout...");
				client.Connect(timeout: 1000);
				LogText("[SendToServer] Connected to server.");

				using var writer = new StreamWriter(client, Encoding.UTF8) { AutoFlush = true };
				string lineToSend = string.Join("\u001F", args);
				LogText("[SendToServer] Sending line over pipe: " + lineToSend);
				writer.WriteLine(lineToSend);

				LogText("[SendToServer] Command sent. Client exiting normally.");
				return 0;
			}
			catch (TimeoutException) {
				LogText("[SendToServer] TimeoutException: Could not connect to server. Is it running?");
				return -1;
			}
			catch (Exception exception) {
				LogText("[SendToServer] ERROR sending to server: " + exception);
				return -1;
			}
		}

		// ─────────────────────────────────────────────────────
		// COMMAND PROCESSING
		// ─────────────────────────────────────────────────────

		private static void SafeProcessCommand(string[] args) {
			try {
				LogText("[SafeProcessCommand] Enter. Args: " + string.Join(" | ", args));
				ProcessCommand(args);
				LogText("[SafeProcessCommand] Exit OK.");
			}
			catch (Exception exception) {
				LogText("[SafeProcessCommand] ERROR: " + exception);
			}
		}

		/// <summary>
		/// Parses arguments coming from Unity and decides how to open the file.
		/// </summary>
		private static void ProcessCommand(string[] args) {
			bool forceNewInstance = false;
			string? filePath = null;
			int? lineNumber = null;
			int? columnNumber = null;

			int currentIndex = 0;

			LogText("[ProcessCommand] Starting parse of args.");

			// Flags first (e.g. --new)
			while (currentIndex < args.Length && args[currentIndex].StartsWith("--", StringComparison.OrdinalIgnoreCase)) {
				string flag = args[currentIndex];
				LogText("[ProcessCommand] Flag detected: " + flag);

				if (string.Equals(flag, "--new", StringComparison.OrdinalIgnoreCase)) {
					forceNewInstance = true;
					LogText("[ProcessCommand] forceNewInstance set TRUE.");
				}
				currentIndex++;
			}

			// Unity convention: "$(File)" "$(Line)" "$(Column)"
			if (currentIndex < args.Length) {
				filePath = args[currentIndex++];
				LogText("[ProcessCommand] Parsed file: " + filePath);
			}

			if (currentIndex < args.Length && int.TryParse(args[currentIndex], out var parsedLine)) {
				lineNumber = parsedLine;
				LogText("[ProcessCommand] Parsed line: " + lineNumber);
				currentIndex++;
			} else if (currentIndex < args.Length) {
				LogText("[ProcessCommand] Could not parse line from: " + args[currentIndex]);
				currentIndex++;
			}

			if (currentIndex < args.Length && int.TryParse(args[currentIndex], out var parsedColumn)) {
				columnNumber = parsedColumn;
				LogText("[ProcessCommand] Parsed column: " + columnNumber);
				currentIndex++;
			} else if (currentIndex < args.Length) {
				LogText("[ProcessCommand] Could not parse column from: " + args[currentIndex]);
				currentIndex++;
			}

			// Try to find a .sln for this file
			string? solutionPath = TryFindSolutionForFile(filePath);
			if (!string.IsNullOrEmpty(solutionPath)) {
				LogText($"[ProcessCommand] Resolved solution for file: '{solutionPath}' (Exists={File.Exists(solutionPath)})");
			} else {
				LogText("[ProcessCommand] No solution found for file; will launch VS without explicit .sln.");
			}

			LogText($"[ProcessCommand] Final parsed -> file='{filePath}', line={lineNumber?.ToString() ?? "null"}, column={columnNumber?.ToString() ?? "null"}, forceNew={forceNewInstance}");

			if (forceNewInstance)
				LaunchVisualStudioNewInstance(solutionPath, filePath, lineNumber, columnNumber);
			else
				LaunchVisualStudioReuseInstance(solutionPath, filePath, lineNumber, columnNumber);
		}

		// ─────────────────────────────────────────────────────
		// DTE (COM) DISCOVERY
		// ─────────────────────────────────────────────────────

		/// <summary>
		/// Attempts to find a running Visual Studio DTE instance that has the given solution loaded.
		/// </summary>
		private static object? TryGetDteForSolution(string? solutionPath) {
			LogText("[TryGetDteForSolution] ENTER. solutionPath=" + (solutionPath ?? "<null>"));
			if (string.IsNullOrEmpty(solutionPath))
				return null;

			string normalizedTarget;
			try {
				normalizedTarget = NormalizePath(solutionPath);
			}
			catch (Exception normalizationException) {
				LogText("[TryGetDteForSolution] ERROR normalizing target solution path: " + normalizationException);
				return null;
			}

			try {
				if (!TryGetRunningObjectTable(out var runningObjectTable))
					return null;

				if (!TryCreateBindContext(out var bindContext))
					return null;

				if (!TryEnumerateRunningObjects(runningObjectTable, bindContext, normalizedTarget, out var matchedDte))
					return matchedDte;

				return matchedDte;
			}
			catch (Exception exception) {
				LogText("[TryGetDteForSolution] ERROR: " + exception);
			}

			LogText("[TryGetDteForSolution] EXIT with null (no matching DTE found).");
			return null;
		}

		private static bool TryGetRunningObjectTable(out IRunningObjectTable runningObjectTable) {
			int hr = GetRunningObjectTable(0, out runningObjectTable);
			if (hr != 0 || runningObjectTable == null) {
				LogText("[TryGetDteForSolution] GetRunningObjectTable failed. hr=" + hr);
				return false;
			}
			return true;
		}

		private static bool TryCreateBindContext(out IBindCtx bindContext) {
			int hr = CreateBindCtx(0, out bindContext);
			if (hr != 0 || bindContext == null) {
				LogText("[TryGetDteForSolution] CreateBindCtx failed. hr=" + hr);
				return false;
			}
			return true;
		}

		private static bool TryEnumerateRunningObjects(
			IRunningObjectTable runningObjectTable,
			IBindCtx bindContext,
			string normalizedTargetSolution,
			out object? matchedDte) {

			matchedDte = null;
			runningObjectTable.EnumRunning(out IEnumMoniker monikerEnumerator);
			if (monikerEnumerator == null) {
				LogText("[TryGetDteForSolution] EnumRunning returned null.");
				return true;
			}

			IMoniker[] monikers = new IMoniker[1];

			while (monikerEnumerator.Next(1, monikers, IntPtr.Zero) == 0) {
				IMoniker moniker = monikers[0];
				if (moniker == null)
					continue;

				if (!TryGetMonikerDisplayName(moniker, bindContext, out var displayName))
					continue;

				LogText("[TryGetDteForSolution] ROT entry: " + displayName);

				if (!displayName.Contains("VisualStudio.DTE.", StringComparison.OrdinalIgnoreCase))
					continue;

				if (!TryGetDteForMoniker(runningObjectTable, moniker, normalizedTargetSolution, out matchedDte))
					continue;

				if (matchedDte != null)
					return true;
			}

			return true;
		}

		private static bool TryGetMonikerDisplayName(IMoniker moniker, IBindCtx bindContext, out string displayName) {
			displayName = string.Empty;
			try {
				moniker.GetDisplayName(bindContext, null, out displayName);
				return true;
			}
			catch (Exception exception) {
				LogText("[TryGetDteForSolution] ERROR getting moniker display name: " + exception);
				return false;
			}
		}

		private static bool TryGetDteForMoniker(
			IRunningObjectTable runningObjectTable,
			IMoniker moniker,
			string normalizedTargetSolution,
			out object? matchedDte) {

			matchedDte = null;

			try {
				runningObjectTable.GetObject(moniker, out object dteObject);
				if (dteObject == null) {
					LogText("[TryGetDteForSolution] ROT GetObject returned null for VS entry.");
					return false;
				}

				dynamic dteDynamic = dteObject;
				string? solutionFullName = null;

				try {
					solutionFullName = dteDynamic?.Solution?.FullName as string;
				}
				catch (Exception solutionException) {
					LogText("[TryGetDteForSolution] ERROR reading Solution.FullName: " + solutionException);
				}

				if (string.IsNullOrEmpty(solutionFullName)) {
					LogText("[TryGetDteForSolution] DTE instance has no loaded solution (Solution.FullName is null/empty).");
					return false;
				}

				string normalizedExistingSolution;
				try {
					normalizedExistingSolution = NormalizePath(solutionFullName);
				}
				catch (Exception normalizationException) {
					LogText("[TryGetDteForSolution] ERROR normalizing existing solution path: " + normalizationException);
					return false;
				}

				LogText($"[TryGetDteForSolution] DTE instance solution: '{solutionFullName}' (normalized='{normalizedExistingSolution}')");

				if (normalizedExistingSolution == normalizedTargetSolution) {
					LogText("[TryGetDteForSolution] MATCH found for solution. Returning DTE instance.");
					matchedDte = dteObject;
					return true;
				}

				LogText("[TryGetDteForSolution] Solution does not match target.");
				return false;
			}
			catch (Exception exception) {
				LogText("[TryGetDteForSolution] ERROR inspecting ROT object: " + exception);
				return false;
			}
		}

		/// <summary>
		/// Helper to normalize filesystem paths for comparison.
		/// </summary>
		private static string NormalizePath(string path) {
			return Path.GetFullPath(path)
				.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
				.ToLowerInvariant();
		}

		/// <summary>
		/// Helper to safely read a COM string property and log on error.
		/// </summary>
		private static string SafeComString(Func<string?> getter, string label) {
			try {
				var value = getter();
				return value ?? "<null>";
			}
			catch (Exception exception) {
				LogText($"[OpenFileViaDte] ERROR reading {label}: {exception}");
				return "<error>";
			}
		}

		// ─────────────────────────────────────────────────────
		// DTE FILE OPEN HELPERS
		// ─────────────────────────────────────────────────────

		/// <summary>
		/// High-level helper that uses DTE to bring VS to front and open a file.
		/// </summary>
		private static void OpenFileViaDte(object dteObject, string? filePath, int? lineNumber) {
			LogText("[OpenFileViaDte] ENTER");

			if (string.IsNullOrEmpty(filePath)) {
				LogText("[OpenFileViaDte] filePath is null/empty. Nothing to open.");
				return;
			}

			try {
				var dteType = dteObject.GetType();

				LogActiveDocumentAndDocumentsSnapshot(dteObject, dteType, "BEFORE open");

				bool fileExists = CheckFileExists(filePath);

				string solutionFullName = SafeComString(
					() => (string?)dteType
						.InvokeMember("Solution", System.Reflection.BindingFlags.GetProperty, null, dteObject, null)?
						.GetType()
						.GetProperty("FullName")?
						.GetValue(dteType.InvokeMember("Solution", System.Reflection.BindingFlags.GetProperty, null, dteObject, null))!,
					"Solution.FullName");

				LogText("[OpenFileViaDte] Using DTE with Solution='" + solutionFullName +
						"' to open file='" + filePath + "'. Exists=" + fileExists);

				TryActivateMainWindow(dteObject, dteType);

				bool openSucceeded =
					TryOpenFileWithItemOperations(dteObject, dteType, filePath) ||
					TryOpenFileWithExecuteCommand(dteObject, dteType, filePath);

				LogActiveDocumentAndDocumentsSnapshot(dteObject, dteType, "AFTER open");

				if (!openSucceeded) {
					LogText("[OpenFileViaDte] All OpenFile attempts failed; not attempting goto-line.");
				} else if (lineNumber.HasValue && lineNumber.Value > 0) {
					TryGoToLine(dteObject, dteType, lineNumber.Value);
				} else {
					LogText("[OpenFileViaDte] No line provided; skipping goto.");
				}
			}
			catch (Exception exception) {
				LogText("[OpenFileViaDte] ERROR (outer): " + exception);
			}
			finally {
				try {
					Marshal.ReleaseComObject(dteObject);
				}
				catch {
					// ignore COM release failures
				}
				LogText("[OpenFileViaDte] EXIT");
			}
		}

		private static void LogActiveDocumentAndDocumentsSnapshot(object dteObject, Type dteType, string stageLabel) {
			try {
				object? activeDocumentObject =
					dteType.InvokeMember("ActiveDocument", System.Reflection.BindingFlags.GetProperty,
						null, dteObject, null);

				string activeDocumentName = "<null>";
				string activeDocumentFullName = "<null>";

				if (activeDocumentObject != null) {
					var activeDocumentType = activeDocumentObject.GetType();
					activeDocumentName = SafeComString(
						() => activeDocumentType.GetProperty("Name")?.GetValue(activeDocumentObject) as string,
						"ActiveDocument.Name (" + stageLabel + ")");
					activeDocumentFullName = SafeComString(
						() => activeDocumentType.GetProperty("FullName")?.GetValue(activeDocumentObject) as string,
						"ActiveDocument.FullName (" + stageLabel + ")");
				}

				LogText($"[OpenFileViaDte] {stageLabel}: ActiveDocument Name='{activeDocumentName}', FullName='{activeDocumentFullName}'");

				// Snapshot of a few open documents
				LogDocumentsSnapshot(dteObject, dteType);
			}
			catch (Exception exception) {
				LogText("[OpenFileViaDte] ERROR logging active document snapshot (" + stageLabel + "): " + exception);
			}
		}

		private static void LogDocumentsSnapshot(object dteObject, Type dteType) {
			try {
				object? documentsObject = dteType.InvokeMember("Documents",
					System.Reflection.BindingFlags.GetProperty, null, dteObject, null);

				if (documentsObject == null) {
					LogText("[OpenFileViaDte] Documents is null.");
					return;
				}

				var documentsType = documentsObject.GetType();
				int documentCount = 0;

				try {
					documentCount = (int)(documentsType.GetProperty("Count")?.GetValue(documentsObject) ?? 0);
				}
				catch (Exception countException) {
					LogText("[OpenFileViaDte] ERROR reading Documents.Count: " + countException);
				}

				LogText("[OpenFileViaDte] Documents.Count=" + documentCount);

				int maxDocumentsToLog = Math.Min(documentCount, 5);
				for (int index = 1; index <= maxDocumentsToLog; index++) {
					try {
						object? documentObject = documentsType.InvokeMember("Item",
							System.Reflection.BindingFlags.InvokeMethod, null, documentsObject, new object[] { index });
						if (documentObject == null)
							continue;

						var documentType = documentObject.GetType();
						string? documentName = documentType.GetProperty("Name")?.GetValue(documentObject) as string;
						string? documentFullName = documentType.GetProperty("FullName")?.GetValue(documentObject) as string;

						LogText($"[OpenFileViaDte] Document[{index}]: Name='{documentName ?? "<null>"}', FullName='{documentFullName ?? "<null>"}'");
					}
					catch (Exception documentException) {
						LogText("[OpenFileViaDte] ERROR reading Document[" + index + "]: " + documentException);
					}
				}
			}
			catch (Exception documentsException) {
				LogText("[OpenFileViaDte] ERROR enumerating Documents: " + documentsException);
			}
		}

		private static bool CheckFileExists(string filePath) {
			try {
				return File.Exists(filePath);
			}
			catch (Exception exception) {
				LogText("[OpenFileViaDte] ERROR checking File.Exists: " + exception);
				return false;
			}
		}

		private static void TryActivateMainWindow(object dteObject, Type dteType) {
			try {
				object? mainWindowObject = dteType.InvokeMember("MainWindow",
					System.Reflection.BindingFlags.GetProperty, null, dteObject, null);
				if (mainWindowObject != null) {
					var mainWindowType = mainWindowObject.GetType();
					LogText("[OpenFileViaDte] Activating MainWindow...");
					mainWindowType.InvokeMember("Activate",
						System.Reflection.BindingFlags.InvokeMethod, null, mainWindowObject, null);
				} else {
					LogText("[OpenFileViaDte] MainWindow is null.");
				}
			}
			catch (Exception exception) {
				LogText("[OpenFileViaDte] ERROR Activating MainWindow: " + exception);
			}
		}

		private static bool TryOpenFileWithItemOperations(object dteObject, Type dteType, string filePath) {
			try {
				LogText("[OpenFileViaDte] Trying ItemOperations.OpenFile via reflection...");
				object? itemOperationsObject = dteType.InvokeMember("ItemOperations",
					System.Reflection.BindingFlags.GetProperty, null, dteObject, null);

				if (itemOperationsObject == null) {
					LogText("[OpenFileViaDte] ItemOperations is null.");
					return false;
				}

				var itemOperationsType = itemOperationsObject.GetType();
				itemOperationsType.InvokeMember("OpenFile",
					System.Reflection.BindingFlags.InvokeMethod, null, itemOperationsObject, new object[] { filePath });

				LogText("[OpenFileViaDte] ItemOperations.OpenFile reflection call succeeded.");
				return true;
			}
			catch (Exception exception) {
				LogText("[OpenFileViaDte] ERROR calling ItemOperations.OpenFile via reflection: " + exception);
				return false;
			}
		}

		private static bool TryOpenFileWithExecuteCommand(object dteObject, Type dteType, string filePath) {
			try {
				LogText("[OpenFileViaDte] Falling back to ExecuteCommand('File.OpenFile', <file>) via reflection...");
				object[] commandArguments = { "File.OpenFile", "\"" + filePath + "\"" };
				dteType.InvokeMember("ExecuteCommand",
					System.Reflection.BindingFlags.InvokeMethod, null, dteObject, commandArguments);

				LogText("[OpenFileViaDte] ExecuteCommand('File.OpenFile', ...) reflection call succeeded.");
				return true;
			}
			catch (Exception exception) {
				LogText("[OpenFileViaDte] ERROR calling ExecuteCommand('File.OpenFile', ...) via reflection: " + exception);
				return false;
			}
		}

		private static void TryGoToLine(object dteObject, Type dteType, int lineNumber) {
			try {
				object? activeDocumentObject = dteType.InvokeMember("ActiveDocument",
					System.Reflection.BindingFlags.GetProperty, null, dteObject, null);
				if (activeDocumentObject == null) {
					LogText("[OpenFileViaDte] ActiveDocument is null after OpenFile.");
					return;
				}

				var activeDocumentType = activeDocumentObject.GetType();
				object? selectionObject = activeDocumentType.GetProperty("Selection")?.GetValue(activeDocumentObject);
				if (selectionObject == null) {
					LogText("[OpenFileViaDte] ActiveDocument.Selection is null; cannot go to line.");
					return;
				}

				var selectionType = selectionObject.GetType();
				selectionType.InvokeMember("GotoLine",
					System.Reflection.BindingFlags.InvokeMethod, null, selectionObject,
					new object[] { lineNumber, true });
				LogText("[OpenFileViaDte] GotoLine(" + lineNumber + ") executed.");
			}
			catch (Exception exception) {
				LogText("[OpenFileViaDte] ERROR while trying to go to line: " + exception);
			}
		}

		// ─────────────────────────────────────────────────────
		// VS LAUNCH HELPERS
		// ─────────────────────────────────────────────────────

		/// <summary>
		/// Attempts to re-use a VS instance (by solution) and open the file,
		/// falling back to launching VS with that solution if needed.
		/// </summary>
		private static void LaunchVisualStudioReuseInstance(string? solutionPath, string? filePath, int? lineNumber, int? columnNumber) {
			LogText("[LaunchVisualStudioReuseInstance] ENTER");

			if (string.IsNullOrEmpty(solutionPath)) {
				LogText("[LaunchVisualStudioReuseInstance] No solutionPath; will fall back to /Edit into whatever VS chooses.");
				LaunchVisualStudioNewInstance(null, filePath, lineNumber, columnNumber);
				LogText("[LaunchVisualStudioReuseInstance] EXIT (no solutionPath).");
				return;
			}

			// Spawn a background STA worker so we do not block Unity.
			var workerThread = new Thread(() => LaunchOrAttachToSolutionWorker(solutionPath, filePath, lineNumber)) {
				IsBackground = true
			};
			workerThread.SetApartmentState(ApartmentState.STA);
			workerThread.Start();

			LogText("[LaunchVisualStudioReuseInstance] Spawned STA worker to handle solution/file. EXIT (main).");
		}

		/// <summary>
		/// Worker that runs on an STA thread to find or create the VS instance for a solution and open the file via DTE.
		/// </summary>
		private static void LaunchOrAttachToSolutionWorker(string solutionPath, string? filePath, int? lineNumber) {
			try {
				LogText("[LaunchOrAttachToSolutionWorker] ENTER. solution=" + solutionPath + ", file=" + (filePath ?? "<null>"));

				// 1) Try existing instance
				object? dteInstance = TryGetDteForSolution(solutionPath);
				if (dteInstance != null) {
					LogText("[LaunchOrAttachToSolutionWorker] Found existing DTE; opening file via DTE.");
					OpenFileViaDte(dteInstance, filePath, lineNumber);
					LogText("[LaunchOrAttachToSolutionWorker] EXIT (used existing DTE).");
					return;
				}

				// 2) Launch VS with solution
				LogText("[LaunchOrAttachToSolutionWorker] No existing DTE. Launching VS with solution only.");
				RunDevenv("EnsureSolutionOpen", $"\"{solutionPath}\"", waitForExit: false);

				// 3) Wait for DTE to appear and then open the file
				object? dteAfterLaunch = WaitForDteForSolution(solutionPath, 15000, 500);
				if (dteAfterLaunch != null) {
					LogText("[LaunchOrAttachToSolutionWorker] DTE appeared after launch; opening file via DTE.");
					OpenFileViaDte(dteAfterLaunch, filePath, lineNumber);
					LogText("[LaunchOrAttachToSolutionWorker] EXIT (opened via DTE after launch).");
					return;
				}

				// 4) Fallback
				LogText("[LaunchOrAttachToSolutionWorker] No DTE after waiting; falling back to devenv /Edit.");
				string fallbackArguments = BuildDevenvFallbackArguments(solutionPath, filePath, lineNumber);
				RunDevenv("OpenFileReuseFallback", fallbackArguments, waitForExit: false);
				LogText("[LaunchOrAttachToSolutionWorker] EXIT (fallback /Edit).");
			}
			catch (Exception exception) {
				LogText("[LaunchOrAttachToSolutionWorker] ERROR: " + exception);
			}
		}

		private static object? WaitForDteForSolution(string solutionPath, int maxWaitMilliseconds, int stepMilliseconds) {
			int accumulatedWait = 0;
			while (accumulatedWait < maxWaitMilliseconds) {
				Thread.Sleep(stepMilliseconds);
				accumulatedWait += stepMilliseconds;

				LogText("[LaunchOrAttachToSolutionWorker] Polling for DTE (waited " + accumulatedWait + " ms)...");
				object? dteInstance = TryGetDteForSolution(solutionPath);
				if (dteInstance != null)
					return dteInstance;
			}
			return null;
		}

		private static string BuildDevenvFallbackArguments(string solutionPath, string? filePath, int? lineNumber) {
			var argumentsBuilder = new StringBuilder();
			argumentsBuilder.Append('"').Append(solutionPath).Append('"');

			if (!string.IsNullOrEmpty(filePath)) {
				argumentsBuilder.Append(' ')
					.Append("/Edit ")
					.Append('"')
					.Append(filePath)
					.Append('"');
			}

			if (lineNumber.HasValue && lineNumber.Value > 0) {
				argumentsBuilder.Append(' ')
					.Append("/Command ")
					.Append('"')
					.Append("Edit.GoTo ")
					.Append(lineNumber.Value)
					.Append('"');
			}

			return argumentsBuilder.ToString();
		}

		/// <summary>
		/// Always opens in a new instance (or whatever VS chooses) using /Edit.
		/// </summary>
		private static void LaunchVisualStudioNewInstance(string? solutionPath, string? filePath, int? lineNumber, int? columnNumber) {
			LogText("[LaunchVisualStudioNewInstance] ENTER");

			var argumentsBuilder = new StringBuilder();

			if (!string.IsNullOrEmpty(solutionPath)) {
				LogText("[LaunchVisualStudioNewInstance] Using solutionPath=" + solutionPath);
				argumentsBuilder.Append('"')
					.Append(solutionPath)
					.Append('"')
					.Append(' ');
			} else {
				LogText("[LaunchVisualStudioNewInstance] No solutionPath; will rely on VS default instance selection.");
			}

			argumentsBuilder.Append("/Edit");

			if (!string.IsNullOrEmpty(filePath)) {
				bool fileExists = false;
				try { fileExists = File.Exists(filePath); }
				catch { }
				LogText("[LaunchVisualStudioNewInstance] Will /Edit file=" + filePath + " (Exists=" + fileExists + ")");
				argumentsBuilder.Append(' ')
					.Append('"')
					.Append(filePath)
					.Append('"');
			} else {
				LogText("[LaunchVisualStudioNewInstance] No file specified; /Edit will be called with no file.");
			}

			if (lineNumber.HasValue && lineNumber.Value > 0) {
				LogText("[LaunchVisualStudioNewInstance] Adding Edit.GoTo line=" + lineNumber);
				argumentsBuilder.Append(' ')
					.Append("/Command ")
					.Append('"')
					.Append("Edit.GoTo ")
					.Append(lineNumber.Value)
					.Append('"');
			} else {
				LogText("[LaunchVisualStudioNewInstance] No valid line provided; skipping Edit.GoTo.");
			}

			string vsArguments = argumentsBuilder.ToString();
			LogText("[LaunchVisualStudioNewInstance] Final args: " + vsArguments);

			RunDevenv("OpenFileNew", vsArguments, waitForExit: false);
			LogText("[LaunchVisualStudioNewInstance] EXIT");
		}

		/// <summary>
		/// Low-level helper to start devenv with a given set of arguments.
		/// </summary>
		private static void RunDevenv(string operationLabel, string arguments, bool waitForExit) {
			string devenvPath = ResolveDevenvPath();

			LogText($"[RunDevenv:{operationLabel}] Preparing to start devenv. Path='{devenvPath}', Args='{arguments}'");
			LogText($"[RunDevenv:{operationLabel}] File.Exists(devenvPath)={File.Exists(devenvPath)}");

			var processStartInfo = new ProcessStartInfo {
				FileName = devenvPath,
				Arguments = arguments,
				UseShellExecute = true,
				CreateNoWindow = false,
				WorkingDirectory = Path.GetDirectoryName(devenvPath) ?? Environment.CurrentDirectory
			};

			LogText($"[RunDevenv:{operationLabel}] WorkingDirectory='{processStartInfo.WorkingDirectory}'");

			try {
				using var process = new Process { StartInfo = processStartInfo };

				LogText($"[RunDevenv:{operationLabel}] Starting process...");
				process.Start();
				LogText($"[RunDevenv:{operationLabel}] Started, PID={process.Id}");

				if (waitForExit) {
					LogText($"[RunDevenv:{operationLabel}] Waiting for exit...");
					process.WaitForExit();
					LogText($"[RunDevenv:{operationLabel}] Process exited with code: {process.ExitCode}");
				} else {
					LogText($"[RunDevenv:{operationLabel}] Not waiting for exit (fire-and-forget).");
				}
			}
			catch (Exception exception) {
				LogText($"[RunDevenv:{operationLabel}] ERROR starting Visual Studio: {exception}");
			}
		}

		private static string ResolveDevenvPath() {
			try {
				if (!File.Exists(DefaultDevenvPath)) {
					LogText("[ResolveDevenvPath] WARNING: DefaultDevenvPath does not exist on disk: " + DefaultDevenvPath);
				}
			}
			catch (Exception exception) {
				LogText("[ResolveDevenvPath] Exception while checking DefaultDevenvPath: " + exception);
			}

			return DefaultDevenvPath;
		}

		// ─────────────────────────────────────────────────────
		// SOLUTION DISCOVERY (FILE → .SLN)
		// ─────────────────────────────────────────────────────

		/// <summary>
		/// Walks up the directory hierarchy from file to root, looking for a .sln.
		/// </summary>
		private static string? TryFindSolutionForFile(string? filePath) {
			LogText("[TryFindSolutionForFile] ENTER. file=" + (filePath ?? "<null>"));

			if (string.IsNullOrEmpty(filePath)) {
				LogText("[TryFindSolutionForFile] file is null/empty; returning null.");
				return null;
			}

			try {
				string? currentDirectory = Path.GetDirectoryName(filePath);
				LogText("[TryFindSolutionForFile] Initial directory: " + (currentDirectory ?? "<null>"));

				while (!string.IsNullOrEmpty(currentDirectory)) {
					LogText("[TryFindSolutionForFile] Checking directory: " + currentDirectory);

					string[] solutionFiles = Array.Empty<string>();
					try {
						solutionFiles = Directory.GetFiles(currentDirectory, "*.sln", SearchOption.TopDirectoryOnly);
					}
					catch (Exception directoryException) {
						LogText("[TryFindSolutionForFile] ERROR enumerating .sln in dir=" + currentDirectory + " -> " + directoryException);
					}

					if (solutionFiles.Length > 0) {
						LogText("[TryFindSolutionForFile] Found .sln count=" + solutionFiles.Length + " in dir=" + currentDirectory);
						foreach (var solution in solutionFiles) {
							LogText("[TryFindSolutionForFile] Candidate .sln: " + solution);
						}
						LogText("[TryFindSolutionForFile] Returning first .sln: " + solutionFiles[0]);
						return solutionFiles[0];
					}

					var parent = Directory.GetParent(currentDirectory);
					if (parent == null) {
						LogText("[TryFindSolutionForFile] Reached filesystem root; no .sln found.");
						break;
					}

					currentDirectory = parent.FullName;
				}
			}
			catch (Exception exception) {
				LogText("[TryFindSolutionForFile] ERROR: " + exception);
			}

			LogText("[TryFindSolutionForFile] EXIT with null (no solution found).");
			return null;
		}

		// ─────────────────────────────────────────────────────────────────────
		// LOGGING
		// ─────────────────────────────────────────────────────────────────────

		private static void LogHeader(string kind, string[] args) {
			var headerBuilder = new StringBuilder();
			headerBuilder.AppendLine();
			headerBuilder.AppendLine("────────────────────────────────────────────────────────");
			headerBuilder.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] VsInterceptor {kind}");
			headerBuilder.AppendLine("Raw args: " + (args == null ? "<null>" : string.Join(" | ", args)));

			AppendToLog(headerBuilder.ToString());
			Console.Write(headerBuilder.ToString());
		}

		private static void LogText(string text) {
			string line = $"[{DateTime.Now:HH:mm:ss.fff}] {text}";
			AppendToLog(line + Environment.NewLine);
			try {
				Console.WriteLine(line);
			}
			catch {
				// If no console (client mode), ignore.
			}
		}

		private static void AppendToLog(string text) {
			try {
				File.AppendAllText(LogFilePath, text, Encoding.UTF8);
			}
			catch {
				// Ignore logging failures.
			}
		}
	}
}
