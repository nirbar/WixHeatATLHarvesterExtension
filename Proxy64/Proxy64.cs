using System.Runtime.InteropServices.ComTypes;

namespace ATLCOMHarvester
{
	using System;
	using System.IO;
	using System.Reflection;
	using System.Reflection.Emit;
	using System.Runtime.InteropServices;
	using System.Text;

	using System.Diagnostics;
	using Microsoft.Win32;

	public class Proxy64
	{
		/// <summary>
		/// Error code indicating invalid arguments to process.
		/// </summary>
		public const int InvalidArguments = -1;
		/// <summary>
		/// Error code indicating that DLL/EXE file does not exist.
		/// </summary>
		public const int FileNonexistent = -2;
		/// <summary>
		/// Error code indicating that attempt to invoke DllRegisterServer on DLL failed.
		/// </summary>
		public const int DLLError = -3;
		/// <summary>
		/// Error code indicating failure to start EXE.
		/// </summary>
		public const int EXEError = -4;

		/// <summary>
		/// Name of 64-bit registry redirection DLL.
		/// </summary>
		private const string RegRedirDLL = "RegRedir64.dll";

		static string sCandidatePath = "";


		/// <summary>
		/// Entry point for proxy process.
		/// </summary>
		/// <param name="args">see ShowHelp below</param>
		/// <returns>error code, or 0 if no error</returns>
		public static int Main(string[] args)
		{
			if (args.Length < 2)
			{
				ShowHelp();
				return InvalidArguments;
			}
			else
			{
				if (args[0].Equals("dll", StringComparison.InvariantCultureIgnoreCase))
				{
					//StartRegistryRedirection();

					if (File.Exists(args[1]))
					{
						return LoadDLL(args[1]);
					}
					else
					{
						return FileNonexistent;
					}
				}
				else if (args[0].Equals("exe", StringComparison.InvariantCultureIgnoreCase))
				{
					if (File.Exists(args[1]))
					{
						string arguments = string.Empty;

						if (args.Length > 2)
						{
							arguments = args[2];
						}

						return StartEXE(args[1], arguments);
					}
					else
					{
						return FileNonexistent;
					}
				}
				else
				{
					ShowHelp();
					return InvalidArguments;
				}
			}
		}

		/// <summary>
		/// Show usage/parameters.
		/// </summary>
		static void ShowHelp()
		{
			Console.WriteLine("Usage: ATLHarvesterProxy64 <exe|dll> <path to EXE or DLL> [EXE arguments]");
			Console.WriteLine();
		}

		private static bool redirectionStarted = false;
		/// <summary>
		/// Load registry redirection DLL.
		/// </summary>
		static void StartRegistryRedirection()
		{
			if (!redirectionStarted)
			{
				NativeMethods.LoadLibrary(RegRedirDLL);

				redirectionStarted = true;
			}
		}

		/// <summary>
		/// Load DLL and invoke its DllRegisterServer method.
		/// </summary>
		/// <param name="file">path to DLL</param>
		/// <returns>error code, or 0 if no error</returns>
		static int LoadDLL(string file)
		{
			//SCAN-5789 and UBMT-25108 (and also follow-up to dependency loading UBMT-27890)
			sCandidatePath = Path.GetDirectoryName(file);
			AppDomain myDomain = AppDomain.CurrentDomain;
			myDomain.AssemblyResolve += new ResolveEventHandler(MyResolveEventHandler2);

			try
			{//try to harvest .Net x64 dll
				Assembly assembly = Assembly.LoadFile(file);

				// must call this before overriding registry hives to prevent binding failures
				// on exported types during RegisterAssembly
				assembly.GetExportedTypes();

				//the redirection MUST be loaded after loading the assembly
				//otherwise, the load crashes with COMException 0x80090017 (NTE_PROV_TYPE_NOT_DEF)
				StartRegistryRedirection();
				NativeMethods.SetPerUserTypelibs();

				RegistrationServices regSvcs = new RegistrationServices();
				regSvcs.RegisterAssembly(assembly, AssemblyRegistrationFlags.SetCodeBase);
			}
			catch (BadImageFormatException)
			{//try to harvest win64 dll exporting DllRegisterServer

				NativeMethods.SetPerUserTypelibs();

				NativeMethods.LoadLibrary(file);
				StartRegistryRedirection();

				try
				{
					DynamicPInvoke(file, "DllRegisterServer", typeof(int), null, null);
				}
				catch (Exception ex)
				{
					Console.WriteLine(ex);
					return DLLError;
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
				return DLLError;
			}

			return 0;
		}

		private static Assembly MyResolveEventHandler2(object sender, ResolveEventArgs args)
		{
			try
			{
				AssemblyName assemblyName = new AssemblyName(args.Name);
				string assemblyFileName = assemblyName.Name + ".dll";

				string directory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
				string assemblyPath = Path.Combine(sCandidatePath, assemblyFileName);

				if (!File.Exists(assemblyPath))
				{
					if (assemblyPath.EndsWith(".resources.dll", StringComparison.OrdinalIgnoreCase))
					{
						Console.WriteLine("CustomResolve: ignore resource: " + assemblyPath);
						Trace.WriteLine("Ignore non-existing assembly: " + assemblyPath); //DEBUG
						return null;
					}
					else
					{
						Console.WriteLine("PCG: fallback to EXE");
						assemblyFileName = assemblyName.Name + ".exe";
						assemblyPath = Path.Combine(sCandidatePath, assemblyFileName);
					}
				}

				if (!File.Exists(assemblyPath))
				{
					Console.WriteLine("CustomResolve: not found: " + args.Name);
					return null;
				}
				else
				{
					return Assembly.LoadFrom(assemblyPath);
				}
			}
			catch (System.Exception ex)
			{
				Console.WriteLine("CustomResolve: exception: " + ex.Message);
				Trace.WriteLine("Exception while resolving assembly: " + ex.Message); //ERROR
				return null;
			}
		}

		/// <summary>
		/// Start EXE with specified arguments (typically "/RegServer"), and inject
		/// 64-bit registry redirector DLL.
		/// </summary>
		/// <param name="file">path to EXE</param>
		/// <param name="arguments">arguments to EXE</param>
		/// <returns>error code, or 0 if no error</returns>
		static int StartEXE(string file, string arguments)
		{
			try
			{
				ProcessWithInjectedDll.Start(file, arguments, RegRedirDLL);
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
				return EXEError;
			}

			return 0;
		}

		/// <summary>
		/// Dynamically PInvokes into a DLL.
		/// </summary>
		/// <param name="dll">Dynamic link library containing the entry point.</param>
		/// <param name="entryPoint">Entry point into dynamic link library.</param>
		/// <param name="returnType">Return type of entry point.</param>
		/// <param name="parameterTypes">Type of parameters to entry point.</param>
		/// <param name="parameterValues">Value of parameters to entry point.</param>
		/// <returns>Value from invoked code.</returns>
		private static object DynamicPInvoke(string dll, string entryPoint, Type returnType, Type[] parameterTypes, object[] parameterValues)
		{
			AssemblyName assemblyName = new AssemblyName();
			assemblyName.Name = "wixTempAssembly";

			AssemblyBuilder dynamicAssembly = AppDomain.CurrentDomain.DefineDynamicAssembly(assemblyName, AssemblyBuilderAccess.Run);
			ModuleBuilder dynamicModule = dynamicAssembly.DefineDynamicModule("wixTempModule");

			MethodBuilder dynamicMethod = dynamicModule.DefinePInvokeMethod(entryPoint, dll, MethodAttributes.Static | MethodAttributes.Public | MethodAttributes.PinvokeImpl, CallingConventions.Standard, returnType, parameterTypes, CallingConvention.Winapi, CharSet.Ansi);
			dynamicModule.CreateGlobalFunctions();

			MethodInfo methodInfo = dynamicModule.GetMethod(entryPoint);
			return methodInfo.Invoke(null, parameterValues);
		}

		/// <summary>
		/// Native methods for loading libraries.
		/// </summary>
		private sealed class NativeMethods
		{
			private const UInt32 LOAD_WITH_ALTERED_SEARCH_PATH = 0x00000008;

			/// <summary>
			/// Load a DLL library.
			/// </summary>
			/// <param name="file">The file name of the executable module.</param>
			/// <returns>If the function succeeds, the return value is a handle to the mapped executable module.</returns>
			internal static IntPtr LoadLibrary(string file)
			{
				IntPtr dllHandle = LoadLibraryEx(file, IntPtr.Zero, NativeMethods.LOAD_WITH_ALTERED_SEARCH_PATH);

				if (IntPtr.Zero == dllHandle)
				{
					int lastError = Marshal.GetLastWin32Error();
					throw new Exception(String.Format("Unable to load file: {0}, error: {1}", file, lastError));
				}

				return dllHandle;
			}

			/// <summary>
			/// Set typelibs to be registered per-user so that 'run as admin' isn't needed.
			/// </summary>
			internal static void SetPerUserTypelibs()
			{
				SetErrorMode(SEM_FAILCRITICALERRORS | SEM_NOOPENFILEERRORBOX);

				try
				{
					OaEnablePerUserTLibRegistration();
				}
				catch (EntryPointNotFoundException)
				{
				}
			}

			/// <summary>
			/// Maps the specified executable module into the address space of the calling process.
			/// </summary>
			/// <param name="file">The file name of the executable module.</param>
			/// <param name="fileHandle">This parameter is reserved for future use. It must be NULL.</param>
			/// <param name="flags">Action to take when loading the module.</param>
			/// <returns>If the function succeeds, the return value is a handle to the mapped executable module.</returns>
			[DllImport("kernel32.dll", CallingConvention = CallingConvention.Winapi, CharSet = CharSet.Unicode, SetLastError = true)]
			private static extern IntPtr LoadLibraryEx(string file, IntPtr fileHandle, UInt32 flags);

			// Flags for SetErrorMode() native method.
			private const UInt32 SEM_FAILCRITICALERRORS = 0x0001;
			private const UInt32 SEM_NOGPFAULTERRORBOX = 0x0002;
			private const UInt32 SEM_NOALIGNMENTFAULTEXCEPT = 0x0004;
			private const UInt32 SEM_NOOPENFILEERRORBOX = 0x8000;

			/// <summary>
			/// allow process to handle serious system errors.
			/// </summary>
			[DllImport("Kernel32.dll")]
			private static extern void SetErrorMode(UInt32 uiMode);

			/// <summary>
			/// enable the RegisterTypeLib API to use the appropriate override mapping for non-admin users on Vista
			/// </summary>
			[DllImport("Oleaut32.dll")]
			private static extern void OaEnablePerUserTLibRegistration();
		}

		/// <summary>
		/// Found on the net from a patched version of Tallow (Wix 2.0).
		/// </summary>
		internal class ProcessWithInjectedDll
		{
			[DllImport("kernel32.dll")]
			private static extern int CloseHandle(IntPtr hObject);

			private struct PROCESS_INFORMATION
			{
				public IntPtr hProcess;
				public IntPtr hThread;
				public uint dwProcessId;
				public uint dwThreadId;
			}

			private struct STARTUPINFO
			{
				public uint cb;
				public string lpReserved;
				public string lpDesktop;
				public string lpTitle;
				public uint dwX;
				public uint dwY;
				public uint dwXSize;
				public uint dwYSize;
				public uint dwXCountChars;
				public uint dwYCountChars;
				public uint dwFillAttribute;
				public uint dwFlags;
				public short wShowWindow;
				public short cbReserved2;
				public IntPtr lpReserved2;
				public IntPtr hStdInput;
				public IntPtr hStdOutput;
				public IntPtr hStdError;
			}

			private static uint CREATE_SUSPENDED = 0x4;
			private static uint NORMAL_PRIORITY_CLASS = 0x20;
			private static uint STARTF_USESHOWWINDOW = 0x00000001;
			private static short SW_NORMAL = 1;

			[DllImport("kernel32.dll")]
			private static extern bool CreateProcess(
				string lpApplicationName,
				string lpCommandLine,
				IntPtr lpProcessAttributes,
				IntPtr lpThreadAttributes,
				bool bInheritHandles,
				uint dwCreationFlags,
				IntPtr lpEnvironment,
				string lpCurrentDirectory,
				ref STARTUPINFO lpStartupInfo,
				out PROCESS_INFORMATION lpProcessInformation);

			[DllImport("kernel32.dll")]
			private static extern int ResumeThread(IntPtr hThread);

			private static uint PROCESS_CREATE_THREAD = 0x0002;
			private static uint PROCESS_VM_OPERATION = 0x0008;
			private static uint PROCESS_VM_READ = 0x0010;
			private static uint PROCESS_VM_WRITE = 0x0020;
			private static uint PROCESS_QUERY_INFORMATION = 0x0400;

			[DllImport("kernel32.dll")]
			private static extern IntPtr OpenProcess(
				uint fdwAccess,
				bool fInherit,
				uint IDProcess);


			private static uint MEM_COMMIT = 0x1000;
			private static uint PAGE_READWRITE = 0x04;

			[DllImport("kernel32.dll")]
			private static extern IntPtr VirtualAllocEx(
				  IntPtr hProcess,
				  IntPtr lpAddress,
				  uint dwSize,
				  uint flAllocationType,
				  uint flProtect);

			[DllImport("kernel32.dll")]
			private static extern bool WriteProcessMemory(
				IntPtr hProcess,
				IntPtr lpBaseAddress,
				string lpBuffer,
				uint nSize,
				ref uint lpNumberOfBytesWritten);

			[DllImport("kernel32.dll")]
			private static extern IntPtr GetModuleHandle(string lpModuleName);

			[DllImport("kernel32.dll")]
			private static extern IntPtr GetProcAddress(
				IntPtr hModule,
				string lpProcName);

			[DllImport("kernel32.dll")]
			private static extern IntPtr CreateRemoteThread(
				IntPtr hProcess,
				IntPtr lpThreadAttributes,
				uint dwStackSize,
				IntPtr lpStartAddress,
				IntPtr lpParameter,
				uint dwCreationFlags,
				IntPtr lpThreadId);

			private static uint INFINITE = (uint)0xffffffff;

			[DllImport("kernel32.dll")]
			private static extern uint WaitForSingleObject(
				IntPtr hHandle,
				uint dwMilliseconds);

			[DllImport("kernel32.dll")]
			private static extern uint GetLastError();

			[DllImport("kernel32.dll")]
			private static extern uint SearchPath(
				string lpPath,
				string lpFileName,
				string lpExtension,
				uint nBufferLength,
				StringBuilder lpBuffer,
				IntPtr lpFilePart);

			internal static void Start(string exeFileName, string arguments, string dllFileName)
			{
				STARTUPINFO si;
				PROCESS_INFORMATION pi;
				pi.hProcess = IntPtr.Zero;
				pi.hThread = IntPtr.Zero;
				IntPtr hRemoteProcess = IntPtr.Zero;
				IntPtr hRemoteThread = IntPtr.Zero;
				IntPtr lpBaseAddr;
				uint dwWritten = 0;
				IntPtr pfnLoadLibrary;

				try
				{
					// try to find dll in path
					StringBuilder lpBuffer = new StringBuilder(1000);
					if (SearchPath(null, dllFileName, null, (uint)lpBuffer.Capacity, lpBuffer, IntPtr.Zero) == 0)
					{
						throw new ApplicationException(String.Format("Unable to find dll: {0}", dllFileName));
					}

					dllFileName = lpBuffer.ToString();
					// create process in suspended mode
					si = new STARTUPINFO();
					si.cb = (uint)Marshal.SizeOf(typeof(STARTUPINFO));
					si.dwFlags = STARTF_USESHOWWINDOW;
					si.wShowWindow = SW_NORMAL;
					//                    if (!CreateProcess(exeFileName, arguments, IntPtr.Zero, IntPtr.Zero, false,
					if (!CreateProcess(null, exeFileName + " " + arguments, IntPtr.Zero, IntPtr.Zero, false,
							CREATE_SUSPENDED | NORMAL_PRIORITY_CLASS, IntPtr.Zero, null, ref si, out pi))
					{
						throw new ApplicationException(String.Format("CreateProcess failed, last error={0}", GetLastError()));
					}

					// insert dll file name in remote process
					hRemoteProcess = OpenProcess(PROCESS_CREATE_THREAD | PROCESS_QUERY_INFORMATION | PROCESS_VM_OPERATION |
							PROCESS_VM_WRITE | PROCESS_VM_READ, false, pi.dwProcessId);
					if (hRemoteProcess == IntPtr.Zero)
					{
						throw new ApplicationException(String.Format("OpenProcess failed, last error={0}", GetLastError()));
					}

					lpBaseAddr = VirtualAllocEx(hRemoteProcess, IntPtr.Zero, (uint)dllFileName.Length + 1, MEM_COMMIT, PAGE_READWRITE);
					if (lpBaseAddr == IntPtr.Zero)
					{
						throw new ApplicationException(String.Format("VirtualAllocEx failed, last error={0}", GetLastError()));
					}

					if (!WriteProcessMemory(hRemoteProcess, lpBaseAddr, dllFileName, (uint)dllFileName.Length + 1, ref dwWritten))
					{
						throw new ApplicationException(String.Format("WriteProcessMemory failed, last error={0}", GetLastError()));
					}

					pfnLoadLibrary = GetProcAddress(GetModuleHandle("kernel32.dll"), "LoadLibraryA");
					if (pfnLoadLibrary == IntPtr.Zero)
					{
						throw new ApplicationException(String.Format("GetProcAddress failed, last error={0}", GetLastError()));
					}

					hRemoteThread = CreateRemoteThread(hRemoteProcess, IntPtr.Zero, 0, pfnLoadLibrary, lpBaseAddr, 0, IntPtr.Zero);
					if (hRemoteThread == IntPtr.Zero)
					{
						throw new ApplicationException(String.Format("CreateRemoteThread failed, last error={0}", GetLastError()));
					}

					WaitForSingleObject(hRemoteThread, INFINITE);

					ResumeThread(pi.hThread);
					WaitForSingleObject(pi.hProcess, INFINITE);
				}
				finally
				{
					if (hRemoteThread != IntPtr.Zero)
						CloseHandle(hRemoteThread);
					if (hRemoteProcess != IntPtr.Zero)
						CloseHandle(hRemoteProcess);
					if (pi.hProcess != IntPtr.Zero)
						CloseHandle(pi.hProcess);
					if (pi.hThread != IntPtr.Zero)
						CloseHandle(pi.hThread);
				}
			}
		}
	}
}
