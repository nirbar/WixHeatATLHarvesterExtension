namespace ATLCOMHarvester
{
    using System;
    using System.Diagnostics;
    using System.Runtime.InteropServices;
    using System.Text;

    using Microsoft.Win32;

    using Wix = Microsoft.Tools.WindowsInstallerXml.Serialize;
    using Microsoft.Tools.WindowsInstallerXml.Extensions;

    /// <summary>
    /// Similar to ATLDLLHarvester, but for an EXE server.
    /// </summary>
    public class ATLExeHarvester
    {
        bool useDash;
        bool isFile64Bit;

        public ATLExeHarvester(bool useDash)
        {
            this.useDash = useDash;
        }

        /// <summary>
        /// Harvest the registry values written by calling DllRegisterServer on the specified file.
        /// </summary>
        /// <param name="file">The file to harvest registry values from.</param>
        /// <returns>The harvested registry values.</returns>
        public Wix.RegistryValue[] HarvestRegistryValues(string file)
        {
            this.isFile64Bit = ATLDllHarvester.Is64Bit(file);

            using (var registryHarvester = new ATLRegistryHarvester(true))
            {
                try
                {
                    string options = (useDash ? "-RegServer" : "/RegServer");

                    if (isFile64Bit)
                    {
                        StartProxyProcess64(file, options);
                    }
                    else
                    {
                        ProcessWithInjectedDll.Start(file, options, "RegRedir.dll");
                    }

                    return registryHarvester.HarvestRegistry();
                }
                catch (Exception e)
                {
                    e.Data["file"] = file;
                    throw;
                }
            }
        }

        /// <summary>
        /// Return true if 64-bit file, or false otherwise.
        /// </summary>
        public bool Win64
        {
            get { return this.isFile64Bit; }
        }

        protected void StartProxyProcess64(string file, string arguments)
        {
            Process proc = new Process();
            proc.StartInfo.FileName = "ATLHarvesterProxy64.exe";
            proc.StartInfo.Arguments = string.Format("exe \"{0}\" {1}", file, arguments);
            proc.StartInfo.UseShellExecute = false;
            proc.StartInfo.CreateNoWindow = true;
            proc.StartInfo.RedirectStandardOutput = true;
            proc.StartInfo.RedirectStandardError = false;
            proc.Start();

            string proxyOutput = proc.StandardOutput.ReadToEnd();

            proc.WaitForExit();

            if (proc.ExitCode != 0)
            {
                throw new ApplicationException(proxyOutput);
            }
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
