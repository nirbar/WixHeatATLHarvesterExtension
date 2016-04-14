namespace ATLCOMHarvester
{
    using System;
    using System.Diagnostics;
    using System.IO;
    using System.Reflection;
    using System.Reflection.Emit;
    using System.Runtime.InteropServices;

    using Wix = Microsoft.Tools.WindowsInstallerXml.Serialize;

    /// <summary>
    /// This is basically the same as DllHarvester, but uses an ATLRegistryHarvester instead.
    /// </summary>
    public class ATLDllHarvester
    {
        private bool isFile64Bit;

        /// <summary>
        /// Harvest the registry values written by calling DllRegisterServer on the specified file.
        /// </summary>
        /// <param name="file">The file to harvest registry values from.</param>
        /// <param name="addShellExtensionKey">true to add shell extension key to registry before harvesting</param>
        /// <returns>The harvested registry values.</returns>
        public Wix.RegistryValue[] HarvestRegistryValues(string file, bool addShellExtensionKey)
        {
            this.isFile64Bit = Is64Bit(file);

            // load the DLL if 32-bit.
            if (!isFile64Bit)
            {
                NativeMethods.LoadLibrary(file);
            }

            using (var registryHarvester = new ATLRegistryHarvester(true, addShellExtensionKey))
            {
                try
                {
                    if (isFile64Bit)
                    {
                        StartProxyProcess64(file);
                    }
                    else
                    {
                        DynamicPInvoke(file, "DllRegisterServer", typeof(int), null, null);
                    }

                    return registryHarvester.HarvestRegistry();
                }
                catch (TargetInvocationException e)
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

        internal enum MachineType : ushort
        {
            IMAGE_FILE_MACHINE_UNKNOWN = 0x0,
            IMAGE_FILE_MACHINE_AM33 = 0x1d3,
            IMAGE_FILE_MACHINE_AMD64 = 0x8664,
            IMAGE_FILE_MACHINE_ARM = 0x1c0,
            IMAGE_FILE_MACHINE_EBC = 0xebc,
            IMAGE_FILE_MACHINE_I386 = 0x14c,
            IMAGE_FILE_MACHINE_IA64 = 0x200,
            IMAGE_FILE_MACHINE_M32R = 0x9041,
            IMAGE_FILE_MACHINE_MIPS16 = 0x266,
            IMAGE_FILE_MACHINE_MIPSFPU = 0x366,
            IMAGE_FILE_MACHINE_MIPSFPU16 = 0x466,
            IMAGE_FILE_MACHINE_POWERPC = 0x1f0,
            IMAGE_FILE_MACHINE_POWERPCFP = 0x1f1,
            IMAGE_FILE_MACHINE_R4000 = 0x166,
            IMAGE_FILE_MACHINE_SH3 = 0x1a2,
            IMAGE_FILE_MACHINE_SH3DSP = 0x1a3,
            IMAGE_FILE_MACHINE_SH4 = 0x1a6,
            IMAGE_FILE_MACHINE_SH5 = 0x1a8,
            IMAGE_FILE_MACHINE_THUMB = 0x1c2,
            IMAGE_FILE_MACHINE_WCEMIPSV2 = 0x169,
        }

        internal static MachineType GetMachineType(string file)
        {
            //see http://www.microsoft.com/whdc/system/platform/firmware/PECOFF.mspx
            //offset to PE header is always at 0x3C
            //PE header starts with "PE\0\0" =  0x50 0x45 0x00 0x00
            //followed by 2-byte machine type field (see document above for enum)
            using (var fs = new FileStream(file, FileMode.Open, FileAccess.Read))
            {
                using (var br = new BinaryReader(fs))
                {
                    fs.Seek(0x3c, SeekOrigin.Begin);
                    Int32 peOffset = br.ReadInt32();
                    fs.Seek(peOffset, SeekOrigin.Begin);
                    UInt32 peHead = br.ReadUInt32();
                    if (peHead != 0x00004550) // "PE\0\0", little-endian
                        throw new Exception("Can't find PE header");
                    return (MachineType)br.ReadUInt16();
                }
            }
        }

        internal static bool Is64Bit(string filePath)
        {
            var machineType = GetMachineType(filePath);

            if ((machineType == MachineType.IMAGE_FILE_MACHINE_AMD64) ||
                (machineType == MachineType.IMAGE_FILE_MACHINE_IA64))
            {
                return true;
            }
            else
            {
                return false;
            }
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

        protected void StartProxyProcess64(string file)
        {
            Process proc = new Process();
            proc.StartInfo.FileName = "ATLHarvesterProxy64.exe";
            proc.StartInfo.Arguments = string.Format("dll \"{0}\"", file);
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
            /// Maps the specified executable module into the address space of the calling process.
            /// </summary>
            /// <param name="file">The file name of the executable module.</param>
            /// <param name="fileHandle">This parameter is reserved for future use. It must be NULL.</param>
            /// <param name="flags">Action to take when loading the module.</param>
            /// <returns>If the function succeeds, the return value is a handle to the mapped executable module.</returns>
            [DllImport("kernel32.dll", CallingConvention = CallingConvention.Winapi, CharSet = CharSet.Unicode, SetLastError = true)]
            private static extern IntPtr LoadLibraryEx(string file, IntPtr fileHandle, UInt32 flags);
        }
    }
}
