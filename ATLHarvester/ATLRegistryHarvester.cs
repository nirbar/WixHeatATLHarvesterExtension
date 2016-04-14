namespace ATLCOMHarvester
{
    using System;

    using Microsoft.Win32;
    using Microsoft.Tools.WindowsInstallerXml.Extensions;
    using Wix = Microsoft.Tools.WindowsInstallerXml.Serialize;

    /// <summary>
    /// This class extends RegistryHarvester via containment / delegation.
    /// </summary>
    public class ATLRegistryHarvester : IDisposable
    {
        protected internal const string ATLRegistrarKey = @"CLSID\{44EC053A-400F-11D0-9DCD-00A0C90391D3}\InprocServer32";
        protected internal const string ShellKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved";

        /// <summary>
        /// Contained registry harvester.
        /// </summary>
        protected RegistryHarvester registryHarvester;

        /// <summary>
        /// ATL Registrar COM registration info.
        /// </summary>
        protected string registrarPath;
        protected string registrarThreadingModel;

        /// <summary>
        /// Instantiate a new ATLRegistryHarvester.
        /// </summary>
        /// <param name="remap">Set to true to remap the entire registry to a private location for this process.</param>
        /// <param name="addShellExtensionKey">Set to true to set shell extension key in remapped registry</param>
        public ATLRegistryHarvester(bool remap, bool addShellExtensionKey = false)
        {
            // Get current ATL Registrar COM registration info
            if (remap)
            {
                GetAtlRegistrarInfo();
            }

            // Delegate to contained registry harvester.
            this.registryHarvester = new RegistryHarvester(remap);

            // Write ATL Registrar info to remapped registry
            if (remap)
            {
                SaveAtlRegistrarInfo();

                // Add Shell Extension key, if desired
                if (addShellExtensionKey)
                {
                    CreateShellExtensionKey();
                }
            }
        }

        /// <summary>
        /// Close the ATLRegistryHarvester.
        /// </summary>
        public void Close()
        {
            this.registryHarvester.Close();
        }

        /// <summary>
        /// Dispose the ATLRegistryHarvester.
        /// </summary>
        public void Dispose()
        {
            this.Close();
        }

        /// <summary>
        /// Harvest all registry roots supported by Windows Installer.
        /// </summary>
        /// <returns>The registry keys and values in the registry.</returns>
        public Wix.RegistryValue[] HarvestRegistry()
        {
            // delegate to RegistryHarvester
            return this.registryHarvester.HarvestRegistry();
        }

        /// <summary>
        /// Harvest a registry key.
        /// </summary>
        /// <param name="path">The path of the registry key to harvest.</param>
        /// <returns>The registry keys and values under the key.</returns>
        public Wix.RegistryValue[] HarvestRegistryKey(string path)
        {
            // delegate to RegistryHarvester
            return this.registryHarvester.HarvestRegistryKey(path);
        }

        /// <summary>
        /// Get ATL Registrar info from registry.
        /// </summary>
        protected void GetAtlRegistrarInfo()
        {
            RegistryKey atlKey = null;

            try
            {
                atlKey = Registry.ClassesRoot.OpenSubKey(ATLRegistrarKey);

                if (atlKey != null)
                {
                    registrarPath = (string) atlKey.GetValue("", null);
                    registrarThreadingModel = (string)atlKey.GetValue("ThreadingModel", "Both");
                }
            }
            finally
            {
                if (atlKey != null)
                {
                    atlKey.Close();
                }
            }
        }

        protected void SaveAtlRegistrarInfo()
        {
            if (!string.IsNullOrEmpty(registrarPath))
            {
                RegistryKey atlKey = null;

                try
                {
                    atlKey = Registry.ClassesRoot.CreateSubKey(ATLRegistrarKey);

                    if (atlKey != null)
                    {
                        atlKey.SetValue("", registrarPath, RegistryValueKind.String);
                        atlKey.SetValue("ThreadingModel", registrarThreadingModel);
                    }
                }
                finally
                {
                    if (atlKey != null)
                    {
                        atlKey.Close();
                    }
                }
            }
        }

        protected void CreateShellExtensionKey()
        {
            using (var shellKey = Registry.LocalMachine.CreateSubKey(ShellKey)) { } 
        }
    }
}
