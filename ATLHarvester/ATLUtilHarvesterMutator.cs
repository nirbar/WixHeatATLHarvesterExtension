namespace ATLCOMHarvester
{
    using System;
    using System.Collections;
    using System.IO;
    using System.Reflection;
    using System.Runtime.InteropServices;

    using Microsoft.Tools.WindowsInstallerXml;
    using Microsoft.Tools.WindowsInstallerXml.Extensions;
    using Wix = Microsoft.Tools.WindowsInstallerXml.Serialize;
    
    /// <summary>
    /// Basically the same as UtilHarvesterMutator, but with the ability to preload ATL Registrar info for DLL's.
    /// </summary>
    public class ATLUtilHarvesterMutator : MutatorExtension
    {
        // Flags for SetErrorMode() native method.
        private const UInt32 SEM_FAILCRITICALERRORS = 0x0001;
        private const UInt32 SEM_NOGPFAULTERRORBOX = 0x0002;
        private const UInt32 SEM_NOALIGNMENTFAULTEXCEPT = 0x0004;
        private const UInt32 SEM_NOOPENFILEERRORBOX = 0x8000;

        // Remember whether we were able to call OaEnablePerUserTLibRegistration
        private bool calledPerUserTLibReg;

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

        protected bool useDash;
        protected bool addShellExtensionKey;
        protected bool x64TypeLib;

        public ATLUtilHarvesterMutator()
        {
            calledPerUserTLibReg = false;

            SetErrorMode(SEM_FAILCRITICALERRORS | SEM_NOOPENFILEERRORBOX);

            try
            {
                OaEnablePerUserTLibRegistration();
                calledPerUserTLibReg = true;
            }
            catch (EntryPointNotFoundException)
            {
            }
        }

        /// <summary>
        /// True to use "-RegServer" instead of "/RegServer"
        /// </summary>
        public bool UseDash
        {
            get { return useDash; }
            set { useDash = value; }
        }

        public bool AddShellExtensionKey
        {
            get { return addShellExtensionKey; }
            set { addShellExtensionKey = value; }
        }

        /// <summary>
        /// Gets the sequence of this mutator extension.
        /// </summary>
        /// <value>The sequence of this mutator extension.</value>
        public override int Sequence
        {
            get { return 100; }
        }

        /// <summary>
        /// Mutate a WiX document.
        /// </summary>
        /// <param name="wix">The Wix document element.</param>
        public override void Mutate(Wix.Wix wix)
        {
            this.MutateElement(null, wix);
        }

        /// <summary>
        /// Mutate an element.
        /// </summary>
        /// <param name="parentElement">The parent of the element to mutate.</param>
        /// <param name="element">The element to mutate.</param>
        protected void MutateElement(Wix.IParentElement parentElement, Wix.ISchemaElement element)
        {
            if (element is Wix.File)
            {
                this.MutateFile(parentElement, (Wix.File)element);
            }

            // mutate the child elements
            if (element is Wix.IParentElement)
            {
                ArrayList childElements = new ArrayList();

                // copy the child elements to a temporary array (to allow them to be deleted/moved)
                foreach (Wix.ISchemaElement childElement in ((Wix.IParentElement)element).Children)
                {
                    childElements.Add(childElement);
                }

                foreach (Wix.ISchemaElement childElement in childElements)
                {
                    this.MutateElement((Wix.IParentElement)element, childElement);
                }
            }
        }

        /// <summary>
        /// Mutate a file.
        /// </summary>
        /// <param name="parentElement">The parent of the element to mutate.</param>
        /// <param name="file">The file to mutate.</param>
        protected void MutateFile(Wix.IParentElement parentElement, Wix.File file)
        {
			if (null == file.Source) return;

			string fileExtension = Path.GetExtension(file.Source);
			string fileSource = this.Core.ResolveFilePath(file.Source);

			if (String.Equals(".dll", fileExtension, StringComparison.OrdinalIgnoreCase) ||
				String.Equals(".ocx", fileExtension, StringComparison.OrdinalIgnoreCase)) // ActiveX
			{
				mutateDllComServer(parentElement, fileSource);
			}
			else if (String.Equals(".exe", fileExtension, StringComparison.OrdinalIgnoreCase))
			{
				mutateExeComServer(parentElement, fileSource);
			}
			else if (string.Equals(".plb", fileExtension, StringComparison.OrdinalIgnoreCase) || 
					 string.Equals(".tlb", fileExtension, StringComparison.OrdinalIgnoreCase))
			{
				// try the type library harvester
				try
				{
					ATLTypeLibraryHarvester atlTypeLibHarvester = new ATLTypeLibraryHarvester();

					this.Core.OnMessage(UtilVerboses.HarvestingTypeLib(fileSource));
					Wix.RegistryValue[] registryValues = atlTypeLibHarvester.HarvestRegistryValues(fileSource);

					foreach (Wix.RegistryValue registryValue in registryValues)
					{
						parentElement.AddChild(registryValue);
					}
				}
				catch (COMException ce)
				{
					//  0x8002801C (TYPE_E_REGISTRYACCESS)
					// If we don't have permission to harvest typelibs, it's likely because we're on
					// Vista or higher and aren't an Admin, or don't have the appropriate QFE installed.
					if (!this.calledPerUserTLibReg && (0x8002801c == unchecked((uint)ce.ErrorCode)))
					{
						this.Core.OnMessage(WixWarnings.InsufficientPermissionHarvestTypeLib());
					}
					else if (0x80029C4A == unchecked((uint)ce.ErrorCode)) // generic can't load type library
					{
						this.Core.OnMessage(UtilWarnings.TypeLibLoadFailed(fileSource, ce.Message));
					}
				}
			}
		}

		private void mutateExeComServer(Wix.IParentElement parentElement, string fileSource)
		{
			try
			{
				ATLExeHarvester exeHarvester = new ATLExeHarvester(useDash);

				this.Core.OnMessage(UtilVerboses.HarvestingSelfReg(fileSource));
				Wix.RegistryValue[] registryValues = exeHarvester.HarvestRegistryValues(fileSource);

				// Set Win64 on parent component if 64-bit PE.
				Wix.Component component = parentElement as Wix.Component;

				if ((component != null) && exeHarvester.Win64)
				{
					component.Win64 = Wix.YesNoType.yes;
				}

				foreach (Wix.RegistryValue registryValue in registryValues)
				{
					if ((Wix.RegistryValue.ActionType.write == registryValue.Action) &&
						(Wix.RegistryRootType.HKCR == registryValue.Root) &&
						string.Equals(registryValue.Key, ATLRegistryHarvester.ATLRegistrarKey))
					{
						continue; // ignore ATL Registrar values
					}
					else if (addShellExtensionKey &&
							 (Wix.RegistryValue.ActionType.write == registryValue.Action) &&
							 (Wix.RegistryRootType.HKLM == registryValue.Root) &&
							 string.Equals(registryValue.Key, ATLRegistryHarvester.ShellKey, StringComparison.InvariantCultureIgnoreCase) &&
							 string.IsNullOrEmpty(registryValue.Name))
					{
						continue; // ignore Shell Extension base key
					}
					else if ((Wix.RegistryValue.ActionType.write == registryValue.Action) &&
							 (Wix.RegistryRootType.HKCR == registryValue.Root) &&
							 registryValue.Key.StartsWith(@"CLSID\{") &&
							 registryValue.Key.EndsWith(@"}\LocalServer32", StringComparison.OrdinalIgnoreCase))
					{
						// Fix double (or double-double) quotes around LocalServer32 value, if present.

						if (registryValue.Value.StartsWith("\"") && registryValue.Value.EndsWith("\""))
						{
							registryValue.Value = registryValue.Value.Substring(1, registryValue.Value.Length - 2);
						}
						else if (registryValue.Value.StartsWith("\"\"") && registryValue.Value.EndsWith("\"\""))
						{
							registryValue.Value = registryValue.Value.Substring(2, registryValue.Value.Length - 4);
						}

						parentElement.AddChild(registryValue);
					}
					else if ((Wix.RegistryValue.ActionType.write == registryValue.Action) &&
							 (Wix.RegistryRootType.HKCR == registryValue.Root) &&
							 string.Equals(registryValue.Key, "Interface", StringComparison.OrdinalIgnoreCase))
					{
						continue; // ignore extra HKCR\Interface key
					}
					else if ((Wix.RegistryValue.ActionType.write == registryValue.Action) &&
							 (Wix.RegistryRootType.HKLM == registryValue.Root) &&
							 !registryValue.Key.StartsWith(@"SOFTWARE\Classes\Root", StringComparison.InvariantCultureIgnoreCase))
					{
						continue; // ignore anything written to HKLM for now, unless it's under SW\Classes\Root ..
					}
					else
					{
						parentElement.AddChild(registryValue);
					}
				}
			}
			catch (Exception ex)
			{
				this.Core.OnMessage(UtilWarnings.SelfRegHarvestFailed(fileSource, ex.Message));
			}
		}

		private void mutateDllComServer(Wix.IParentElement parentElement, string fileSource)
		{
			try
			{
				ATLDllHarvester dllHarvester = new ATLDllHarvester();

				this.Core.OnMessage(UtilVerboses.HarvestingSelfReg(fileSource));
				Wix.RegistryValue[] registryValues = dllHarvester.HarvestRegistryValues(fileSource, addShellExtensionKey);

				// Set Win64 on parent component if 64-bit PE.
				Wix.Component component = parentElement as Wix.Component;

				if ((component != null) && dllHarvester.Win64)
				{
					component.Win64 = Wix.YesNoType.yes;
				}

				foreach (Wix.RegistryValue registryValue in registryValues)
				{
					if ((Wix.RegistryValue.ActionType.write == registryValue.Action) &&
						(Wix.RegistryRootType.HKCR == registryValue.Root) &&
						string.Equals(registryValue.Key, ATLRegistryHarvester.ATLRegistrarKey, StringComparison.InvariantCultureIgnoreCase))
					{
						continue; // ignore ATL Registrar values
					}
					else if (addShellExtensionKey &&
							 (Wix.RegistryValue.ActionType.write == registryValue.Action) &&
							 (Wix.RegistryRootType.HKLM == registryValue.Root) &&
							 string.Equals(registryValue.Key, ATLRegistryHarvester.ShellKey, StringComparison.InvariantCultureIgnoreCase) &&
							 string.IsNullOrEmpty(registryValue.Name))
					{
						continue; // ignore Shell Extension base key
					}
					else
					{
						parentElement.AddChild(registryValue);
					}
				}
			}
			catch (TargetInvocationException tie)
			{
				if (tie.InnerException is EntryPointNotFoundException)
				{
					// No DllRegisterServer()
				}
				else
				{
					this.Core.OnMessage(UtilWarnings.SelfRegHarvestFailed(fileSource, tie.Message));
				}
			}
			catch (COMException ce)
			{
				//  0x8002801C (TYPE_E_REGISTRYACCESS)
				// If we don't have permission to harvest typelibs, it's likely because we're on
				// Vista or higher and aren't an Admin, or don't have the appropriate QFE installed.
				if (!this.calledPerUserTLibReg && (0x8002801c == unchecked((uint) ce.ErrorCode)))
				{
					this.Core.OnMessage(WixWarnings.InsufficientPermissionHarvestTypeLib());
				}
				else
				{
					this.Core.OnMessage(UtilWarnings.SelfRegHarvestFailed(fileSource, ce.Message));
				}
			}
			catch (Exception ex)
			{
				this.Core.OnMessage(UtilWarnings.SelfRegHarvestFailed(fileSource, ex.Message));
			}
		}
	}
}
