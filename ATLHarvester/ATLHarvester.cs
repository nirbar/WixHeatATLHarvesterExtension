namespace ATLCOMHarvester
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO;

    using Microsoft.Tools.WindowsInstallerXml;
    using Microsoft.Tools.WindowsInstallerXml.Tools;
    using Microsoft.Tools.WindowsInstallerXml.Extensions;
    using Wix = Microsoft.Tools.WindowsInstallerXml.Serialize;

    public class ATLHarvesterHeatExtension : HeatExtension
    {
        /// <SUMMARY>
        /// Gets the supported command line types for this extension.
        /// </SUMMARY>
        /// <VALUE>The supported command line types for this extension.</VALUE>
        public override HeatCommandLineOption[] CommandLineTypes
        {
            get
            {
                return new HeatCommandLineOption[]
                    {
                        new HeatCommandLineOption("atlcom", "harvest COM registration info from DLL/OCX/EXE using ATL registrar info"),
                        new HeatCommandLineOption("-dash",  "use -RegServer instead of /RegServer when harvesting EXE registration info (used with 'atlcom')"),
                        new HeatCommandLineOption("-sstd",  "suppress stdole2.tlb entries (used with 'atlcom')"),
                        new HeatCommandLineOption("-shell", @"add shell extension key (used with 'atlcom') :
SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved"),
                    };
            }
        }

        /// <SUMMARY>
        /// Parse the command line options for this extension.
        /// </SUMMARY>
        /// <PARAM name="type">The active harvester type.</PARAM>
        /// <PARAM name="args">The option arguments.</PARAM>
        public override void ParseOptions(string type, string[] args)
        {
            bool activated = false;

            if (String.Equals(type, "atlcom", StringComparison.OrdinalIgnoreCase))
            {
                this.Core.Harvester.Extension = new FileHarvester();
                activated = true;
            }

			if (!activated) return;

			bool useDash = false;
			bool suppressStdole2 = false;
			bool addShellExtensionKey = false;
			bool generateGuids = false;
			bool autogenerateGuids = false;
			bool suppressRootDirectory = false;
			string directoryRefId = null;
			string preprocessorVariable = null;

			// TODO: parse command-line options more completely, eg. -g1, etc.

			for(int k=0; k < args.Length; k++)
			{
				string option = args[k];
				if (String.Equals(option, "-dash", StringComparison.OrdinalIgnoreCase) ||
					String.Equals(option, "/dash", StringComparison.OrdinalIgnoreCase))
				{
					useDash = true;
				}
				else if (string.Equals(option, "-sstd", StringComparison.OrdinalIgnoreCase) ||
						 string.Equals(option, "/sstd", StringComparison.OrdinalIgnoreCase))
				{
					suppressStdole2 = true;
				}
				else if (string.Equals(option, "-shell", StringComparison.OrdinalIgnoreCase) ||
						 string.Equals(option, "/shell", StringComparison.OrdinalIgnoreCase))
				{
					addShellExtensionKey = true;
				}
				else if (string.Equals(option, "-gg", StringComparison.OrdinalIgnoreCase) ||
						 string.Equals(option, "/gg", StringComparison.OrdinalIgnoreCase))
				{
					generateGuids = true;
				}
				else if (string.Equals(option, "-ag", StringComparison.OrdinalIgnoreCase) ||
						 string.Equals(option, "/ag", StringComparison.OrdinalIgnoreCase))
				{
					autogenerateGuids = true;
				}
				else if (string.Equals(option, "-srd", StringComparison.OrdinalIgnoreCase) ||
						 string.Equals(option, "/srd", StringComparison.OrdinalIgnoreCase))
				{
					suppressRootDirectory = true;
				}
				else if (string.Equals(option, "-dr", StringComparison.OrdinalIgnoreCase) ||
						 string.Equals(option, "/dr", StringComparison.OrdinalIgnoreCase))
				{
					k++;
					if (k < args.Length) directoryRefId = args[k];
				}
				else if (string.Equals(option, "-var", StringComparison.OrdinalIgnoreCase) ||
						 string.Equals(option, "/var", StringComparison.OrdinalIgnoreCase))
				{
					k++;
					if (k < args.Length) preprocessorVariable = args[k];
				}
			}

			this.Core.Harvester.Core.RootDirectory = Path.GetDirectoryName(Path.GetFullPath(this.Core.Harvester.Core.ExtensionArgument));
			if (!suppressRootDirectory) this.Core.Harvester.Core.RootDirectory = Path.GetDirectoryName(this.Core.Harvester.Core.RootDirectory);

			// GetDirectoryName() returns null for root paths such as "c:\", so make sure to support that as well
			if (null == this.Core.Harvester.Core.RootDirectory)
			{
				this.Core.Harvester.Core.RootDirectory = Path.GetPathRoot(Path.GetDirectoryName(Path.GetFullPath(this.Core.Harvester.Core.ExtensionArgument)));
			}

			var fh = (FileHarvester)this.Core.Harvester.Extension;
			fh.SuppressRootDirectory = suppressRootDirectory;
			if (!string.IsNullOrEmpty(directoryRefId)) fh.RootedDirectoryRef = directoryRefId;

			var hm = new ATLUtilHarvesterMutator();
			hm.UseDash = useDash;
			hm.AddShellExtensionKey = addShellExtensionKey;
			this.Core.Mutator.AddExtension(hm);

			var mutator = new UtilFinalizeHarvesterMutator64();
			mutator.SuppressVB6COMElements = true;
			mutator.SuppressSTDOLE2Elements = suppressStdole2;
			if (!string.IsNullOrEmpty(preprocessorVariable)) mutator.PreprocessorVariable = preprocessorVariable;
			this.Core.Mutator.AddExtension(mutator);

			var utilMutator = new ATLUtilMutator();
			// TODO: set utilMutator options, eg. GenerateGuids
			utilMutator.GenerateGuids = generateGuids;
			utilMutator.AutogenerateGuids = autogenerateGuids;
			this.Core.Mutator.AddExtension(utilMutator);
		}
    }

    /// <summary>
    /// This class contains a UtilMutator instance and delegates to it.
    /// 
    /// Necessary because by the time ATLHarvesterHeatExtension.ParseOptions gets called, a UtilMutator is already present.
    /// </summary>
    public class ATLUtilMutator : MutatorExtension
    {
        /// <summary>
        /// Contained instance
        /// </summary>
        protected UtilMutator utilMutator;

        public ATLUtilMutator()
        {
            utilMutator = new UtilMutator();
        }

        /// <summary>
        /// Gets the sequence of the extension.
        /// </summary>
        /// <value>The sequence of the extension.</value>
        public override int Sequence
        {
            get { return 900; }
        }

        /// <summary>
        /// Mutate a WiX document.
        /// </summary>
        /// <param name="wix">The Wix document element.</param>
        public override void Mutate(Wix.Wix wix)
        {
            // delegate to contained instance.
            utilMutator.Mutate(wix);
        }

        public bool GenerateGuids
        {
            get { return utilMutator.GenerateGuids; }
            set { utilMutator.GenerateGuids = value; }
        }

        public bool AutogenerateGuids
        {
            get { return utilMutator.AutogenerateGuids; }
            set { utilMutator.AutogenerateGuids = value; }
        }
    }
}
