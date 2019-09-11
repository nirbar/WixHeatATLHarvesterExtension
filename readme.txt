ATL Harvester Heat Extension
----------------------------

contributed by Robert Yang (wix@rhysw.com)
9-November-2013


1. What Is It ?
---------------
This is an extension to Wix Heat (see footnote 1) that offers the following capabilities:

- Harvest COM EXE servers (32- and 64-bit) (see footnote 2)
- Harvest COM DLL servers (64-bit) (see fotnote 3)
- Preload harvesting registry with ATL and Shell Extension keys.

The installer presently works with Wix 3.5 - 3.8, but that is just the set of
versions it looks for right now.


2. How to Install it ?
----------------------
Install ATLHarvester.msi from the ATLHarvesterInstaller\bin\Release folder.
This puts the necessary binaries in the Wix installation folder and adds an
entry to heat's .exe.config file to load the extension.


3. How to Run it ?
------------------
Invoke heat.exe similarly to using the "file" harvest type, as follows :

    heat.exe atlcom [options] <DLL/OCX/EXE file to harvest> -o <output .wxs filename>

The following options are supported :

    -dash : use "-RegServer" instead of "/RegServer" when harvesting EXE servers 
            (sometimes useful for buggy old self-registration code).
    -sstd : suppress stdole2 references from output.
    -shell : add HKLM\Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved
             to registry for shell extensions

The -svb6 (suppress VB6 references) and -gg (generate GUID) options are automatically applied.


4. To Do :
----------
The command-line option parsing probably needs to be expanded for some of the more obscure
heat options (-g1, -ag, etc).  See ATLHarvesterHeatExtension.ParseOptions.

I was able do a lot of subclassing/containment to get the behavior I wanted.  One notable exception
is UtilFinalizeHarvesterMutator64 : I basically had to copy most of this code and make some small
adds/changes to get it to handle 64-bit typelibs and suppress stdole2.tlb entries.


5. How Does It Work ?
---------------------
Harvesting of COM server self-registration info is done via registry redirection, using
capabilities already found in heat.exe.  For 32-bit EXE's, a C++ DLL is injected into
the process which redirects the registry hives.  

For 64-bit DLL's and EXE's a 64-bit proxy process is used.  In the case of a 64-bit DLL
the proxy process loads the DLL and performs the harvesting.  For 64-bit EXE's the proxy
process injects a 64-bit DLL into the EXE process for harvesting.


6. How do I build it ?
----------------------
Open ATLHarvester.sln in Visual Studio 2015 and build the projects.  You need to have C#
and C++ installed, obviously.

What are all the projects for ?

- RegRedir : C++ registry redirector, for harvesting EXE servers.  Built as both 32- and 64-bit binary.
- Proxy64 : 64-bit proxy process.
- ATLHarvester : The heat extension itself.
- ATLHarvesterInstaller : Installer for all components.
- HarvesterCustomActions : Simple custom actions to backup/restore heat.exe.config.


7. Brief History
----------------
I started learning Wix in summer 2011 and wrote the 32-bit version of this extension for my employer
later that year.  Some of our projects have a large number of legacy COM components and the lack of 
EXE server harvesting in heat (see fotnote 2) had become problematic.

Just to give an idea, in our build process we have an auto-harvester process which runs 
every build and harvests all the COM and .NET assembly registration info created by heat.
This step does some preprocessing and creates a couple of large .wxs files for the installer.  
Some of the preprocessing includes:

- changing the component ID's to indicate the filename
- changing the component GUID's to be consistent from build to build (for patching)

A year or so later I moved on to a different team in the same company and added the ability
to harvest 64-bit COM servers. (see footnote 3)
We don't do a lot of this, so that functionality is not all that well-tested. Some folks on 
the wix-devs list ran some tests and I made a few changes based on their comments (such as 
addition of the Shell Extension keys).

I have permission from my employer to contribute this project to Wix as open-source, minus
any company-specific stuff, of course.

--
Thanks to Roger Orr for reviewing my initial fork and giving me helpful feedback.


----
Fotnotes:

1) heat extensions: https://blogs.msdn.microsoft.com/icumove/2009/07/02/wix-heat-extension-setting-up-a-custom-extension-project/
2) heat vs. out-of-process (exe): https://github.com/wixtoolset/issues/issues/103
3) heat.exe vs. x64 harvesting: https://github.com/wixtoolset/issues/issues/1661
