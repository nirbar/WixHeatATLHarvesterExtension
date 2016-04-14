using System;
using System.IO;

using Microsoft.Deployment.WindowsInstaller;

namespace ATLCOMHarvester
{
    public class CustomActions
    {
        /// <summary>
        /// Back up heat.exe.config on installation, before copying new file.
        /// </summary>>
        [CustomAction]
        public static ActionResult BackupConfig(Session session)
        {
            session.Log("Begin BackupConfig");

            try
            {
                string installDir = session.CustomActionData["InstallDir"];

                if (Directory.Exists(installDir))
                {
                    string src = Path.Combine(installDir, "heat.exe.config");

                    if (File.Exists(src))
                    {
                        string dest = Path.Combine(installDir, "heat.exe.config.original");

                        try
                        {
                            session.Log("Copying from {0} to {1}", src, dest);
                            File.Copy(src, dest, true);
                        }
                        catch (Exception ex)
                        {
                            session.Log("BackupConfig: error '{0}' while copying", ex.Message);
                            return ActionResult.Failure;
                        }
                    }
                    else
                    {
                        session.Log("BackupConfig: '{0}' does not exist", src);
                        return ActionResult.Failure;
                    }
                }
                else
                {
                    session.Log("BackupConfig: '{0}' does not exist", installDir);
                    return ActionResult.Failure;
                }

                return ActionResult.Success;
            }
            catch (Exception exc)
            {
                session.Log("BackupConfig exception caught: {0}", exc.Message);
                throw;
            }
            finally
            {
                session.Log("End BackupConfig");
            }
        }

        /// <summary>
        /// Restore original heat.exe.config on uninstallation, after new file removed.
        /// </summary>
        [CustomAction]
        public static ActionResult RestoreConfig(Session session)
        {
            session.Log("Begin RestoreConfig");

            try
            {
                string installDir = session.CustomActionData["InstallDir"];

                if (Directory.Exists(installDir))
                {
                    string src = Path.Combine(installDir, "heat.exe.config.original");

                    if (File.Exists(src))
                    {
                        string dest = Path.Combine(installDir, "heat.exe.config");

                        try
                        {
                            session.Log("Copying from {0} to {1}", src, dest);
                            File.Copy(src, dest, true);
                        }
                        catch (Exception ex)
                        {
                            session.Log("BackupConfig: error '{0}' while copying", ex.Message);
                            return ActionResult.Failure;
                        }
                    }
                    else
                    {
                        session.Log("RestoreConfig: '{0}' does not exist", src);
                        return ActionResult.Failure;
                    }
                }
                else
                {
                    session.Log("RestoreConfig: '{0}' does not exist", installDir);
                    return ActionResult.Failure;
                }

                return ActionResult.Success;
            }
            catch (Exception exc)
            {
                session.Log("RestoreConfig exception caught: {0}", exc.Message);
                throw;
            }
            finally
            {
                session.Log("End RestoreConfig");
            }
        }
    }
}
