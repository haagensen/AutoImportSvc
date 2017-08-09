using System;
using System.Collections;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Configuration.Install;

namespace AutoImportSvc
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                // https://stackoverflow.com/questions/1195478/how-to-make-a-net-windows-service-start-right-after-the-installation/1195621#1195621
                // roda o serviço normalmente
                ServiceBase[] ServicesToRun;
                ServicesToRun = new ServiceBase[] { new AutoImportSvc() };
                ServiceBase.Run(ServicesToRun);
            } else if (args.Length==1)
                {
                    switch(args[0])
                    {
                        case "-install":
                        case "-instalar":
                        case "-i":
                            InstallService();
                            StartService();
                            break;
                        case "-uninstall":
                        case "-u":
                        case "-remover":
                        case "-r":
                        case "-desinstalar":
                        case "-d":
                            StopService();
                            UninstallService();
                            break;
                        default:
                            throw new NotImplementedException();
                    }
            }
        }
        
        private static bool IsInstalled()
        {
            using (ServiceController ct = new ServiceController("AutoImportSvc"))
            {
                try
                {
                    ServiceControllerStatus status = ct.Status;
                }
                catch
                {
                    return false;
                }
                return true;
            }
        }
        
        private static bool IsRunning()
        {
            using (ServiceController ct = new ServiceController("AutoImportSvc"))
            {
                if (!IsInstalled()) return false;
                return (ct.Status == ServiceControllerStatus.Running);
            }
        }

        private static AssemblyInstaller GetInstaller()
        {
            AssemblyInstaller installer = new AssemblyInstaller(typeof(AutoImportSvc).Assembly, null);
            installer.UseNewContext = true;
            return installer;
        }

        private static void InstallService()
        {
            if (IsInstalled()) return;

            try
            {
                using (AssemblyInstaller installer = GetInstaller())
                {
                    IDictionary state = new Hashtable();
                    try
                    {
                        installer.Install(state);
                        installer.Commit(state);
                    }
                    catch
                    {
                        try
                        {
                            installer.Rollback(state);
                        }
                        catch { }
                        throw;
                    }
                }
            }
            catch
            {
                throw;
            }
        }

        private static void UninstallService()
        {
            if (!IsInstalled()) return;
            try
            {
                using (AssemblyInstaller installer = GetInstaller())
                {
                    IDictionary state = new Hashtable();
                    try
                    {
                        installer.Uninstall(state);
                    }
                    catch
                    {
                        throw;
                    }
                }
            }
            catch
            {
                throw;
            }
        }
        
        private static void StartService()
        {
            if (!IsInstalled()) return;

            using (ServiceController controller = new ServiceController("AutoImportSvc"))
            {
                try
                {
                    if (controller.Status != ServiceControllerStatus.Running)
                    {
                        controller.Start();
                        controller.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(10));
                    }
                }
                catch
                {
                    throw;
                }
            }
        }
        
        private static void StopService()
        {
            if (!IsInstalled()) return;
            using (ServiceController controller = new ServiceController("AutoImportSvc"))
            {
                try
                {
                    if (controller.Status != ServiceControllerStatus.Stopped)
                    {
                        controller.Stop();
                        controller.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(10));
                    }
                }
                catch
                {
                    throw;
                }
            }
        }
        
    }
}
