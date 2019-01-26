using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.ServiceProcess;
using System.Threading.Tasks;

namespace Arxhivista
{
    [RunInstaller(true)]
    public partial class ArxhivistaInstaller : System.Configuration.Install.Installer
    {
        private ServiceProcessInstaller serviceProcessInstaller;
        private ServiceInstaller serviceInstaller;

        public ArxhivistaInstaller()
        {
            InitializeComponent();

            serviceProcessInstaller = new ServiceProcessInstaller();
            serviceInstaller = new ServiceInstaller();
            // Here you can set properties on serviceProcessInstaller
            //or register event handlers
            serviceProcessInstaller.Account = ServiceAccount.User;

            serviceInstaller.ServiceName = Arxhivista.SERVICE_NAME;
            this.Installers.AddRange(new Installer[] {
                serviceProcessInstaller, serviceInstaller
            });
        }
    }
}
