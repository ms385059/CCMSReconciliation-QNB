namespace ReconciliationScheduleService
{
    partial class ProjectInstaller
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.ReconciliationServiceProcessInstaller1 = new System.ServiceProcess.ServiceProcessInstaller();
            this.ReconciliationService = new System.ServiceProcess.ServiceInstaller();
            // 
            // ReconciliationServiceProcessInstaller1
            // 
            this.ReconciliationServiceProcessInstaller1.Account = System.ServiceProcess.ServiceAccount.LocalSystem;
            this.ReconciliationServiceProcessInstaller1.Password = null;
            this.ReconciliationServiceProcessInstaller1.Username = null;
            // 
            // ReconciliationService
            // 
            this.ReconciliationService.Description = "Schedules Batch and import data from switch and host";
            this.ReconciliationService.DisplayName = "CCMSReconciliationScheduleService";
            this.ReconciliationService.ServiceName = "CCMSReconciliationScheduleService";
            // 
            // ProjectInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.ReconciliationServiceProcessInstaller1,
            this.ReconciliationService});

        }

        #endregion

        private System.ServiceProcess.ServiceProcessInstaller ReconciliationServiceProcessInstaller1;
        private System.ServiceProcess.ServiceInstaller ReconciliationService;
    }
}