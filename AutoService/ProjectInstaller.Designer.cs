namespace AutoService
{
    partial class ProjectAutoServiceInstaller
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
            this.AutoServiceProcessInstaller = new System.ServiceProcess.ServiceProcessInstaller();
            this.FastAutoService = new System.ServiceProcess.ServiceInstaller();
            // 
            // AutoServiceProcessInstaller
            // 
            this.AutoServiceProcessInstaller.Account = System.ServiceProcess.ServiceAccount.LocalSystem;
            this.AutoServiceProcessInstaller.Password = null;
            this.AutoServiceProcessInstaller.Username = null;
            // 
            // FastAutoService
            // 
            this.FastAutoService.ServiceName = "FastAutoService";
            // 
            // ProjectAutoServiceInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.AutoServiceProcessInstaller,
            this.FastAutoService});

        }

        #endregion

        private System.ServiceProcess.ServiceProcessInstaller AutoServiceProcessInstaller;
        private System.ServiceProcess.ServiceInstaller FastAutoService;
    }
}