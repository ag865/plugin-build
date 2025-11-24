namespace DavigoldPowerpointAddin
{
    partial class PPMainWindow
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
            this.mainFormHost = new System.Windows.Forms.Integration.ElementHost();
            this.SuspendLayout();
            // 
            // mainFormHost
            // 
            this.mainFormHost.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainFormHost.Location = new System.Drawing.Point(0, 0);
            this.mainFormHost.Name = "mainFormHost";
            this.mainFormHost.Size = new System.Drawing.Size(438, 790);
            this.mainFormHost.TabIndex = 0;
            this.mainFormHost.Text = "elementHost1";
            this.mainFormHost.Child = null;
            // 
            // PPMainWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.mainFormHost);
            this.Name = "PPMainWindow";
            this.Size = new System.Drawing.Size(438, 790);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Integration.ElementHost mainFormHost;
    }
}
