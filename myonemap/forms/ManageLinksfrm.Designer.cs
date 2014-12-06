namespace myonemap.forms
{
    partial class ManageLinksfrm
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.lstLinks = new System.Windows.Forms.ListBox();
            this.btnSync = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lstLinks
            // 
            this.lstLinks.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.lstLinks.FormattingEnabled = true;
            this.lstLinks.Location = new System.Drawing.Point(12, 21);
            this.lstLinks.Name = "lstLinks";
            this.lstLinks.Size = new System.Drawing.Size(555, 186);
            this.lstLinks.TabIndex = 0;
            this.lstLinks.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.lstLinks_DrawItem);
            // 
            // btnSync
            // 
            this.btnSync.Location = new System.Drawing.Point(582, 41);
            this.btnSync.Name = "btnSync";
            this.btnSync.Size = new System.Drawing.Size(138, 45);
            this.btnSync.TabIndex = 1;
            this.btnSync.Text = "Sync with Onenote";
            this.btnSync.UseVisualStyleBackColor = true;
            this.btnSync.Click += new System.EventHandler(this.btnSync_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.Location = new System.Drawing.Point(582, 111);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(138, 46);
            this.btnDelete.TabIndex = 2;
            this.btnDelete.Text = "Delete Not found Links";
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(619, 173);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 3;
            this.button3.Text = "button3";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // ManageLinksfrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(732, 476);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.btnSync);
            this.Controls.Add(this.lstLinks);
            this.Name = "ManageLinksfrm";
            this.Text = "ManageLinksfrm";
            this.Load += new System.EventHandler(this.ManageLinksfrm_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox lstLinks;
        private System.Windows.Forms.Button btnSync;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.Button button3;
    }
}