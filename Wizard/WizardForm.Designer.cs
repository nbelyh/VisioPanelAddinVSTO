namespace PanelAddinWizard
{
    partial class WizardForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WizardForm));
            this.checkSupportCommandBars = new System.Windows.Forms.CheckBox();
            this.checkSupportRibbon = new System.Windows.Forms.CheckBox();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.buttonNext = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.checkSupportTaskPane = new System.Windows.Forms.CheckBox();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // checkSupportCommandBars
            // 
            this.checkSupportCommandBars.AutoSize = true;
            this.checkSupportCommandBars.Location = new System.Drawing.Point(15, 197);
            this.checkSupportCommandBars.Name = "checkSupportCommandBars";
            this.checkSupportCommandBars.Size = new System.Drawing.Size(247, 17);
            this.checkSupportCommandBars.TabIndex = 5;
            this.checkSupportCommandBars.Text = "Support Command Bars (Visio 2007 and below)";
            this.checkSupportCommandBars.UseVisualStyleBackColor = true;
            // 
            // checkSupportRibbon
            // 
            this.checkSupportRibbon.AutoSize = true;
            this.checkSupportRibbon.Checked = true;
            this.checkSupportRibbon.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkSupportRibbon.Location = new System.Drawing.Point(15, 140);
            this.checkSupportRibbon.Name = "checkSupportRibbon";
            this.checkSupportRibbon.Size = new System.Drawing.Size(212, 17);
            this.checkSupportRibbon.TabIndex = 4;
            this.checkSupportRibbon.Text = "Support Ribbon (Visio 2010 and above)";
            this.checkSupportRibbon.UseVisualStyleBackColor = true;
            // 
            // buttonCancel
            // 
            this.buttonCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonCancel.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.buttonCancel.Location = new System.Drawing.Point(344, 279);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 25);
            this.buttonCancel.TabIndex = 1;
            this.buttonCancel.Text = "Cancel";
            // 
            // buttonNext
            // 
            this.buttonNext.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonNext.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.buttonNext.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.buttonNext.Location = new System.Drawing.Point(263, 279);
            this.buttonNext.Name = "buttonNext";
            this.buttonNext.Size = new System.Drawing.Size(75, 25);
            this.buttonNext.TabIndex = 0;
            this.buttonNext.Text = "Create";
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.pictureBox1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Margin = new System.Windows.Forms.Padding(0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(431, 61);
            this.panel1.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(25, 31);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(175, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Please select options for the project";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(12, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(211, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Create new VSTO-based visio addin";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(387, 12);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(32, 32);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pictureBox1.TabIndex = 13;
            this.pictureBox1.TabStop = false;
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(39, 160);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(380, 29);
            this.label3.TabIndex = 7;
            this.label3.Text = "Add a ribbon with buttons and custom images and state";
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(39, 217);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(380, 30);
            this.label4.TabIndex = 8;
            this.label4.Text = "Add a toolbar with custom images (check if you need to support old Visio)";
            // 
            // label6
            // 
            this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.label6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label6.Location = new System.Drawing.Point(0, 267);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(431, 2);
            this.label6.TabIndex = 9;
            // 
            // label5
            // 
            this.label5.Location = new System.Drawing.Point(39, 103);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(380, 29);
            this.label5.TabIndex = 6;
            this.label5.Text = "Adds a docking panel which can be controlelled with a toggle button";
            // 
            // checkSupportTaskPane
            // 
            this.checkSupportTaskPane.AutoSize = true;
            this.checkSupportTaskPane.Checked = true;
            this.checkSupportTaskPane.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkSupportTaskPane.Location = new System.Drawing.Point(15, 83);
            this.checkSupportTaskPane.Name = "checkSupportTaskPane";
            this.checkSupportTaskPane.Size = new System.Drawing.Size(118, 17);
            this.checkSupportTaskPane.TabIndex = 3;
            this.checkSupportTaskPane.Text = "Support Task Pane";
            this.checkSupportTaskPane.UseVisualStyleBackColor = true;
            // 
            // WizardForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(431, 316);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.checkSupportTaskPane);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonNext);
            this.Controls.Add(this.checkSupportCommandBars);
            this.Controls.Add(this.checkSupportRibbon);
            this.Name = "WizardForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "New Panel Addin Project";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox checkSupportCommandBars;
        private System.Windows.Forms.CheckBox checkSupportRibbon;
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.Button buttonNext;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.CheckBox checkSupportTaskPane;
    }
}