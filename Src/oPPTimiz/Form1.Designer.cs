//oPPTimiz is a powerpoint addin that allow users to reduce presentations size.
//Copyright (C) 2025 EDF
//This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.
//This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
//You should have received a copy of the GNU General Public License along with this program. If not, see https://www.gnu.org/licenses/.  

namespace oPPTimiz
{
    partial class Form1
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.labelFormDescription = new System.Windows.Forms.Label();
            this.radioMaximal = new System.Windows.Forms.RadioButton();
            this.radioIntermediate = new System.Windows.Forms.RadioButton();
            this.validateKeepButton = new System.Windows.Forms.Button();
            this.groupBoxLevel = new System.Windows.Forms.GroupBox();
            this.leafBox5 = new System.Windows.Forms.PictureBox();
            this.leafBox4 = new System.Windows.Forms.PictureBox();
            this.leafBox3 = new System.Windows.Forms.PictureBox();
            this.leafBox2 = new System.Windows.Forms.PictureBox();
            this.leafBox1 = new System.Windows.Forms.PictureBox();
            this.cancelButton = new System.Windows.Forms.Button();
            this.logoBox = new System.Windows.Forms.PictureBox();
            this.labelVersion = new System.Windows.Forms.Label();
            this.validateOverrideButton = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.groupBoxLevel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.leafBox5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.leafBox4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.leafBox3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.leafBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.leafBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.logoBox)).BeginInit();
            this.SuspendLayout();
            // 
            // labelFormDescription
            // 
            this.labelFormDescription.AutoSize = true;
            this.labelFormDescription.Location = new System.Drawing.Point(20, 28);
            this.labelFormDescription.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelFormDescription.Name = "labelFormDescription";
            this.labelFormDescription.Size = new System.Drawing.Size(443, 100);
            this.labelFormDescription.TabIndex = 0;
            this.labelFormDescription.Text = resources.GetString("labelFormDescription.Text");
            // 
            // radioMaximal
            // 
            this.radioMaximal.AutoSize = true;
            this.radioMaximal.Checked = true;
            this.radioMaximal.Location = new System.Drawing.Point(22, 32);
            this.radioMaximal.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.radioMaximal.Name = "radioMaximal";
            this.radioMaximal.Size = new System.Drawing.Size(267, 44);
            this.radioMaximal.TabIndex = 1;
            this.radioMaximal.TabStop = true;
            this.radioMaximal.Text = "Optimisé\r\n(Pour tous les usages standards)";
            this.radioMaximal.UseVisualStyleBackColor = true;
            // 
            // radioIntermediate
            // 
            this.radioIntermediate.AutoSize = true;
            this.radioIntermediate.Location = new System.Drawing.Point(321, 32);
            this.radioIntermediate.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.radioIntermediate.Name = "radioIntermediate";
            this.radioIntermediate.Size = new System.Drawing.Size(281, 44);
            this.radioIntermediate.TabIndex = 2;
            this.radioIntermediate.Text = "Standard\r\n(Pour la projection sur grand écran)";
            this.radioIntermediate.UseVisualStyleBackColor = true;
            // 
            // validateKeepButton
            // 
            this.validateKeepButton.Location = new System.Drawing.Point(411, 325);
            this.validateKeepButton.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.validateKeepButton.Name = "validateKeepButton";
            this.validateKeepButton.Size = new System.Drawing.Size(112, 63);
            this.validateKeepButton.TabIndex = 4;
            this.validateKeepButton.Text = "Optimiser et conserver";
            this.toolTip1.SetToolTip(this.validateKeepButton, "La présentation est sauvegardée et fermée.\r\nLa présentation optimisée devient la " +
        "présentation courante.\r\nElle est stockée au même endroit que la présentation ini" +
        "tiale avec le suffixe \"_oPPTimiz\".");
            this.validateKeepButton.UseVisualStyleBackColor = true;
            this.validateKeepButton.Click += new System.EventHandler(this.startOptimization);
            // 
            // groupBoxLevel
            // 
            this.groupBoxLevel.Controls.Add(this.leafBox5);
            this.groupBoxLevel.Controls.Add(this.leafBox4);
            this.groupBoxLevel.Controls.Add(this.leafBox3);
            this.groupBoxLevel.Controls.Add(this.leafBox2);
            this.groupBoxLevel.Controls.Add(this.leafBox1);
            this.groupBoxLevel.Controls.Add(this.radioMaximal);
            this.groupBoxLevel.Controls.Add(this.radioIntermediate);
            this.groupBoxLevel.Location = new System.Drawing.Point(24, 168);
            this.groupBoxLevel.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBoxLevel.Name = "groupBoxLevel";
            this.groupBoxLevel.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBoxLevel.Size = new System.Drawing.Size(622, 142);
            this.groupBoxLevel.TabIndex = 5;
            this.groupBoxLevel.TabStop = false;
            this.groupBoxLevel.Text = "Niveau d\'optimisation";
            // 
            // leafBox5
            // 
            this.leafBox5.Image = global::oPPTimiz.Properties.Resources.leaf;
            this.leafBox5.Location = new System.Drawing.Point(410, 88);
            this.leafBox5.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.leafBox5.Name = "leafBox5";
            this.leafBox5.Size = new System.Drawing.Size(30, 31);
            this.leafBox5.TabIndex = 12;
            this.leafBox5.TabStop = false;
            // 
            // leafBox4
            // 
            this.leafBox4.Image = global::oPPTimiz.Properties.Resources.leaf;
            this.leafBox4.Location = new System.Drawing.Point(370, 88);
            this.leafBox4.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.leafBox4.Name = "leafBox4";
            this.leafBox4.Size = new System.Drawing.Size(30, 31);
            this.leafBox4.TabIndex = 11;
            this.leafBox4.TabStop = false;
            // 
            // leafBox3
            // 
            this.leafBox3.Image = global::oPPTimiz.Properties.Resources.leaf;
            this.leafBox3.Location = new System.Drawing.Point(140, 88);
            this.leafBox3.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.leafBox3.Name = "leafBox3";
            this.leafBox3.Size = new System.Drawing.Size(30, 31);
            this.leafBox3.TabIndex = 10;
            this.leafBox3.TabStop = false;
            // 
            // leafBox2
            // 
            this.leafBox2.Image = global::oPPTimiz.Properties.Resources.leaf;
            this.leafBox2.Location = new System.Drawing.Point(100, 88);
            this.leafBox2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.leafBox2.Name = "leafBox2";
            this.leafBox2.Size = new System.Drawing.Size(30, 31);
            this.leafBox2.TabIndex = 9;
            this.leafBox2.TabStop = false;
            // 
            // leafBox1
            // 
            this.leafBox1.Image = global::oPPTimiz.Properties.Resources.leaf;
            this.leafBox1.Location = new System.Drawing.Point(62, 88);
            this.leafBox1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.leafBox1.Name = "leafBox1";
            this.leafBox1.Size = new System.Drawing.Size(30, 31);
            this.leafBox1.TabIndex = 8;
            this.leafBox1.TabStop = false;
            // 
            // cancelButton
            // 
            this.cancelButton.Location = new System.Drawing.Point(532, 325);
            this.cancelButton.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(112, 63);
            this.cancelButton.TabIndex = 9;
            this.cancelButton.Text = "Annuler";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.CancelOptimization);
            // 
            // logoBox
            // 
            this.logoBox.Image = global::oPPTimiz.Properties.Resources.Logo;
            this.logoBox.InitialImage = null;
            this.logoBox.Location = new System.Drawing.Point(495, 5);
            this.logoBox.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.logoBox.Name = "logoBox";
            this.logoBox.Size = new System.Drawing.Size(150, 154);
            this.logoBox.TabIndex = 7;
            this.logoBox.TabStop = false;
            // 
            // labelVersion
            // 
            this.labelVersion.AutoSize = true;
            this.labelVersion.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelVersion.Location = new System.Drawing.Point(3, 366);
            this.labelVersion.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelVersion.Name = "labelVersion";
            this.labelVersion.Size = new System.Drawing.Size(54, 17);
            this.labelVersion.TabIndex = 8;
            this.labelVersion.Text = "version";
            // 
            // validateOverrideButton
            // 
            this.validateOverrideButton.Location = new System.Drawing.Point(290, 325);
            this.validateOverrideButton.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.validateOverrideButton.Name = "validateOverrideButton";
            this.validateOverrideButton.Size = new System.Drawing.Size(112, 63);
            this.validateOverrideButton.TabIndex = 3;
            this.validateOverrideButton.Text = "Optimiser et remplacer";
            this.toolTip1.SetToolTip(this.validateOverrideButton, "L\'optimisation est faite dans la présentation courante.");
            this.validateOverrideButton.UseVisualStyleBackColor = true;
            this.validateOverrideButton.Click += new System.EventHandler(this.startOptimization);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(664, 403);
            this.Controls.Add(this.validateOverrideButton);
            this.Controls.Add(this.labelVersion);
            this.Controls.Add(this.logoBox);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.groupBoxLevel);
            this.Controls.Add(this.validateKeepButton);
            this.Controls.Add(this.labelFormDescription);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.Text = "oPPTimiz - Optimisation de fichier PowerPoint";
            this.groupBoxLevel.ResumeLayout(false);
            this.groupBoxLevel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.leafBox5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.leafBox4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.leafBox3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.leafBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.leafBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.logoBox)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelFormDescription;
        private System.Windows.Forms.RadioButton radioMaximal;
        private System.Windows.Forms.RadioButton radioIntermediate;
        private System.Windows.Forms.Button validateKeepButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.PictureBox logoBox;
        private System.Windows.Forms.PictureBox leafBox5;
        private System.Windows.Forms.PictureBox leafBox4;
        private System.Windows.Forms.PictureBox leafBox3;
        private System.Windows.Forms.PictureBox leafBox2;
        private System.Windows.Forms.PictureBox leafBox1;
        private System.Windows.Forms.Label labelVersion;
        private System.Windows.Forms.GroupBox groupBoxLevel;
        private System.Windows.Forms.Button validateOverrideButton;
        private System.Windows.Forms.ToolTip toolTip1;
    }
}