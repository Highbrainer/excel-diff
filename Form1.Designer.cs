
namespace ExcelDiff
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.buttonA = new System.Windows.Forms.Button();
            this.textBoxA = new System.Windows.Forms.TextBox();
            this.textBoxB = new System.Windows.Forms.TextBox();
            this.buttonB = new System.Windows.Forms.Button();
            this.buttonLaunch = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.labelB = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.progressBarA = new System.Windows.Forms.ProgressBar();
            this.columnsComboBox = new System.Windows.Forms.ComboBox();
            this.sheetsComboBox = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // buttonA
            // 
            this.buttonA.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonA.Location = new System.Drawing.Point(494, 21);
            this.buttonA.Name = "buttonA";
            this.buttonA.Size = new System.Drawing.Size(35, 23);
            this.buttonA.TabIndex = 0;
            this.buttonA.Text = "...";
            this.buttonA.UseVisualStyleBackColor = true;
            this.buttonA.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBoxA
            // 
            this.textBoxA.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxA.Location = new System.Drawing.Point(126, 21);
            this.textBoxA.Name = "textBoxA";
            this.textBoxA.Size = new System.Drawing.Size(362, 23);
            this.textBoxA.TabIndex = 1;
            this.textBoxA.TextChanged += new System.EventHandler(this.textBoxA_TextChanged);
            // 
            // textBoxB
            // 
            this.textBoxB.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxB.Location = new System.Drawing.Point(126, 50);
            this.textBoxB.Name = "textBoxB";
            this.textBoxB.Size = new System.Drawing.Size(362, 23);
            this.textBoxB.TabIndex = 2;
            this.textBoxB.TextChanged += new System.EventHandler(this.textBoxB_TextChanged);
            // 
            // buttonB
            // 
            this.buttonB.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonB.Location = new System.Drawing.Point(494, 50);
            this.buttonB.Name = "buttonB";
            this.buttonB.Size = new System.Drawing.Size(34, 23);
            this.buttonB.TabIndex = 3;
            this.buttonB.Text = "...";
            this.buttonB.UseVisualStyleBackColor = true;
            this.buttonB.Click += new System.EventHandler(this.buttonB_Click);
            // 
            // buttonLaunch
            // 
            this.buttonLaunch.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonLaunch.Location = new System.Drawing.Point(439, 91);
            this.buttonLaunch.Name = "buttonLaunch";
            this.buttonLaunch.Size = new System.Drawing.Size(90, 26);
            this.buttonLaunch.TabIndex = 5;
            this.buttonLaunch.Text = "Go !";
            this.buttonLaunch.UseVisualStyleBackColor = true;
            this.buttonLaunch.Click += new System.EventHandler(this.buttonLaunch_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(65, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 15);
            this.label1.TabIndex = 6;
            this.label1.Text = "Fichier A";
            // 
            // labelB
            // 
            this.labelB.AutoSize = true;
            this.labelB.Location = new System.Drawing.Point(65, 53);
            this.labelB.Name = "labelB";
            this.labelB.Size = new System.Drawing.Size(52, 15);
            this.labelB.TabIndex = 7;
            this.labelB.Text = "Fichier B";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 82);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(106, 15);
            this.label2.TabIndex = 8;
            this.label2.Text = "Onglet à comparer";
            // 
            // progressBarA
            // 
            this.progressBarA.Location = new System.Drawing.Point(12, 123);
            this.progressBarA.Name = "progressBarA";
            this.progressBarA.Size = new System.Drawing.Size(516, 10);
            this.progressBarA.TabIndex = 9;
            // 
            // columnsComboBox
            // 
            this.columnsComboBox.FormattingEnabled = true;
            this.columnsComboBox.Location = new System.Drawing.Point(346, 78);
            this.columnsComboBox.Name = "columnsComboBox";
            this.columnsComboBox.Size = new System.Drawing.Size(76, 23);
            this.columnsComboBox.TabIndex = 11;
            // 
            // sheetsComboBox
            // 
            this.sheetsComboBox.FormattingEnabled = true;
            this.sheetsComboBox.Location = new System.Drawing.Point(126, 78);
            this.sheetsComboBox.Name = "sheetsComboBox";
            this.sheetsComboBox.Size = new System.Drawing.Size(108, 23);
            this.sheetsComboBox.TabIndex = 12;
            this.sheetsComboBox.TextChanged += new System.EventHandler(this.sheetsComboBox_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(261, 82);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(82, 15);
            this.label3.TabIndex = 13;
            this.label3.Text = "Colonne de tri";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(546, 140);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.sheetsComboBox);
            this.Controls.Add(this.columnsComboBox);
            this.Controls.Add(this.progressBarA);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.labelB);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.buttonLaunch);
            this.Controls.Add(this.buttonB);
            this.Controls.Add(this.textBoxB);
            this.Controls.Add(this.textBoxA);
            this.Controls.Add(this.buttonA);
            this.Name = "Form1";
            this.Text = "Onglet à comparer";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button buttonA;
        private System.Windows.Forms.TextBox textBoxA;
        private System.Windows.Forms.TextBox textBoxB;
        private System.Windows.Forms.Button buttonB;
        private System.Windows.Forms.Button buttonLaunch;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label labelB;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ProgressBar progressBarA;
        private System.Windows.Forms.ComboBox columnsComboBox;
        private System.Windows.Forms.ComboBox sheetsComboBox;
        private System.Windows.Forms.Label label3;
    }
}

