
namespace EJTool
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
            this.label1 = new System.Windows.Forms.Label();
            this.tbEJPath = new System.Windows.Forms.TextBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(2, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Đường dẫn EJ";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // tbEJPath
            // 
            this.tbEJPath.Location = new System.Drawing.Point(83, 12);
            this.tbEJPath.Name = "tbEJPath";
            this.tbEJPath.Size = new System.Drawing.Size(316, 20);
            this.tbEJPath.TabIndex = 1;
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(420, 6);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(97, 31);
            this.btnBrowse.TabIndex = 2;
            this.btnBrowse.Text = "Chọn Thư Mục";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(186, 59);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(122, 30);
            this.button2.TabIndex = 3;
            this.button2.Text = "Phân Tích Dữ Liệu";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(523, 113);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.tbEJPath);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "DVN";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbEJPath;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.Button button2;
    }
}

