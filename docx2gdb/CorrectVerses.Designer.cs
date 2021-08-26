namespace docx2gdb
{
    partial class CorrectVerses
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
            this.txtRight = new System.Windows.Forms.TextBox();
            this.txtLeft = new System.Windows.Forms.TextBox();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnAbort = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(347, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "مصرعها را به گونه‌ای مرتب کنید که تعداد مصرعهای راست و چپ برابر شود:";
            // 
            // txtRight
            // 
            this.txtRight.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.txtRight.Location = new System.Drawing.Point(15, 35);
            this.txtRight.Multiline = true;
            this.txtRight.Name = "txtRight";
            this.txtRight.Size = new System.Drawing.Size(296, 463);
            this.txtRight.TabIndex = 1;
            this.txtRight.TextChanged += new System.EventHandler(this.txtRightOrLeft_TextChanged);
            // 
            // txtLeft
            // 
            this.txtLeft.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.txtLeft.Location = new System.Drawing.Point(357, 35);
            this.txtLeft.Multiline = true;
            this.txtLeft.Name = "txtLeft";
            this.txtLeft.Size = new System.Drawing.Size(296, 463);
            this.txtLeft.TabIndex = 2;
            this.txtLeft.TextChanged += new System.EventHandler(this.txtRightOrLeft_TextChanged);
            // 
            // btnOK
            // 
            this.btnOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.Yes;
            this.btnOK.Enabled = false;
            this.btnOK.Location = new System.Drawing.Point(15, 514);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 25);
            this.btnOK.TabIndex = 3;
            this.btnOK.Text = "تأیید";
            this.btnOK.UseVisualStyleBackColor = true;
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.No;
            this.btnCancel.Location = new System.Drawing.Point(106, 515);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(205, 25);
            this.btnCancel.TabIndex = 4;
            this.btnCancel.Text = "نادیده انگاشتن این شعر";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnAbort
            // 
            this.btnAbort.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnAbort.DialogResult = System.Windows.Forms.DialogResult.Abort;
            this.btnAbort.Location = new System.Drawing.Point(448, 515);
            this.btnAbort.Name = "btnAbort";
            this.btnAbort.Size = new System.Drawing.Size(205, 25);
            this.btnAbort.TabIndex = 5;
            this.btnAbort.Text = "انصراف از بازخوانی باقی شعرها";
            this.btnAbort.UseVisualStyleBackColor = true;
            // 
            // CorrectVerses
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(665, 552);
            this.Controls.Add(this.btnAbort);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.txtLeft);
            this.Controls.Add(this.txtRight);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
            this.Name = "CorrectVerses";
            this.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.RightToLeftLayout = true;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "تصحیح مصاریع";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtRight;
        private System.Windows.Forms.TextBox txtLeft;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnAbort;
    }
}