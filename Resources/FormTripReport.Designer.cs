namespace SemtechAssistant.Resources
{
    partial class FormTripReport
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormTripReport));
            this.labelDate = new System.Windows.Forms.Label();
            this.labelCustomer = new System.Windows.Forms.Label();
            this.labelAttendee = new System.Windows.Forms.Label();
            this.textBoxDate = new System.Windows.Forms.TextBox();
            this.textBoxCustomer = new System.Windows.Forms.TextBox();
            this.textBoxAttendee = new System.Windows.Forms.TextBox();
            this.buttonCreate = new System.Windows.Forms.Button();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // labelDate
            // 
            this.labelDate.AutoSize = true;
            this.labelDate.Location = new System.Drawing.Point(12, 15);
            this.labelDate.Name = "labelDate";
            this.labelDate.Size = new System.Drawing.Size(30, 13);
            this.labelDate.TabIndex = 0;
            this.labelDate.Text = "Date";
            // 
            // labelCustomer
            // 
            this.labelCustomer.AutoSize = true;
            this.labelCustomer.Location = new System.Drawing.Point(12, 41);
            this.labelCustomer.Name = "labelCustomer";
            this.labelCustomer.Size = new System.Drawing.Size(51, 13);
            this.labelCustomer.TabIndex = 1;
            this.labelCustomer.Text = "Customer";
            // 
            // labelAttendee
            // 
            this.labelAttendee.AutoSize = true;
            this.labelAttendee.Location = new System.Drawing.Point(11, 67);
            this.labelAttendee.Name = "labelAttendee";
            this.labelAttendee.Size = new System.Drawing.Size(50, 13);
            this.labelAttendee.TabIndex = 2;
            this.labelAttendee.Text = "Attendee";
            // 
            // textBoxDate
            // 
            this.textBoxDate.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxDate.Location = new System.Drawing.Point(67, 12);
            this.textBoxDate.Name = "textBoxDate";
            this.textBoxDate.Size = new System.Drawing.Size(200, 20);
            this.textBoxDate.TabIndex = 3;
            this.textBoxDate.Text = "2013.9.10";
            // 
            // textBoxCustomer
            // 
            this.textBoxCustomer.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxCustomer.Location = new System.Drawing.Point(67, 38);
            this.textBoxCustomer.Name = "textBoxCustomer";
            this.textBoxCustomer.Size = new System.Drawing.Size(200, 20);
            this.textBoxCustomer.TabIndex = 4;
            this.textBoxCustomer.Text = "VS-I";
            // 
            // textBoxAttendee
            // 
            this.textBoxAttendee.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxAttendee.Location = new System.Drawing.Point(67, 64);
            this.textBoxAttendee.Name = "textBoxAttendee";
            this.textBoxAttendee.Size = new System.Drawing.Size(200, 20);
            this.textBoxAttendee.TabIndex = 5;
            this.textBoxAttendee.Text = "maliyu";
            // 
            // buttonCreate
            // 
            this.buttonCreate.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.buttonCreate.Image = ((System.Drawing.Image)(resources.GetObject("buttonCreate.Image")));
            this.buttonCreate.Location = new System.Drawing.Point(115, 147);
            this.buttonCreate.Name = "buttonCreate";
            this.buttonCreate.Size = new System.Drawing.Size(70, 70);
            this.buttonCreate.TabIndex = 6;
            this.buttonCreate.UseVisualStyleBackColor = true;
            this.buttonCreate.Click += new System.EventHandler(this.buttonCreate_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.pictureBox1.InitialImage = null;
            this.pictureBox1.Location = new System.Drawing.Point(115, 147);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(70, 70);
            this.pictureBox1.TabIndex = 7;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Visible = false;
            // 
            // FormTripReport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(292, 262);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.buttonCreate);
            this.Controls.Add(this.textBoxAttendee);
            this.Controls.Add(this.textBoxCustomer);
            this.Controls.Add(this.textBoxDate);
            this.Controls.Add(this.labelAttendee);
            this.Controls.Add(this.labelCustomer);
            this.Controls.Add(this.labelDate);
            this.Name = "FormTripReport";
            this.Text = "Generate Form Trip Report";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FormTripReport_FormClosed);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelDate;
        private System.Windows.Forms.Label labelCustomer;
        private System.Windows.Forms.Label labelAttendee;
        private System.Windows.Forms.TextBox textBoxDate;
        private System.Windows.Forms.TextBox textBoxCustomer;
        private System.Windows.Forms.TextBox textBoxAttendee;
        private System.Windows.Forms.Button buttonCreate;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}