namespace LiXcel
{
    partial class TaskPaneView
    {
        /// <summary> 
        /// Variabile di progettazione necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Liberare le risorse in uso.
        /// </summary>
        /// <param name="disposing">ha valore true se le risorse gestite devono essere eliminate, false in caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Codice generato da Progettazione componenti

        /// <summary> 
        /// Metodo necessario per il supporto della finestra di progettazione. Non modificare 
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InitializeComponent()
        {
            this.StartButton = new System.Windows.Forms.Button();
            this.StatusLabel = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.iterationsNumericUpDown = new System.Windows.Forms.NumericUpDown();
            ((System.ComponentModel.ISupportInitialize)(this.iterationsNumericUpDown)).BeginInit();
            this.SuspendLayout();
            // 
            // StartButton
            // 
            this.StartButton.Location = new System.Drawing.Point(4, 4);
            this.StartButton.Name = "StartButton";
            this.StartButton.Size = new System.Drawing.Size(164, 23);
            this.StartButton.TabIndex = 0;
            this.StartButton.Text = "Start simulation";
            this.StartButton.UseVisualStyleBackColor = true;
            this.StartButton.Click += new System.EventHandler(this.button1_Click);
            // 
            // StatusLabel
            // 
            this.StatusLabel.AutoSize = true;
            this.StatusLabel.Location = new System.Drawing.Point(3, 59);
            this.StatusLabel.Name = "StatusLabel";
            this.StatusLabel.Size = new System.Drawing.Size(128, 13);
            this.StatusLabel.TabIndex = 1;
            this.StatusLabel.Text = "LiXcel simulation statistics";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 34);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "iteration size";
            // 
            // iterationsNumericUpDown
            // 
            this.iterationsNumericUpDown.Increment = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.iterationsNumericUpDown.Location = new System.Drawing.Point(78, 34);
            this.iterationsNumericUpDown.Maximum = new decimal(new int[] {
            100000000,
            0,
            0,
            0});
            this.iterationsNumericUpDown.Minimum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.iterationsNumericUpDown.Name = "iterationsNumericUpDown";
            this.iterationsNumericUpDown.Size = new System.Drawing.Size(90, 20);
            this.iterationsNumericUpDown.TabIndex = 3;
            this.iterationsNumericUpDown.Value = new decimal(new int[] {
            100000,
            0,
            0,
            0});
            // 
            // TaskPaneView
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.iterationsNumericUpDown);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.StatusLabel);
            this.Controls.Add(this.StartButton);
            this.Name = "TaskPaneView";
            this.Size = new System.Drawing.Size(171, 286);
            ((System.ComponentModel.ISupportInitialize)(this.iterationsNumericUpDown)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button StartButton;
        private System.Windows.Forms.Label StatusLabel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown iterationsNumericUpDown;
    }
}
