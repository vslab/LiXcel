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
            this.minTextBox = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.maxTextBox = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.maxIterationsNumericUpDown = new System.Windows.Forms.NumericUpDown();
            this.label4 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.iterationsNumericUpDown)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.maxIterationsNumericUpDown)).BeginInit();
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
            this.StatusLabel.Location = new System.Drawing.Point(6, 157);
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
            this.label1.Size = new System.Drawing.Size(68, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Refresh after";
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
            // minTextBox
            // 
            this.minTextBox.Location = new System.Drawing.Point(69, 17);
            this.minTextBox.Name = "minTextBox";
            this.minTextBox.Size = new System.Drawing.Size(77, 20);
            this.minTextBox.TabIndex = 4;
            this.minTextBox.Text = "auto";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.maxTextBox);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.minTextBox);
            this.groupBox1.Location = new System.Drawing.Point(9, 82);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(152, 72);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "range";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(7, 46);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(26, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "max";
            // 
            // maxTextBox
            // 
            this.maxTextBox.Location = new System.Drawing.Point(69, 43);
            this.maxTextBox.Name = "maxTextBox";
            this.maxTextBox.Size = new System.Drawing.Size(77, 20);
            this.maxTextBox.TabIndex = 6;
            this.maxTextBox.Text = "auto";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(7, 20);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(23, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "min";
            // 
            // maxIterationsNumericUpDown
            // 
            this.maxIterationsNumericUpDown.Increment = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.maxIterationsNumericUpDown.Location = new System.Drawing.Point(78, 60);
            this.maxIterationsNumericUpDown.Maximum = new decimal(new int[] {
            100000000,
            0,
            0,
            0});
            this.maxIterationsNumericUpDown.Minimum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.maxIterationsNumericUpDown.Name = "maxIterationsNumericUpDown";
            this.maxIterationsNumericUpDown.Size = new System.Drawing.Size(90, 20);
            this.maxIterationsNumericUpDown.TabIndex = 7;
            this.maxIterationsNumericUpDown.Value = new decimal(new int[] {
            10000000,
            0,
            0,
            0});
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 60);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(72, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "Max iterations";
            // 
            // TaskPaneView
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.maxIterationsNumericUpDown);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.iterationsNumericUpDown);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.StatusLabel);
            this.Controls.Add(this.StartButton);
            this.Name = "TaskPaneView";
            this.Size = new System.Drawing.Size(171, 286);
            ((System.ComponentModel.ISupportInitialize)(this.iterationsNumericUpDown)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.maxIterationsNumericUpDown)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button StartButton;
        private System.Windows.Forms.Label StatusLabel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown iterationsNumericUpDown;
        private System.Windows.Forms.TextBox minTextBox;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox maxTextBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.NumericUpDown maxIterationsNumericUpDown;
        private System.Windows.Forms.Label label4;
    }
}
