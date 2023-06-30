namespace ClosedXML_Sample
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
            button1 = new Button();
            button2 = new Button();
            button3 = new Button();
            button4 = new Button();
            button5 = new Button();
            label1 = new Label();
            label2 = new Label();
            label3 = new Label();
            button6 = new Button();
            button7 = new Button();
            button8 = new Button();
            label4 = new Label();
            button9 = new Button();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Location = new Point(12, 27);
            button1.Name = "button1";
            button1.Size = new Size(1249, 23);
            button1.TabIndex = 0;
            button1.Text = "Hello World";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // button2
            // 
            button2.Location = new Point(12, 81);
            button2.Name = "button2";
            button2.Size = new Size(619, 23);
            button2.TabIndex = 1;
            button2.Text = "DataTable";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // button3
            // 
            button3.Location = new Point(642, 81);
            button3.Name = "button3";
            button3.Size = new Size(619, 23);
            button3.TabIndex = 2;
            button3.Text = "Objects or structs";
            button3.UseVisualStyleBackColor = true;
            button3.Click += button3_Click;
            // 
            // button4
            // 
            button4.Location = new Point(12, 110);
            button4.Name = "button4";
            button4.Size = new Size(619, 23);
            button4.TabIndex = 3;
            button4.Text = "IEnumerable";
            button4.UseVisualStyleBackColor = true;
            button4.Click += button4_Click;
            // 
            // button5
            // 
            button5.Location = new Point(642, 110);
            button5.Name = "button5";
            button5.Size = new Size(619, 23);
            button5.TabIndex = 4;
            button5.Text = "Untyped Enumerable";
            button5.UseVisualStyleBackColor = true;
            button5.Click += button5_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(12, 9);
            label1.Name = "label1";
            label1.Size = new Size(76, 15);
            label1.TabIndex = 5;
            label1.Text = "クイックスタート";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(12, 63);
            label2.Name = "label2";
            label2.Size = new Size(57, 15);
            label2.TabIndex = 6;
            label2.Text = "データ挿入";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(12, 147);
            label3.Name = "label3";
            label3.Size = new Size(43, 15);
            label3.TabIndex = 7;
            label3.Text = "テーブル";
            // 
            // button6
            // 
            button6.Location = new Point(13, 165);
            button6.Name = "button6";
            button6.Size = new Size(618, 23);
            button6.TabIndex = 8;
            button6.Text = "TableCreate";
            button6.UseVisualStyleBackColor = true;
            button6.Click += button6_Click;
            // 
            // button7
            // 
            button7.Location = new Point(642, 165);
            button7.Name = "button7";
            button7.Size = new Size(619, 23);
            button7.TabIndex = 9;
            button7.Text = "Table Theme";
            button7.UseVisualStyleBackColor = true;
            button7.Click += button7_Click;
            // 
            // button8
            // 
            button8.Location = new Point(12, 194);
            button8.Name = "button8";
            button8.RightToLeft = RightToLeft.Yes;
            button8.Size = new Size(619, 23);
            button8.TabIndex = 10;
            button8.Text = "Table style options";
            button8.UseVisualStyleBackColor = true;
            button8.Click += button8_Click;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(12, 231);
            label4.Name = "label4";
            label4.Size = new Size(40, 15);
            label4.TabIndex = 11;
            label4.Text = "フォント";
            // 
            // button9
            // 
            button9.Location = new Point(12, 249);
            button9.Name = "button9";
            button9.Size = new Size(619, 23);
            button9.TabIndex = 12;
            button9.Text = "Font";
            button9.UseVisualStyleBackColor = true;
            button9.Click += button9_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1273, 628);
            Controls.Add(button9);
            Controls.Add(label4);
            Controls.Add(button8);
            Controls.Add(button7);
            Controls.Add(button6);
            Controls.Add(label3);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(button5);
            Controls.Add(button4);
            Controls.Add(button3);
            Controls.Add(button2);
            Controls.Add(button1);
            Name = "Form1";
            Text = "Form1";
            Load += Form1_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button button1;
        private Button button2;
        private Button button3;
        private Button button4;
        private Button button5;
        private Label label1;
        private Label label2;
        private Label label3;
        private Button button6;
        private Button button7;
        private Button button8;
        private Label label4;
        private Button button9;
    }
}