namespace Korektor
{
    partial class FormCorrectWord
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
            this.lstSuggestions = new System.Windows.Forms.ListBox();
            this.btnIgnore = new System.Windows.Forms.Button();
            this.btnIgnoreAll = new System.Windows.Forms.Button();
            this.btnAddToDictionary = new System.Windows.Forms.Button();
            this.btnChange = new System.Windows.Forms.Button();
            this.btnChangeAll = new System.Windows.Forms.Button();
            this.tbxReplace = new System.Windows.Forms.TextBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.SuspendLayout();
            // 
            // lstSuggestions
            // 
            this.lstSuggestions.FormattingEnabled = true;
            this.lstSuggestions.ItemHeight = 15;
            this.lstSuggestions.Location = new System.Drawing.Point(12, 85);
            this.lstSuggestions.Name = "lstSuggestions";
            this.lstSuggestions.Size = new System.Drawing.Size(331, 79);
            this.lstSuggestions.TabIndex = 1;
            this.lstSuggestions.SelectedIndexChanged += new System.EventHandler(this.lstSuggestions_SelectedIndexChanged);
            // 
            // btnIgnore
            // 
            this.btnIgnore.Location = new System.Drawing.Point(12, 47);
            this.btnIgnore.Name = "btnIgnore";
            this.btnIgnore.Size = new System.Drawing.Size(104, 27);
            this.btnIgnore.TabIndex = 0;
            this.btnIgnore.Text = "Прескочи";
            this.btnIgnore.UseVisualStyleBackColor = true;
            this.btnIgnore.Click += new System.EventHandler(this.btnIgnore_Click);
            // 
            // btnIgnoreAll
            // 
            this.btnIgnoreAll.Location = new System.Drawing.Point(122, 47);
            this.btnIgnoreAll.Name = "btnIgnoreAll";
            this.btnIgnoreAll.Size = new System.Drawing.Size(104, 27);
            this.btnIgnoreAll.TabIndex = 3;
            this.btnIgnoreAll.Text = "Прескочи све";
            this.btnIgnoreAll.UseVisualStyleBackColor = true;
            this.btnIgnoreAll.Click += new System.EventHandler(this.btnIgnoreAll_Click);
            // 
            // btnAddToDictionary
            // 
            this.btnAddToDictionary.Location = new System.Drawing.Point(233, 47);
            this.btnAddToDictionary.Name = "btnAddToDictionary";
            this.btnAddToDictionary.Size = new System.Drawing.Size(110, 27);
            this.btnAddToDictionary.TabIndex = 2;
            this.btnAddToDictionary.Text = "Додај у речник";
            this.btnAddToDictionary.UseVisualStyleBackColor = true;
            this.btnAddToDictionary.Click += new System.EventHandler(this.btnAddToDictionary_Click);
            // 
            // btnChange
            // 
            this.btnChange.Location = new System.Drawing.Point(12, 175);
            this.btnChange.Name = "btnChange";
            this.btnChange.Size = new System.Drawing.Size(104, 27);
            this.btnChange.TabIndex = 6;
            this.btnChange.Text = "Замени";
            this.btnChange.UseVisualStyleBackColor = true;
            this.btnChange.Click += new System.EventHandler(this.btnChange_Click);
            // 
            // btnChangeAll
            // 
            this.btnChangeAll.Location = new System.Drawing.Point(122, 175);
            this.btnChangeAll.Name = "btnChangeAll";
            this.btnChangeAll.Size = new System.Drawing.Size(104, 27);
            this.btnChangeAll.TabIndex = 7;
            this.btnChangeAll.Text = "Замени све";
            this.btnChangeAll.UseVisualStyleBackColor = true;
            this.btnChangeAll.Click += new System.EventHandler(this.btnChangeAll_Click);
            // 
            // tbxReplace
            // 
            this.tbxReplace.Location = new System.Drawing.Point(12, 14);
            this.tbxReplace.Name = "tbxReplace";
            this.tbxReplace.Size = new System.Drawing.Size(331, 21);
            this.tbxReplace.TabIndex = 8;
            this.tbxReplace.TextChanged += new System.EventHandler(this.tbxReplace_TextChanged);
            // 
            // FormCorrectWord
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(356, 211);
            this.Controls.Add(this.tbxReplace);
            this.Controls.Add(this.btnChangeAll);
            this.Controls.Add(this.btnChange);
            this.Controls.Add(this.btnIgnoreAll);
            this.Controls.Add(this.btnAddToDictionary);
            this.Controls.Add(this.lstSuggestions);
            this.Controls.Add(this.btnIgnore);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormCorrectWord";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Исправи Реч";
            this.Activated += new System.EventHandler(this.FormCorrectWord_Activated);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FormCorrectWord_FormClosed);
            this.Load += new System.EventHandler(this.FormCorrectWord_Load);
            this.VisibleChanged += new System.EventHandler(this.FormCorrectWord_VisibleChanged);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnIgnore;
        private System.Windows.Forms.ListBox lstSuggestions;
        private System.Windows.Forms.Button btnAddToDictionary;
        private System.Windows.Forms.Button btnIgnoreAll;
        private System.Windows.Forms.Button btnChange;
        private System.Windows.Forms.Button btnChangeAll;
        private System.Windows.Forms.TextBox tbxReplace;
        private System.Windows.Forms.ToolTip toolTip1;
    }
}