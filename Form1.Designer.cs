namespace BUPicksList
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
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.RoomList = new System.Windows.Forms.ListBox();
            this.RoomTextBox = new System.Windows.Forms.TextBox();
            this.BUTextBox = new System.Windows.Forms.TextBox();
            this.BUListBox = new System.Windows.Forms.ListBox();
            this.AddBUButton = new System.Windows.Forms.Button();
            this.AddRoomButton = new System.Windows.Forms.Button();
            this.BrowseButton = new System.Windows.Forms.Button();
            this.CreateButton = new System.Windows.Forms.Button();
            this.FileLabel = new System.Windows.Forms.Label();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.RoomList, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.RoomTextBox, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.BUTextBox, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.BUListBox, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.AddBUButton, 1, 2);
            this.tableLayoutPanel1.Controls.Add(this.AddRoomButton, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.BrowseButton, 0, 4);
            this.tableLayoutPanel1.Controls.Add(this.CreateButton, 1, 4);
            this.tableLayoutPanel1.Controls.Add(this.FileLabel, 0, 3);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 5;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 44F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 8F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(451, 490);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // RoomList
            // 
            this.RoomList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.RoomList.FormattingEnabled = true;
            this.RoomList.ItemHeight = 25;
            this.RoomList.Location = new System.Drawing.Point(3, 3);
            this.RoomList.Name = "RoomList";
            this.RoomList.Size = new System.Drawing.Size(219, 319);
            this.RoomList.TabIndex = 0;
            // 
            // RoomTextBox
            // 
            this.RoomTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.RoomTextBox.Location = new System.Drawing.Point(3, 328);
            this.RoomTextBox.Name = "RoomTextBox";
            this.RoomTextBox.Size = new System.Drawing.Size(219, 31);
            this.RoomTextBox.TabIndex = 3;
            // 
            // BUTextBox
            // 
            this.BUTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.BUTextBox.Location = new System.Drawing.Point(228, 328);
            this.BUTextBox.Name = "BUTextBox";
            this.BUTextBox.Size = new System.Drawing.Size(220, 31);
            this.BUTextBox.TabIndex = 4;
            // 
            // BUListBox
            // 
            this.BUListBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.BUListBox.FormattingEnabled = true;
            this.BUListBox.ItemHeight = 25;
            this.BUListBox.Location = new System.Drawing.Point(228, 3);
            this.BUListBox.Name = "BUListBox";
            this.BUListBox.Size = new System.Drawing.Size(220, 319);
            this.BUListBox.TabIndex = 6;
            // 
            // AddBUButton
            // 
            this.AddBUButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.AddBUButton.Location = new System.Drawing.Point(228, 365);
            this.AddBUButton.Name = "AddBUButton";
            this.AddBUButton.Size = new System.Drawing.Size(220, 34);
            this.AddBUButton.TabIndex = 5;
            this.AddBUButton.Text = "Add BU";
            this.AddBUButton.UseVisualStyleBackColor = true;
            this.AddBUButton.Click += new System.EventHandler(this.AddBUButton_Click);
            // 
            // AddRoomButton
            // 
            this.AddRoomButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.AddRoomButton.Location = new System.Drawing.Point(3, 365);
            this.AddRoomButton.Name = "AddRoomButton";
            this.AddRoomButton.Size = new System.Drawing.Size(219, 34);
            this.AddRoomButton.TabIndex = 2;
            this.AddRoomButton.Text = "Add Room";
            this.AddRoomButton.UseVisualStyleBackColor = true;
            this.AddRoomButton.Click += new System.EventHandler(this.AddRoomButton_Click);
            // 
            // BrowseButton
            // 
            this.BrowseButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.BrowseButton.Location = new System.Drawing.Point(3, 449);
            this.BrowseButton.Name = "BrowseButton";
            this.BrowseButton.Size = new System.Drawing.Size(219, 38);
            this.BrowseButton.TabIndex = 7;
            this.BrowseButton.Text = "Browse";
            this.BrowseButton.UseVisualStyleBackColor = true;
            this.BrowseButton.Click += new System.EventHandler(this.BrowseButton_Click);
            // 
            // CreateButton
            // 
            this.CreateButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.CreateButton.Location = new System.Drawing.Point(228, 449);
            this.CreateButton.Name = "CreateButton";
            this.CreateButton.Size = new System.Drawing.Size(220, 38);
            this.CreateButton.TabIndex = 8;
            this.CreateButton.Text = "Create";
            this.CreateButton.UseVisualStyleBackColor = true;
            this.CreateButton.Click += new System.EventHandler(this.CreateButton_Click);
            // 
            // FileLabel
            // 
            this.FileLabel.AutoSize = true;
            this.tableLayoutPanel1.SetColumnSpan(this.FileLabel, 2);
            this.FileLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.FileLabel.Location = new System.Drawing.Point(3, 402);
            this.FileLabel.Name = "FileLabel";
            this.FileLabel.Size = new System.Drawing.Size(445, 44);
            this.FileLabel.TabIndex = 9;
            this.FileLabel.Text = "New File Name";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(451, 490);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "Form1";
            this.Text = "BUPickList";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.ListBox RoomList;
        private System.Windows.Forms.Button AddRoomButton;
        private System.Windows.Forms.TextBox RoomTextBox;
        private System.Windows.Forms.TextBox BUTextBox;
        private System.Windows.Forms.Button AddBUButton;
        private System.Windows.Forms.ListBox BUListBox;
        private System.Windows.Forms.Button BrowseButton;
        private System.Windows.Forms.Button CreateButton;
        private System.Windows.Forms.Label FileLabel;
    }
}

