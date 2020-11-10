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
            this.BuildingListBox = new System.Windows.Forms.ListBox();
            this.BuildingTextBox = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 3;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel1.Controls.Add(this.RoomList, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.RoomTextBox, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.BUTextBox, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.BUListBox, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.AddBUButton, 1, 2);
            this.tableLayoutPanel1.Controls.Add(this.AddRoomButton, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.BrowseButton, 0, 4);
            this.tableLayoutPanel1.Controls.Add(this.CreateButton, 1, 4);
            this.tableLayoutPanel1.Controls.Add(this.FileLabel, 0, 3);
            this.tableLayoutPanel1.Controls.Add(this.BuildingListBox, 2, 0);
            this.tableLayoutPanel1.Controls.Add(this.BuildingTextBox, 2, 1);
            this.tableLayoutPanel1.Controls.Add(this.button1, 2, 2);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 5;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 6F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(735, 392);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // RoomList
            // 
            this.RoomList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.RoomList.FormattingEnabled = true;
            this.RoomList.ItemHeight = 20;
            this.RoomList.Location = new System.Drawing.Point(3, 2);
            this.RoomList.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.RoomList.Name = "RoomList";
            this.RoomList.Size = new System.Drawing.Size(238, 256);
            this.RoomList.TabIndex = 0;
            // 
            // RoomTextBox
            // 
            this.RoomTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.RoomTextBox.Location = new System.Drawing.Point(3, 262);
            this.RoomTextBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.RoomTextBox.Name = "RoomTextBox";
            this.RoomTextBox.Size = new System.Drawing.Size(238, 26);
            this.RoomTextBox.TabIndex = 3;
            // 
            // BUTextBox
            // 
            this.BUTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.BUTextBox.Location = new System.Drawing.Point(247, 262);
            this.BUTextBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.BUTextBox.Name = "BUTextBox";
            this.BUTextBox.Size = new System.Drawing.Size(238, 26);
            this.BUTextBox.TabIndex = 4;
            // 
            // BUListBox
            // 
            this.BUListBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.BUListBox.FormattingEnabled = true;
            this.BUListBox.ItemHeight = 20;
            this.BUListBox.Location = new System.Drawing.Point(247, 2);
            this.BUListBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.BUListBox.Name = "BUListBox";
            this.BUListBox.Size = new System.Drawing.Size(238, 256);
            this.BUListBox.TabIndex = 6;
            // 
            // AddBUButton
            // 
            this.AddBUButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.AddBUButton.Location = new System.Drawing.Point(247, 294);
            this.AddBUButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.AddBUButton.Name = "AddBUButton";
            this.AddBUButton.Size = new System.Drawing.Size(238, 27);
            this.AddBUButton.TabIndex = 5;
            this.AddBUButton.Text = "Add BU";
            this.AddBUButton.UseVisualStyleBackColor = true;
            this.AddBUButton.Click += new System.EventHandler(this.AddBUButton_Click);
            // 
            // AddRoomButton
            // 
            this.AddRoomButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.AddRoomButton.Location = new System.Drawing.Point(3, 294);
            this.AddRoomButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.AddRoomButton.Name = "AddRoomButton";
            this.AddRoomButton.Size = new System.Drawing.Size(238, 27);
            this.AddRoomButton.TabIndex = 2;
            this.AddRoomButton.Text = "Add Room";
            this.AddRoomButton.UseVisualStyleBackColor = true;
            this.AddRoomButton.Click += new System.EventHandler(this.AddRoomButton_Click);
            // 
            // BrowseButton
            // 
            this.BrowseButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.BrowseButton.Location = new System.Drawing.Point(3, 360);
            this.BrowseButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.BrowseButton.Name = "BrowseButton";
            this.BrowseButton.Size = new System.Drawing.Size(238, 30);
            this.BrowseButton.TabIndex = 7;
            this.BrowseButton.Text = "Browse";
            this.BrowseButton.UseVisualStyleBackColor = true;
            this.BrowseButton.Click += new System.EventHandler(this.BrowseButton_Click);
            // 
            // CreateButton
            // 
            this.CreateButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.CreateButton.Location = new System.Drawing.Point(247, 360);
            this.CreateButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.CreateButton.Name = "CreateButton";
            this.CreateButton.Size = new System.Drawing.Size(238, 30);
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
            this.FileLabel.Location = new System.Drawing.Point(3, 323);
            this.FileLabel.Name = "FileLabel";
            this.FileLabel.Size = new System.Drawing.Size(482, 35);
            this.FileLabel.TabIndex = 9;
            this.FileLabel.Text = "New File Name";
            // 
            // BuildingListBox
            // 
            this.BuildingListBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.BuildingListBox.FormattingEnabled = true;
            this.BuildingListBox.ItemHeight = 20;
            this.BuildingListBox.Location = new System.Drawing.Point(491, 3);
            this.BuildingListBox.Name = "BuildingListBox";
            this.BuildingListBox.Size = new System.Drawing.Size(241, 254);
            this.BuildingListBox.TabIndex = 10;
            // 
            // BuildingTextBox
            // 
            this.BuildingTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.BuildingTextBox.Location = new System.Drawing.Point(491, 263);
            this.BuildingTextBox.Name = "BuildingTextBox";
            this.BuildingTextBox.Size = new System.Drawing.Size(241, 26);
            this.BuildingTextBox.TabIndex = 11;
            // 
            // button1
            // 
            this.button1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button1.Location = new System.Drawing.Point(491, 295);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(241, 25);
            this.button1.TabIndex = 12;
            this.button1.Text = "Add Building";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(735, 392);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
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
        private System.Windows.Forms.ListBox BuildingListBox;
        private System.Windows.Forms.TextBox BuildingTextBox;
        private System.Windows.Forms.Button button1;
    }
}

