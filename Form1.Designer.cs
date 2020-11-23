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
            this.BuildingListBox = new System.Windows.Forms.ListBox();
            this.BuildingTextBox = new System.Windows.Forms.TextBox();
            this.AddBuildingButton = new System.Windows.Forms.Button();
            this.SizeListBox = new System.Windows.Forms.ListBox();
            this.SizeTextBox = new System.Windows.Forms.TextBox();
            this.SizeButton = new System.Windows.Forms.Button();
            this.FileLabel = new System.Windows.Forms.Label();
            this.DeleteRoomButton = new System.Windows.Forms.Button();
            this.RemoveBUButton = new System.Windows.Forms.Button();
            this.RemoveBuildingButton = new System.Windows.Forms.Button();
            this.RemoveSizeButton = new System.Windows.Forms.Button();
            this.LocateMissingButton = new System.Windows.Forms.Button();
            this.MissingListLabel = new System.Windows.Forms.Label();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 4;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.Controls.Add(this.RoomList, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.RoomTextBox, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.BUTextBox, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.BUListBox, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.AddBUButton, 1, 2);
            this.tableLayoutPanel1.Controls.Add(this.AddRoomButton, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.BrowseButton, 0, 5);
            this.tableLayoutPanel1.Controls.Add(this.CreateButton, 1, 5);
            this.tableLayoutPanel1.Controls.Add(this.BuildingListBox, 2, 0);
            this.tableLayoutPanel1.Controls.Add(this.BuildingTextBox, 2, 1);
            this.tableLayoutPanel1.Controls.Add(this.AddBuildingButton, 2, 2);
            this.tableLayoutPanel1.Controls.Add(this.SizeListBox, 3, 0);
            this.tableLayoutPanel1.Controls.Add(this.SizeTextBox, 3, 1);
            this.tableLayoutPanel1.Controls.Add(this.SizeButton, 3, 2);
            this.tableLayoutPanel1.Controls.Add(this.FileLabel, 0, 4);
            this.tableLayoutPanel1.Controls.Add(this.DeleteRoomButton, 0, 3);
            this.tableLayoutPanel1.Controls.Add(this.RemoveBUButton, 1, 3);
            this.tableLayoutPanel1.Controls.Add(this.RemoveBuildingButton, 2, 3);
            this.tableLayoutPanel1.Controls.Add(this.RemoveSizeButton, 3, 3);
            this.tableLayoutPanel1.Controls.Add(this.LocateMissingButton, 2, 5);
            this.tableLayoutPanel1.Controls.Add(this.MissingListLabel, 2, 4);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 6;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 36F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 8F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(735, 436);
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
            this.RoomList.Size = new System.Drawing.Size(177, 256);
            this.RoomList.TabIndex = 0;
            // 
            // RoomTextBox
            // 
            this.RoomTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.RoomTextBox.Location = new System.Drawing.Point(3, 262);
            this.RoomTextBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.RoomTextBox.Name = "RoomTextBox";
            this.RoomTextBox.Size = new System.Drawing.Size(177, 26);
            this.RoomTextBox.TabIndex = 3;
            // 
            // BUTextBox
            // 
            this.BUTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.BUTextBox.Location = new System.Drawing.Point(186, 262);
            this.BUTextBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.BUTextBox.Name = "BUTextBox";
            this.BUTextBox.Size = new System.Drawing.Size(177, 26);
            this.BUTextBox.TabIndex = 4;
            // 
            // BUListBox
            // 
            this.BUListBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.BUListBox.FormattingEnabled = true;
            this.BUListBox.ItemHeight = 20;
            this.BUListBox.Location = new System.Drawing.Point(186, 2);
            this.BUListBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.BUListBox.Name = "BUListBox";
            this.BUListBox.Size = new System.Drawing.Size(177, 256);
            this.BUListBox.TabIndex = 6;
            // 
            // AddBUButton
            // 
            this.AddBUButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.AddBUButton.Location = new System.Drawing.Point(186, 294);
            this.AddBUButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.AddBUButton.Name = "AddBUButton";
            this.AddBUButton.Size = new System.Drawing.Size(177, 27);
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
            this.AddRoomButton.Size = new System.Drawing.Size(177, 27);
            this.AddRoomButton.TabIndex = 2;
            this.AddRoomButton.Text = "Add Room";
            this.AddRoomButton.UseVisualStyleBackColor = true;
            this.AddRoomButton.Click += new System.EventHandler(this.AddRoomButton_Click);
            // 
            // BrowseButton
            // 
            this.BrowseButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.BrowseButton.Location = new System.Drawing.Point(3, 396);
            this.BrowseButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.BrowseButton.Name = "BrowseButton";
            this.BrowseButton.Size = new System.Drawing.Size(177, 38);
            this.BrowseButton.TabIndex = 7;
            this.BrowseButton.Text = "Browse";
            this.BrowseButton.UseVisualStyleBackColor = true;
            this.BrowseButton.Click += new System.EventHandler(this.BrowseButton_Click);
            // 
            // CreateButton
            // 
            this.CreateButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.CreateButton.Location = new System.Drawing.Point(186, 396);
            this.CreateButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.CreateButton.Name = "CreateButton";
            this.CreateButton.Size = new System.Drawing.Size(177, 38);
            this.CreateButton.TabIndex = 8;
            this.CreateButton.Text = "Create";
            this.CreateButton.UseVisualStyleBackColor = true;
            this.CreateButton.Click += new System.EventHandler(this.CreateButton_Click);
            // 
            // BuildingListBox
            // 
            this.BuildingListBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.BuildingListBox.FormattingEnabled = true;
            this.BuildingListBox.ItemHeight = 20;
            this.BuildingListBox.Location = new System.Drawing.Point(369, 3);
            this.BuildingListBox.Name = "BuildingListBox";
            this.BuildingListBox.Size = new System.Drawing.Size(177, 254);
            this.BuildingListBox.TabIndex = 10;
            // 
            // BuildingTextBox
            // 
            this.BuildingTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.BuildingTextBox.Location = new System.Drawing.Point(369, 263);
            this.BuildingTextBox.Name = "BuildingTextBox";
            this.BuildingTextBox.Size = new System.Drawing.Size(177, 26);
            this.BuildingTextBox.TabIndex = 11;
            // 
            // AddBuildingButton
            // 
            this.AddBuildingButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.AddBuildingButton.Location = new System.Drawing.Point(369, 295);
            this.AddBuildingButton.Name = "AddBuildingButton";
            this.AddBuildingButton.Size = new System.Drawing.Size(177, 25);
            this.AddBuildingButton.TabIndex = 12;
            this.AddBuildingButton.Text = "Add Building";
            this.AddBuildingButton.UseVisualStyleBackColor = true;
            this.AddBuildingButton.Click += new System.EventHandler(this.AddBuildingButton_Click);
            // 
            // SizeListBox
            // 
            this.SizeListBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.SizeListBox.FormattingEnabled = true;
            this.SizeListBox.ItemHeight = 20;
            this.SizeListBox.Location = new System.Drawing.Point(552, 3);
            this.SizeListBox.Name = "SizeListBox";
            this.SizeListBox.Size = new System.Drawing.Size(180, 254);
            this.SizeListBox.TabIndex = 13;
            // 
            // SizeTextBox
            // 
            this.SizeTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.SizeTextBox.Location = new System.Drawing.Point(552, 263);
            this.SizeTextBox.Name = "SizeTextBox";
            this.SizeTextBox.Size = new System.Drawing.Size(180, 26);
            this.SizeTextBox.TabIndex = 14;
            // 
            // SizeButton
            // 
            this.SizeButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.SizeButton.Location = new System.Drawing.Point(552, 295);
            this.SizeButton.Name = "SizeButton";
            this.SizeButton.Size = new System.Drawing.Size(180, 25);
            this.SizeButton.TabIndex = 15;
            this.SizeButton.Text = "Add Size";
            this.SizeButton.UseVisualStyleBackColor = true;
            this.SizeButton.Click += new System.EventHandler(this.SizeButton_Click);
            // 
            // FileLabel
            // 
            this.FileLabel.AutoSize = true;
            this.tableLayoutPanel1.SetColumnSpan(this.FileLabel, 2);
            this.FileLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.FileLabel.Location = new System.Drawing.Point(3, 358);
            this.FileLabel.Name = "FileLabel";
            this.FileLabel.Size = new System.Drawing.Size(360, 36);
            this.FileLabel.TabIndex = 9;
            this.FileLabel.Text = "New File Name";
            // 
            // DeleteRoomButton
            // 
            this.DeleteRoomButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DeleteRoomButton.Location = new System.Drawing.Point(3, 326);
            this.DeleteRoomButton.Name = "DeleteRoomButton";
            this.DeleteRoomButton.Size = new System.Drawing.Size(177, 29);
            this.DeleteRoomButton.TabIndex = 16;
            this.DeleteRoomButton.Text = "Remove Room";
            this.DeleteRoomButton.UseVisualStyleBackColor = true;
            this.DeleteRoomButton.Click += new System.EventHandler(this.DeleteRoomButton_Click);
            // 
            // RemoveBUButton
            // 
            this.RemoveBUButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.RemoveBUButton.Location = new System.Drawing.Point(186, 326);
            this.RemoveBUButton.Name = "RemoveBUButton";
            this.RemoveBUButton.Size = new System.Drawing.Size(177, 29);
            this.RemoveBUButton.TabIndex = 17;
            this.RemoveBUButton.Text = "Remove BU";
            this.RemoveBUButton.UseVisualStyleBackColor = true;
            this.RemoveBUButton.Click += new System.EventHandler(this.RemoveBUButton_Click);
            // 
            // RemoveBuildingButton
            // 
            this.RemoveBuildingButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.RemoveBuildingButton.Location = new System.Drawing.Point(369, 326);
            this.RemoveBuildingButton.Name = "RemoveBuildingButton";
            this.RemoveBuildingButton.Size = new System.Drawing.Size(177, 29);
            this.RemoveBuildingButton.TabIndex = 18;
            this.RemoveBuildingButton.Text = "Remove Building";
            this.RemoveBuildingButton.UseVisualStyleBackColor = true;
            this.RemoveBuildingButton.Click += new System.EventHandler(this.RemoveBuildingButton_Click);
            // 
            // RemoveSizeButton
            // 
            this.RemoveSizeButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.RemoveSizeButton.Location = new System.Drawing.Point(552, 326);
            this.RemoveSizeButton.Name = "RemoveSizeButton";
            this.RemoveSizeButton.Size = new System.Drawing.Size(180, 29);
            this.RemoveSizeButton.TabIndex = 19;
            this.RemoveSizeButton.Text = "Remove Size";
            this.RemoveSizeButton.UseVisualStyleBackColor = true;
            this.RemoveSizeButton.Click += new System.EventHandler(this.RemoveSizeButton_Click);
            // 
            // LocateMissingButton
            // 
            this.LocateMissingButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.LocateMissingButton.Location = new System.Drawing.Point(369, 397);
            this.LocateMissingButton.Name = "LocateMissingButton";
            this.LocateMissingButton.Size = new System.Drawing.Size(177, 36);
            this.LocateMissingButton.TabIndex = 20;
            this.LocateMissingButton.Text = "Local Missing List";
            this.LocateMissingButton.UseVisualStyleBackColor = true;
            this.LocateMissingButton.Click += new System.EventHandler(this.LocateMissingButton_Click);
            // 
            // MissingListLabel
            // 
            this.MissingListLabel.AutoSize = true;
            this.MissingListLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.MissingListLabel.Location = new System.Drawing.Point(369, 358);
            this.MissingListLabel.Name = "MissingListLabel";
            this.MissingListLabel.Size = new System.Drawing.Size(177, 36);
            this.MissingListLabel.TabIndex = 21;
            this.MissingListLabel.Text = "Missing List Directory";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(735, 436);
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
        private System.Windows.Forms.Button AddBuildingButton;
        private System.Windows.Forms.ListBox SizeListBox;
        private System.Windows.Forms.TextBox SizeTextBox;
        private System.Windows.Forms.Button SizeButton;
        private System.Windows.Forms.Button DeleteRoomButton;
        private System.Windows.Forms.Button RemoveBUButton;
        private System.Windows.Forms.Button RemoveBuildingButton;
        private System.Windows.Forms.Button RemoveSizeButton;
        private System.Windows.Forms.Button LocateMissingButton;
        private System.Windows.Forms.Label MissingListLabel;
    }
}

