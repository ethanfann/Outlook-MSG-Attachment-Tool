using System.ComponentModel;
using System.Windows.Forms;

namespace Outlook_MSG_Attachment_Tool
{
    sealed partial class DragForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private IContainer components = null;

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DragForm));
            this.detailLabel = new System.Windows.Forms.Label();
            this.dragLabel = new System.Windows.Forms.Label();
            this.settingsButton = new System.Windows.Forms.Button();
            this.settingsPanel = new System.Windows.Forms.Panel();
            this.deepSearchBox = new System.Windows.Forms.CheckBox();
            this.browseButton = new System.Windows.Forms.Button();
            this.locationText = new System.Windows.Forms.TextBox();
            this.locationBox = new System.Windows.Forms.CheckBox();
            this.settingsTooltip = new System.Windows.Forms.ToolTip(this.components);
            this.moveMsgTooltip = new System.Windows.Forms.ToolTip(this.components);
            this.folderBrowser = new System.Windows.Forms.FolderBrowserDialog();
            this.deepSearchTip = new System.Windows.Forms.ToolTip(this.components);
            this.locationBoxTip = new System.Windows.Forms.ToolTip(this.components);
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.settingsPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // detailLabel
            // 
            this.detailLabel.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.detailLabel.ForeColor = System.Drawing.SystemColors.ButtonShadow;
            this.detailLabel.Location = new System.Drawing.Point(0, 208);
            this.detailLabel.Name = "detailLabel";
            this.detailLabel.Size = new System.Drawing.Size(284, 51);
            this.detailLabel.TabIndex = 1;
            this.detailLabel.Text = "placeholder";
            this.detailLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.detailLabel.Visible = false;
            // 
            // dragLabel
            // 
            this.dragLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dragLabel.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dragLabel.ForeColor = System.Drawing.SystemColors.ButtonShadow;
            this.dragLabel.Location = new System.Drawing.Point(0, 0);
            this.dragLabel.Name = "dragLabel";
            this.dragLabel.Size = new System.Drawing.Size(284, 261);
            this.dragLabel.TabIndex = 0;
            this.dragLabel.Text = "Drag and Drop .msg file(s)";
            this.dragLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // settingsButton
            // 
            this.settingsButton.BackColor = System.Drawing.Color.Transparent;
            this.settingsButton.BackgroundImage = global::Outlook_MSG_Attachment_Tool.Properties.Resources.settings2;
            this.settingsButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.settingsButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.settingsButton.FlatAppearance.BorderSize = 0;
            this.settingsButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.settingsButton.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.settingsButton.Location = new System.Drawing.Point(4, 0);
            this.settingsButton.Name = "settingsButton";
            this.settingsButton.Size = new System.Drawing.Size(28, 24);
            this.settingsButton.TabIndex = 2;
            this.settingsTooltip.SetToolTip(this.settingsButton, "Show/Hide settings");
            this.settingsButton.UseVisualStyleBackColor = false;
            this.settingsButton.Click += new System.EventHandler(this.settingsButton_Click);
            // 
            // settingsPanel
            // 
            this.settingsPanel.Controls.Add(this.deepSearchBox);
            this.settingsPanel.Controls.Add(this.browseButton);
            this.settingsPanel.Controls.Add(this.locationText);
            this.settingsPanel.Controls.Add(this.locationBox);
            this.settingsPanel.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.settingsPanel.Location = new System.Drawing.Point(0, 30);
            this.settingsPanel.Name = "settingsPanel";
            this.settingsPanel.Size = new System.Drawing.Size(284, 231);
            this.settingsPanel.TabIndex = 6;
            // 
            // deepSearchBox
            // 
            this.deepSearchBox.AutoSize = true;
            this.deepSearchBox.Checked = global::Outlook_MSG_Attachment_Tool.Properties.Settings.Default.deepSearch;
            this.deepSearchBox.DataBindings.Add(new System.Windows.Forms.Binding("Checked", global::Outlook_MSG_Attachment_Tool.Properties.Settings.Default, "deepSearch", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.deepSearchBox.ForeColor = System.Drawing.SystemColors.WindowText;
            this.deepSearchBox.Location = new System.Drawing.Point(12, 55);
            this.deepSearchBox.Name = "deepSearchBox";
            this.deepSearchBox.Size = new System.Drawing.Size(109, 17);
            this.deepSearchBox.TabIndex = 7;
            this.deepSearchBox.Text = "Scan Subfolders";
            this.deepSearchTip.SetToolTip(this.deepSearchBox, "Scan for additional .msg files stored in subfolders. Be default, only .msg files " +
        "found in the original folder are processed.");
            this.deepSearchBox.UseVisualStyleBackColor = true;
            // 
            // browseButton
            // 
            this.browseButton.BackColor = System.Drawing.SystemColors.ControlDark;
            this.browseButton.Enabled = false;
            this.browseButton.ForeColor = System.Drawing.SystemColors.WindowText;
            this.browseButton.ImageAlign = System.Drawing.ContentAlignment.BottomLeft;
            this.browseButton.Location = new System.Drawing.Point(209, 26);
            this.browseButton.Name = "browseButton";
            this.browseButton.Size = new System.Drawing.Size(63, 22);
            this.browseButton.TabIndex = 6;
            this.browseButton.Text = "Browse";
            this.browseButton.UseVisualStyleBackColor = false;
            this.browseButton.Click += new System.EventHandler(this.browseButton_Click);
            // 
            // locationText
            // 
            this.locationText.BackColor = System.Drawing.SystemColors.ControlDark;
            this.locationText.DataBindings.Add(new System.Windows.Forms.Binding("Text", global::Outlook_MSG_Attachment_Tool.Properties.Settings.Default, "savePath", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.locationText.Enabled = false;
            this.locationText.Location = new System.Drawing.Point(23, 26);
            this.locationText.Name = "locationText";
            this.locationText.ReadOnly = true;
            this.locationText.Size = new System.Drawing.Size(180, 22);
            this.locationText.TabIndex = 5;
            this.locationText.Text = global::Outlook_MSG_Attachment_Tool.Properties.Settings.Default.savePath;
            // 
            // locationBox
            // 
            this.locationBox.AutoSize = true;
            this.locationBox.Checked = global::Outlook_MSG_Attachment_Tool.Properties.Settings.Default.singleLocation;
            this.locationBox.DataBindings.Add(new System.Windows.Forms.Binding("Checked", global::Outlook_MSG_Attachment_Tool.Properties.Settings.Default, "singleLocation", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.locationBox.ForeColor = System.Drawing.SystemColors.WindowText;
            this.locationBox.Location = new System.Drawing.Point(12, 3);
            this.locationBox.Name = "locationBox";
            this.locationBox.Size = new System.Drawing.Size(151, 17);
            this.locationBox.TabIndex = 4;
            this.locationBox.Text = "Save to a single location";
            this.locationBoxTip.SetToolTip(this.locationBox, "Output to a specified folder instead of the location of the original file(s).");
            this.locationBox.UseVisualStyleBackColor = true;
            this.locationBox.Click += new System.EventHandler(this.locationBox_Click);
            // 
            // DragForm
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Controls.Add(this.settingsPanel);
            this.Controls.Add(this.settingsButton);
            this.Controls.Add(this.detailLabel);
            this.Controls.Add(this.dragLabel);
            this.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ForeColor = System.Drawing.SystemColors.Control;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "DragForm";
            this.Text = ".MSG Attachment Tool";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.DragForm_FormClosing);
            this.settingsPanel.ResumeLayout(false);
            this.settingsPanel.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private Label detailLabel;
        private Label dragLabel;
        private Button settingsButton;
        private Panel settingsPanel;
        private ToolTip settingsTooltip;
        private ToolTip moveMsgTooltip;
        private Button browseButton;
        private TextBox locationText;
        private CheckBox locationBox;
        private FolderBrowserDialog folderBrowser;
        private CheckBox deepSearchBox;
        private ToolTip deepSearchTip;
        private ToolTip locationBoxTip;
        private ToolTip toolTip1;
    }
}

