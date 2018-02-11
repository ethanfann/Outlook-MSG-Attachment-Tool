using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook_MSG_Attachment_Tool.Properties;
using MsgReader.Outlook;
using Color = System.Drawing.Color;


namespace Outlook_MSG_Attachment_Tool {
    public sealed partial class DragForm : Form {
        BackgroundWorker _worker = new BackgroundWorker();

        public DragForm() {
            InitializeComponent();

            AllowDrop = true;
            DragEnter += form_DragEnter;
            DragDrop += form_DragDrop;
            DragOver += form_DragOver;
            DragLeave += form_DragLeave;

            _worker.DoWork += worker_DoWork;
            _worker.RunWorkerCompleted += worker_Completed;

            if (String.IsNullOrEmpty(Settings.Default.savePath)) {
                Settings.Default.savePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\Outlook MSG Tool";
            }

            settingsPanel.Hide();
        }

        // Updates cursor with a "copy" icon when files are dragged over the form
        void form_DragEnter(object sender, DragEventArgs e) {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
        }

        // Updates background of form depending if .msg files have been found in drag-and-dropped files
        void form_DragOver(object sender, DragEventArgs e) {
            List<string> msgFilePaths = new List<string>((string[])e.Data.GetData(DataFormats.FileDrop)).FindAll(HasMsgExtension);
            List<string> folderPaths = new List<string>((string[])e.Data.GetData(DataFormats.FileDrop)).FindAll(IsDirectory);

            // Scan each submitted folder (and their subfolders, if enabled in settings) for .msg files.
            foreach (string folder in folderPaths) {
                msgFilePaths.AddRange(Settings.Default.deepSearch
                    ? Directory.GetFiles(folder, "*.msg*", SearchOption.AllDirectories)
                    : Directory.GetFiles(folder, "*.msg*"));
            }

            if (msgFilePaths.Count >= 1) {
                BackColor = Color.MediumSeaGreen;
            }
            else {
                BackColor = Color.Red;
            }
            dragLabel.ForeColor = Color.White;
        }

        // Reset the backgroun color of the form when cursor leaves form
        void form_DragLeave(object sender, EventArgs e) {
            BackColor = DefaultBackColor;
            dragLabel.ForeColor = SystemColors.ButtonShadow;
        }

        // Runs the attachment extraction in the background when files have been dragged and dropped
        void form_DragDrop(object sender, DragEventArgs e) {
            BackColor = DefaultBackColor;
            dragLabel.ForeColor = SystemColors.ButtonShadow;

            if (settingsPanel.Visible) {
                settingsPanel.Hide();
                Settings.Default.Save();
            }

            Cursor.Current = Cursors.WaitCursor;

            _worker.RunWorkerAsync(new List<string>((string[])e.Data.GetData(DataFormats.FileDrop)));
        }

        async Task delay(int timeToWait) {
            await Task.Delay(timeToWait);
        }

        // Shows the detail label containing information about the last attachment extraction for 5 seconds, then hides.
        async void ShowDetailLabel(string labelText) {
            detailLabel.Text = labelText;
            detailLabel.Show();

            await delay(5000);

            detailLabel.Hide();
        }

        // Toggles the visibility of the settings checkboxes
        private void settingsButton_Click(object sender, EventArgs e) {
            if (settingsPanel.Visible) {
                settingsPanel.Hide();
            }
            else {
                settingsPanel.Show();

                if (Settings.Default.singleLocation) {
                    locationText.BackColor = SystemColors.Control;
                    browseButton.BackColor = SystemColors.Control;

                    locationText.Enabled = true;
                    browseButton.Enabled = true;
                }
                else {
                    locationText.BackColor = SystemColors.ControlDark;
                    browseButton.BackColor = SystemColors.ControlDark;

                    locationText.Enabled = false;
                    browseButton.Enabled = false;
                }
            }
        }

        /* 
            Enables/Disables the single save location box and browse button based on the state of the 
            single location save checkbox
        */
        private void locationBox_Click(object sender, EventArgs e) {
            Settings.Default.Save();

            if (Settings.Default.singleLocation) {
                locationText.BackColor = SystemColors.Control;
                browseButton.BackColor = SystemColors.Control;

                locationText.Enabled = true;
                browseButton.Enabled = true;
            }
            else {
                locationText.BackColor = SystemColors.ControlDark;
                browseButton.BackColor = SystemColors.ControlDark;

                locationText.Enabled = false;
                browseButton.Enabled = false;
            }
        }

        // Sets the value of the single location save path to the result of the folderbrowser dialog
        private void browseButton_Click(object sender, EventArgs e) {
            if (folderBrowser.ShowDialog() == DialogResult.OK) {
                Settings.Default.savePath = folderBrowser.SelectedPath;

                Settings.Default.Save();
            }
        }

        // Predicate for finding all files with the .msg extension in a list of paths
        private static bool HasMsgExtension(string path) {
            return (Path.GetExtension(path) == ".msg");
        }

        // Predicate for finding all folders in a list of paths
        private static bool IsDirectory(string path) {
            return ((File.GetAttributes(path) & FileAttributes.Directory) == FileAttributes.Directory);
        }

        private void DragForm_FormClosing(object sender, FormClosingEventArgs e) {
            Settings.Default.Save();
        }

        private void worker_DoWork(object sender, DoWorkEventArgs e) {
            Stopwatch timer = new Stopwatch();
            timer.Start();

            List<string> msgFilePaths = ((List<string>)e.Argument).FindAll(HasMsgExtension);
            List<string> folderPaths = ((List<string>)e.Argument).FindAll(IsDirectory);

            // Scan each submitted folder (and their subfolders, if enabled in settings) for .msg files.
            foreach (string folder in folderPaths) {
                msgFilePaths.AddRange(Settings.Default.deepSearch
                    ? Directory.GetFiles(folder, "*.msg*", SearchOption.AllDirectories)
                    : Directory.GetFiles(folder, "*.msg*"));
            }

            int attachmentCount = 0; // Stores total number of attachments. Displayed at end of processing in status text
            foreach (string filePath in msgFilePaths) {

                using (var msg = new Storage.Message(filePath)) {
                    string folderPath = "";
                    if (Settings.Default.singleLocation) {
                        // Test if the single location save option is selected
                        if (String.IsNullOrEmpty(Settings.Default.savePath)) {
                            e.Result = $"Please choose a save location or disable the single save location option in the settings menu";
                            timer.Stop();
                            return;
                        }
                        else {
                            // Combine the single location save path and the file
                            Console.WriteLine(Path.GetFileNameWithoutExtension(filePath));
                            Console.WriteLine(filePath.Remove(filePath.Length - 4));
                            folderPath = Settings.Default.singleLocation
                                ? Path.Combine(Settings.Default.savePath, Path.GetFileNameWithoutExtension(filePath))
                                : filePath.Remove(filePath.Length - 4);
                        }

                    }
                    else {
                        folderPath = filePath.Remove(filePath.Length - 4);
                    }

                    if (Directory.Exists(folderPath)) {
                        // Adds the current time to the end of a foldername in the event of duplicate folders being created
                        folderPath = folderPath + " " + DateTime.Now.ToString("MM-dd-yyyy HHmmss");
                        Directory.CreateDirectory(folderPath);
                    }
                    else {
                        Directory.CreateDirectory(folderPath);
                    }


                    foreach (Storage.Attachment attachment in msg.Attachments) {

                        File.WriteAllBytes(Path.Combine(folderPath, attachment.FileName), attachment.Data);
                        // attachment.SaveAsFile(Path.Combine(folderPath, attachment.FileName));
                        attachmentCount++;
                    }
                }
            }
            timer.Stop();

            e.Result =
                $"Processed {msgFilePaths.Count} .msg Files(s) and extracted {attachmentCount} attachment(s) in {timer.ElapsedMilliseconds / 1000.0} seconds.";

        }

        // Runs when the background worker has finished extracting attachments. Show detail label with time taken and number of attachments extracted
        public void worker_Completed(object sender, RunWorkerCompletedEventArgs e) {
            Cursor.Current = Cursors.Default;
            ShowDetailLabel(e.Result.ToString());
        }

    }
}


