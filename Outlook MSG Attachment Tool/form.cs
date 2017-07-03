using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook_MSG_Attachment_Tool.Properties;
using Microsoft.Office.Interop.Outlook;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using PdfSharp.Pdf;
using Color = System.Drawing.Color;


namespace Outlook_MSG_Attachment_Tool
{
    public sealed partial class DragForm : Form
    {
        BackgroundWorker _worker = new BackgroundWorker();

        public DragForm()
        {
            InitializeComponent();

            AllowDrop = true;
            DragEnter += form_DragEnter;
            DragDrop += form_DragDrop;
            DragOver += form_DragOver;
            DragLeave += form_DragLeave;

            _worker.DoWork += worker_DoWork;
            _worker.RunWorkerCompleted += worker_Completed;

            if(String.IsNullOrEmpty(Settings.Default.savePath))
            {
                Settings.Default.savePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\Outlook MSG Tool";
            }

            settingsPanel.Hide();
        }

        void form_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
        }

        void form_DragOver(object sender, DragEventArgs e)
        {
            List<string> msgFilePaths = new List<string>((string[])e.Data.GetData(DataFormats.FileDrop)).FindAll(HasMsgExtension);
            List<string> folderPaths = new List<string>((string[])e.Data.GetData(DataFormats.FileDrop)).FindAll(IsDirectory);

            // Scan each submitted folder (and their subfolders, if enabled in settings) for .msg files.
            foreach (string folder in folderPaths)
            {
                msgFilePaths.AddRange(Settings.Default.deepSearch
                    ? Directory.GetFiles(folder, "*.msg*", SearchOption.AllDirectories)
                    : Directory.GetFiles(folder, "*.msg*"));
            }

            if (msgFilePaths.Count >= 1)
            {
                BackColor = Color.MediumSeaGreen;

            }
            else
            {
                BackColor = Color.Red;
            }
            dragLabel.ForeColor = Color.White;
        }

        void form_DragLeave(object sender, EventArgs e)
        {
            BackColor = DefaultBackColor;
            dragLabel.ForeColor = SystemColors.ButtonShadow;
        }

        void form_DragDrop(object sender, DragEventArgs e)
        {
            BackColor = DefaultBackColor;
            dragLabel.ForeColor = SystemColors.ButtonShadow;

            if (settingsPanel.Visible)
            {
                settingsPanel.Hide();
                Settings.Default.Save();
            }

            Cursor.Current = Cursors.WaitCursor;

            _worker.RunWorkerAsync(new List<string>((string[])e.Data.GetData(DataFormats.FileDrop)));
        }

        async Task delay()
        {
            await Task.Delay(5000);
        }

        async void ShowDetailLabel(string labelText)
        {
            detailLabel.Text = labelText;
            detailLabel.Show();

            await delay();

            detailLabel.Hide();
        }

        private void settingsButton_Click(object sender, EventArgs e)
        {
            if (settingsPanel.Visible)
            {
                settingsPanel.Hide();
            }
            else
            {
                settingsPanel.Show();

                if (Settings.Default.singleLocation)
                {
                    locationText.BackColor = SystemColors.Control;
                    browseButton.BackColor = SystemColors.Control;

                    locationText.Enabled = true;
                    browseButton.Enabled = true;
                }
                else
                {
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
        private void locationBox_Click(object sender, EventArgs e)
        {
            Settings.Default.Save();

            if (Settings.Default.singleLocation)
            {
                locationText.BackColor = SystemColors.Control;
                browseButton.BackColor = SystemColors.Control;

                locationText.Enabled = true;
                browseButton.Enabled = true;
            }
            else
            {
                locationText.BackColor = SystemColors.ControlDark;
                browseButton.BackColor = SystemColors.ControlDark;

                locationText.Enabled = false;
                browseButton.Enabled = false;
            }
        }

        // Sets the value of the single location save path to the result of the folderbrowser dialog
        private void browseButton_Click(object sender, EventArgs e)
        {
            if (folderBrowser.ShowDialog() == DialogResult.OK)
            {
                Settings.Default.savePath = folderBrowser.SelectedPath;

                Settings.Default.Save();
            }
        }

        // Returns a pdf file containing useful information extracted from the msg file
        private Document GeneratePdf(MailItem msg)
        {
            Document pdf = new Document();
            Section section = pdf.AddSection();

            section.AddParagraph("From: " + msg.SenderEmailAddress);
            section.AddParagraph("To: " + msg.To);
            
            if (msg.CC != null)
            {
                section.AddParagraph("CC: " + msg.CC);
            }

            section.AddParagraph("Date Received: " + msg.ReceivedTime.ToString("M/d/yyyy h:mm tt"));

            section.AddParagraph();

            if (msg.Subject != null)
            {
                section.AddParagraph("Subject: " + msg.Subject);
            }

            section.AddParagraph();

            if (msg.Body != null)
            {
                section.AddParagraph(msg.Body);
            }

            return pdf;
        }

        // Predicate for finding all files with the .msg extension in a list of paths
        private static bool HasMsgExtension(string path)
        {
            return (Path.GetExtension(path) == ".msg");
        }

        // Predicate for finding all folders in a list of paths
        private static bool IsDirectory(string path)
        {
            return ((File.GetAttributes(path) & FileAttributes.Directory) == FileAttributes.Directory);
        }

        private void DragForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            Settings.Default.Save();
        }

        private void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            Stopwatch timer = new Stopwatch();
            timer.Start();

            var ol = new Microsoft.Office.Interop.Outlook.Application();
            List<string> msgFilePaths = ((List<string>) e.Argument).FindAll(HasMsgExtension);
            List<string> folderPaths = ((List<string>) e.Argument).FindAll(IsDirectory);

            // Scan each submitted folder (and their subfolders, if enabled in settings) for .msg files.
            foreach (string folder in folderPaths)
            {
                msgFilePaths.AddRange(Settings.Default.deepSearch
                    ? Directory.GetFiles(folder, "*.msg*", SearchOption.AllDirectories)
                    : Directory.GetFiles(folder, "*.msg*"));
            }

            int attachmentCount = 0; // Stores total number of attachments. Displayed at end of processing in status text
            foreach (string filePath in msgFilePaths)
            {
                // Use the outlook interop libary to create a MailItem object from the given .msg file
                MailItem msg = ol.CreateItemFromTemplate(filePath);

                string folderPath = "";
                if(Settings.Default.singleLocation)
                {
                    // Test if the save path has been set if the single location save option is selected
                    if(String.IsNullOrEmpty(Settings.Default.savePath))
                    {
                        e.Result = $"Please choose a save location or disable the single save location option in the settings menu";
                        timer.Stop();
                        return;
                    }
                    else
                    {
                        // Combine the single location save path and the file
                        Console.WriteLine(Path.GetFileNameWithoutExtension(filePath));
                        Console.WriteLine(filePath.Remove(filePath.Length - 4));
                        folderPath = Settings.Default.singleLocation
                            ? Path.Combine(Settings.Default.savePath, Path.GetFileNameWithoutExtension(filePath))
                            : filePath.Remove(filePath.Length - 4);
                    }
   
                }
                else
                {
                    folderPath = filePath.Remove(filePath.Length - 4);
                }

                if (Directory.Exists(folderPath))
                {
                    // Adds the current time to the end of a foldername in the event of duplicate folders being created
                    folderPath = folderPath + " " + DateTime.Now.ToString("MM-dd-yyyy HHmmss");
                    Directory.CreateDirectory(folderPath);
                }
                else
                {
                    Directory.CreateDirectory(folderPath);
                }

                if (Settings.Default.createPDF)
                {
                    Document pdf = GeneratePdf(msg);

                    PdfDocumentRenderer renderer = new PdfDocumentRenderer(false, PdfFontEmbedding.Always)
                    {
                        Document = pdf
                    };

                    renderer.RenderDocument();
                    renderer.PdfDocument.Save(Path.Combine(folderPath, msg.Subject + ".pdf"));
                }

                foreach (Attachment attachment in msg.Attachments)
                {
                    attachment.SaveAsFile(Path.Combine(folderPath, attachment.FileName));
                    attachmentCount++;
                }
            }
            timer.Stop();

            e.Result =
                $"Processed {msgFilePaths.Count} .msg Files(s) and extracted {attachmentCount} attachment(s) in {timer.ElapsedMilliseconds/1000.0} seconds.";

        }
   
        public void worker_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            Cursor.Current = Cursors.Default;
            ShowDetailLabel(e.Result.ToString());
        }

    }
}


