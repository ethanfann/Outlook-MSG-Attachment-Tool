# Outlook .MSG Attachment Tool

A Windows application which offers a drag-and-drop interface for extracting attachments stored in .msg files

## Getting Started

### Prerequisites

```
Microsoft .Net Framework 4.5 or newer
```

### Installing

1) <a href="https://github.com/ethanfann/Outlook-MSG-Attachment-Tool/releases/download/v1/OutlookMSGAttachmentTool.zip">Download</a>
2) Unzip and run install.exe 

### Usage

Drag and drop a .msg file over the window to extract the attachments store inside and save them to a folder in the same location as the original file. The generated folder will have the same name as the original file.

Or, drag a folder containing multiple .msg files and they will all be processed.

### Options

#### Save to a single location
Instead of creating the folder with the attachments in the same location as the original file, save to a designated area.

#### Scan subfolders
Selecting this option will search any folders contained in the original folder for .msg files.

```
Desktop  
│
└───folder1
│   │   file1.msg
│   │   file2.msg
│   │
│   └───subfolder1
│       │   file3.msg
│       │   file4.msg
│       │   ...

```
For example, folder1 is drag-and-dropped. Without this option selected, only file1 and file2 will be processed while file3 and file4 are ignored.

## Demo

<a href="https://youtu.be/wixZePLY4nc">Demo</a>


## License

This project is licensed under the GPL License - see the [LICENSE.md](LICENSE.md) file for details


