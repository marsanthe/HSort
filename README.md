
# HSort

![alt text](https://github.com/marsanthe/HSort/blob/main/Images/HSortLogo.png)

<p>
HSort creates a <strong> Library </strong> (a folder of a certain structure) from your H-Manga, that is compatible with <strong> Kavita </strong>.

[Kavita Reading Server](https://github.com/Kareadita/Kavita)
</p>

## Please Note

1. HSort only works for Manga that are named by the <strong>E-Hentai Naming Convention</strong><br>
2. Kavita usually needs about <strong> three scans </strong>to display the complete content of a Library correctly<br> - for details see [HSort-Wiki/HSort and Kavita]

## Overview

- <strong>This script does not modify your source folder/files in any way</strong>

- HSort creates a <strong> Library </strong> (folder) from your Manga that is compatible with Kavita - [Kavita Naming convention / File Structure](https://wiki2.kavitareader.com/guides/scanner) <br>

- HSort <strong> sorts </strong> all your Manga by Artists (for generic Manga), Conventions (for Doujinshi) and Anthologies<br>

- HSort automatically adds <strong> metadata </strong> (ComicInfo.xml) to all H-Manga of your library<br>
(<strong> No </strong> online scraping!)


## Features

<p>

- Works with any folder.<br>
As long as the folder contains items that match the E-Hentai naming scheme,
HSort will find them.

- Allows you to update your Libraries at any time
- Supports creating multiple libraries from different sources
- Automatically finds Variants and Duplicates of your Manga
- Automatic tagging of your Manga + support for custom tags !
- Creates detailed reports for:
    - Items that were successfully sorted/copied 
    - Items that were skipped (and why they were skipped)
    - An overview over the number of Manga, Doujinshi and Anthologies<br> 
    and number of different Artists and Conventions in your library
- and more -> check the Wiki !

</p>

## Requirements
- Windows 10 or higher 
- Powershell 5.1 or higher
- 7zip installed

## Getting Started

### Please check the Wiki for detailed information

<p>

[HSort Wiki](https://github.com/marsanthe/HSort/wiki)
</p>

### Step 1: Set-Up
<p>

1. Save the HSort folder anywhere
2. Open the HSort folder
3. Inside the folder: [Shift] + [Right Click] -> Select "Open PowerShell window here"
4. Type in .\HSort and hit enter.
5. A Settings.txt file will open. Edit it to your liking.
6. Save and close it.

</p>

### Step 2: Run
<p>

1. Open PowerShell again as above (if you have closed it).
2. Type in .\HSort again and hit enter.
3. Follow the instructions.

</p>

### Step 3: In Kavita

<p> 
Let's say HSort created a library-folder called "MyHLibrary" for you on your Desktop.<br>
This folder then contains three more folders: Logs, ComicInfoFiles and <strong> another </strong> folder also called "MyHLibrary".<br>
This is the folder you want to select in Kavita as source. 
</p>
