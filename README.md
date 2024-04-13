
# HSort

## Note
This script only works for H-Manga that are named by the E-Hentai [Naming Convention](#naming-convention).<br>

## Overview

- Creates a folder from your Manga that is compatible with Kavita - [Kavita Naming convention / File Structure](https://wiki2.kavitareader.com/guides/scanner) <br>

- Sorts all your Manga by Artists (for generic Manga), Conventions (for Doujinshi) and Anthologies.<br>

- Tags your Manga to help you find your stuff quickly in Kavita.

## Features

<p>
Your source files remain unchanged!

- Creates a Library-folder of your Manga that is compatible with Kavita.<br>
(No pre-sorting of files/folders required.) <br>
- Sorts your Manga <br>
- Allows you to update your Library-folder any time [Updating]
- Supports creating multiple libraries from different sources.
- Creates and adds a custom ComicInfo-file to each Manga [ComicInfo].<br>
- Adds Tags to be used in Kavita like<br> Language, Artist, Convention, Content-Type, Date-Created, Censored/Decensored, Digital,...<br>
- Automatically finds Variants of titles and duplicates [Variants and Duplicates](#variants-and-duplicates).
- Creates a report of files/folders
    - That don't match the E-Hentai naming convention
    - That have the wrong file-type (Movies, Exes,...)
    - That are broken Archives

</p>

## Requirements
- Windows 10 or higher 
- Powershell 5.1 or higher
- 7zip installed in .\Program Files

## Manual

### Set-up
<p>

1. Save the HSort folder anywhere
2. Open the HSort folder
3. Inside the folder: [Shift] + [Right Click] -> Select "Open PowerShell window here"
4. Type in .\HSort and hit enter.
5. A text-file will open that you'll customize.
6. Save and close it.

</p>

### Run
<p>

1. Open PowerShell again as above if you have closed it.
2. Type in .\HSort and hit enter.
3. Follow the instructions.

</p>

## Naming Convention

This script only works for files (Zip, Rar, Cbz, Cbr) and folders named as shown below. 

### Generic Manga
    [Artist Names] Title (Magazine or Tankoubon source) [Language] [Translators] [Special Indicators]
### Doujinshi
    (Convention Name) [Circle Names (Artist Names)] Title (Parody Names) [Language] [Translators] [Special Indicators]
### Anthologies
    [Anthology] Title (Parody Names) [Language] [Translators] [Special Indicators]

#### Example
.\MyEcchiManga\\[BUTA] Otonarisan My Neighbor [COMIC HOTMILK 2022-07] [English] [head empty] [Digital].zip


## Adapting H-Manga to Kavita
<p>
Kavita is primarily geard towards readers of western comic-books and "normal" Manga.
To create a comfortable reading experience for H-Manga some changes had to be made to how Kavita interprets certain tags in a ComicInfo-file.
</p>
<p>
Kavita creates <strong>Collections</strong> based on the [Series Group] tag in a ComicInfo file.<br>
Let's assume we have a properly set-up folder containing Batman-Comics and each has a properly configured ComicInfo file. 
In Kavita this would result in a collection in which one element is "Batman", which holds in turn all volumes of Batman-Comics (gross simplification).
</p>
<P>
HSort on the other hand creates a Kavita compatible folder (local library) that is sorted by
</p>

- Artists for Generic Manga
- Conventions for Doujinshi 
- "Anthology" for all Anthologies

So one element in our collection could be "Buta", which contains all H-Manga from Buta.<br>
Another element could be "C90", contining all Doujinshi publsihed by C90.

## Good to know

### Japanese Manga VS American Comics 
<p>
The structure of western Comic-books and "normal" Manga are <em>somewhat</em> similar.<br> There are overarching <strong>Themes</strong> and <strong>Series</strong> that belong to a certain theme.<br>Each series has its own <strong>Title</strong> and usually consist of mulitple <strong>Issues</strong>.
</p>
<p>
As mentioned above, each Comic-Series has a certain <strong>Theme</strong>, like Batman, Spiderman, Superman... which is narrowly defined and revolves around one character or a group of characters. 
</p>
<p>
These themes are reffered to as <strong>Series Group</strong> in ComicInfo-files.
In Kavita they are called <strong>Collections</strong>.
</p>
<p>
In contrast, Manga can only be loosely grouped into <strong>Genres</strong> like Shôjo, Shônen, Gekiga, etc. - which are far less specific and refer to the main readership or the topic the Manga is concerned with (Sports, Adventure,...).  
</p>

### The Role/s of the Creator/s
<p>
One major difference between Comics and Manga is that a Manga-Series is usually the work of a single person - the Manga-ka/Artist/Creator - while a Comic-Series is created by a host of different people, each responsible for a different task (Artists/Writer/Letterer/Penciller/Inker). This is reflected in the structure of ComicInfo-files.
</p>
<p>
When HSort creates a ComicInfo file (for a Manga), tags like Writer, Letterer, Penciller, Inker are omitted, and Artist is defined as Writer. 
</p>

### Manga VS H-Manga

<p>
While Comics and Manga share some similarities, Manga and H-Manga differ in one key aspect. <br>
H-Manga are mostly single-chapter stories, analog to a One-Shot Comic release.<br>
That is why grouping H-Manga by <strong>Series</strong> would do little to structure a set of H-Manga.<br>
</p>

Therefore I decided to do the following conversion [Conversion Table](#conversion-tables-h-manga-on-kavita)

## Conversion Tables (H-Manga on Kavita)

Read [good-to-know](#good-to-know) for the rationale behind this.


### For Generic H-Manga (Grouped by Artist)

<table>
<tbody>
<tr>
<td>ComicInfo</td>
<td>Kavita</td>
<td>HSort (set as)</td>
</tr>
<tr>
<td>Title</td>
<td>Chapter Title</td>
<td>Title</td>
</tr>
<tr>
<td>Series</td>
<td>Name</td>
<td><strong>Title</strong></td>
</tr>
<tr>
<td>Writer, Penciller, Inker</td>
<td>Writer, Penciller, Inker</td>
<td><strong>Artist</strong></td>
</tr>
<tr>
<td>Series Group</td>
<td>Collection</td>
<td><strong>Artist</strong></td>
</tr>
<tr>
<td>Format</td>
<td>Special</td>
<td><strong>One-Shot</strong></td>
</tr>
</tbody>
</table>
<p>
</p>

### For Doujinshi (Grouped by Convention)

<table>
<tbody>
<tr>
<td>ComicInfo</td>
<td>Kavita</td>
<td>HSort (set as)</td>
</tr>
<tr>
<td>Title</td>
<td>Chapter Title</td>
<td>Title</td>
</tr>
<tr>
<td>Series</td>
<td>Name</td>
<td><strong>Title</strong></td>
</tr>
<tr>
<td>Writer, Penciller, Inker</td>
<td>Writer, Penciller, Inker</td>
<td><strong>Artist</strong></td>
</tr>
<tr>
<td>Series Group</td>
<td>Collection</td>
<td><strong>Convention</strong></td>
</tr>
<tr>
<td>Format</td>
<td>Special</td>
<td><strong>One-Shot</strong></td>
</tr>
</tbody>
</table>
<p>
</p>

### For Anthologies (Grouped by "Anthology")

<table>
<tbody>
<tr>
<td>ComicInfo</td>
<td>Kavita</td>
<td>HSort (set as)</td>
</tr>
<tr>
<td>Title</td>
<td>Chapter Title</td>
<td>Title</td>
</tr>
<tr>
<td>Series</td>
<td>Name</td>
<td><strong>Title</strong></td>
</tr>
<tr>
<td>Writer, Penciller, Inker</td>
<td>Writer, Penciller, Inker</td>
<td><strong>"Anthology"</strong></td>
</tr>
<tr>
<td>Series Group</td>
<td>Collection</td>
<td><strong>"Anthology"</strong></td>
</tr>
<tr>
<td>Format</td>
<td>Special</td>
<td><strong>One-Shot</strong></td>
</tr>
</tbody>
</table>
<p>
</p>

## Implementation Details

### Variants And Duplicates
<p>
HSort differentiates between Duplicates and Variants.
</p>

#### Duplicates

<p>
EXAMPLE<br>
Take the following objects located somewhere in a folder called MyManga:


    A) [Taira Mune Suki Iinkai (Okunoha)] Anata no Machi no Shokushuyasan  Your Neighborhood Tentacle Shop.zip

    B) [Taira Mune Suki Iinkai (Okunoha)] Anata no Machi no Shokushuyasan  Your Neighborhood Tentacle Shop

    C) [Taira Mune Suki Iinkai (Okunoha)] Anata no   Machi no Shokushuyasan  Your Neighborhood Tentacle Shop


Let's assume A) was already added to the Library.<br>
When HSort encounters B) it is treated as a Duplicate of A), since both objects share the exact same name.<br>
HSort does not differentiate between folders, files or file-extensions.

When HSort encounters C) it is initially <strong>not</strong> a duplicate - note the extra space between "no" and "Machi"

But HSort <strong>normalizes</strong> object names before comparing them, which removes the additional space and turns C into another duplicate of A.
</p>

#### Variants

EXAMPLE<br>
<p>
Take the following objects located somewhere in a folder called MyManga:
</p>

    A) [BUTA] Otonarisan My Neighbor [COMIC HOTMILK 2022-07] [English] [head empty] [Digital].zip

    B) [BUTA] Otonarisan My Neighbor [English] [head empty] [Decensored].zip


<p>
Object B is <strong>not</strong> a duplicate of object A, since the object names are not identical.<br>
But since both share the exact same <strong>Title</strong>, object B is considered a <strong>Variant</strong> of object A.<br>
</p>
<p>
The rationale behind this is, that objects with the same title but different meta-information can differ in such a way, that it is worth to add both to the library.
</p>