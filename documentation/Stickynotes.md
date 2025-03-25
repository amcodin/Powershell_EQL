## Location of Sticky Notes Data File:

1. Open **File Explorer** (Windows + E).
2. In the address bar, type or paste the following path and press Enter:
%AppData%\Microsoft\Sticky Notes

This will take you to the folder where the Sticky Notes data is stored.

## The Data Files:

Inside this folder, you will find a file named **StickyNotes.snt** (or a newer version might be named **plum.sqlite** for the UWP version of Sticky Notes).

- **StickyNotes.snt** (if you're using the older version of Sticky Notes) contains your notes.
- If you have the newer **UWP version** of Sticky Notes, the data might be stored in **plum.sqlite**.


## Locating Sticky Notes Data in Windows 10

It's important to know that the location of Sticky Notes data has changed across different Windows 10 versions. Here's a breakdown:

* **For Windows 10 version 1607 and later:**
    * The Sticky Notes data is stored in a `.sqlite` database file.
    * You can find it here: `C:\Users\[YourUserName]\AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\`
    * The key file you're looking for is typically named `plum.sqlite`.

* **For older Windows 10 versions:**
    * The data was stored in a `.snt` file.
    * This file was located here: `C:\Users\[YourUserName]\AppData\Roaming\Microsoft\Sticky Notes\`

**Important Considerations:**

* **Hidden Files:**
    * The `AppData` folder is hidden by default. You'll need to enable "Show hidden items" in File Explorer to see it.
* **Transferring the Data:**
    * When transferring the data to a new computer, ensure you copy all relevant files from the `LocalState` folder.
    * It is important to understand that if the version of sticky notes on the recieving computer is vastly different than the sending computer, that there could be issues with the data transfer.
* **Microsoft Account Sync:**
    * If you're signed in to Sticky Notes with your Microsoft account, your notes should automatically sync. This is often the easiest way to transfer them.
* It is always a good idea to back up the files before moving them, in case of data corruption.

## Overview

On Windows 10 (version 1607 and later) and Windows 11, Sticky Notes data is stored in a SQLite database file named `plum.sqlite`. Here's how to transfer it to a new computer:

## Steps to Transfer Sticky Notes

### 1. Enable Hidden Files:
- Open **File Explorer**, click on the **"View"** tab, and check **"Hidden items"** to make hidden folders visible.

### 2. Locate the File:
- Navigate to the following path and copy the `plum.sqlite` file:

C:\Users<YourUsername>\AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState


### 3. Transfer the File:
- Paste the `plum.sqlite` file into the same location on the new computer.

### 4. Restart Sticky Notes:
- Open the Sticky Notes app on the new computer, and your notes should appear.

## Note for Older Windows Versions

For older versions of Windows 10 (prior to version 1607), Sticky Notes were stored in a file named `StickyNotes.snt` at this location:

C:\Users<YourUsername>\AppData\Roaming\Microsoft\Sticky Notes