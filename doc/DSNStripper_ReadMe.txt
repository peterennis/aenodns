DSNStripper Utility
Author: Paul Litwin, Fred Hutchinson Cancer Research Center
plitwin@fhcrc.org
6/14/2004

Licensing: Anyone can use this program or modify the source code in anyway.
The author makes no warranties as to its use or misuse.

Compatability: should work with Microsoft Access 2000, 2002, and 2003.

There is a problem with Access links to SQL Server in that when you create links through the UI they are dependent on ODBC DSN's (Data Source Names) being on the client desktop. This is pain in the neck if you will be using the database from multiple desktops.

Thus, I saw a need and I created an Access add-in called DSN Stripper. The idea is this: go ahead and create your SQL Server linked tables using the UI and then when you are done use this add-in to convert the links to DSN-less links. That is, it converts the links to links that do not depend on DSNs being created on the user's desktop. I created it as an add-in which makes it easy to use across Access databases.

To install the utility as an Access add-in:
1. Unzip the zip file and place DSNStripper.mda in any directory.
2. Open any Access database and select Tools|Add-Ins|Add-in Manager
3. Click Add New and navigate to the DSNStripper.mda file.
4. Click close at the Add-in dialog.

To use the add-in:
1. Create some links to SQL Server tables using the UI. 
2. Startup the add-in by selecting  Tools|Add-Ins|DSN Stripper.
3. The add-in needs to know the name of the SQL Server server, which defaults to localhost. (For future use, it remembers the last setting).
4. Optionally, it will also strip off the annoying "dbo_" prefix that Access adds to the name of the link.
5. Click the Convert button and close the form when it is done.

Hope it helps,
Paul