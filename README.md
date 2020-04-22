# viewIT
Allows you to display multiple web pages in timed succession such as network monitors, performance indicators, status pages, security cameras, photo-albums, etc. Define how many seconds each page is displayed and the location of pages on the Internet, a network path, or local disk. Great for use in datacenters, NOCs, meeting rooms, reception, or kiosks.
 
Runs as a stand-alone program or Screen Saver and optionally prompts for a password in order to close. Capable of displaying anything Internet Explorer can such as HTML, ASP, XML, PHP, Java, JPG, GIF, TXT, SWF/Flash, Shockwave, and much more.

Alternate versions:

 - viewIT-rev1 v2.5.0.1 ...INI format is ~ as separator instead of comma
 - viewIT-rev2 v2.5.0.2 ...INI format is | as separator instead of comma
 - viewIT-rev3 v2.5.0.3 ...larger window scale to accommodate modern monitor/tablet resolutions
 - viewIT-rev4 v2.5.0.4 ...larger window scale to accommodate modern monitor/tablet resolutions, INI format is | as separator instead of comma

Tested on Windows 2000, XP, Vista, 7, 8, 8.1 and Server 2003, 2008, 2012

<b>Installation:</b>

1) Run viewit-setup.exe and follow the wizard
2) Modify viewit.ini as indicated below
3) Run viewit.exe or click on the viewIT shortcut in the Start menu to launch

viewIT can also be run as a screen saver by coping viewit.exe and viewit.ini to the Windows directory and then renaming viewit.exe to viewit.scr

<b>.ini Settings:</b>

To display your content you need to edit the viewit.ini file in the directory with the viewIT executable, this is usually located in C:\Program Files\viewIT unless you put it somewhere else.

Add on each line: the content to display, followed by how many seconds you want that displayed
Pages can be referenced from a local or network drive such as...

    c:\directory\page1.htm, 12
    w:\directory\page2.htm, 20
      or pages can be referenced by URL...
    http://example.com/page1.htm, 5
    http://example.com/page2.htm, 10

If you have only one page, or don't want a page to advance (unless right arrow is pressed) include a zero...

    http://example.com, 0

In some cases you may be able to include login information with the URL... (see KB834489)

    http://username:password@www.example.com, 20

Un-comment to prevent the program from being closed without the specified password (default is commented as shown, no password)

    ;Password=changeme

Comment to hide the mouse cursor (when run as a screen saver the cursor is always hidden regardless of this setting)

    Cursor=Yes

Comment so the first mouse click will not automatically pause the timing (pressing P will turn timing back on again)

    ClickPause=Yes

<b>Usage:</b>

- Press right or left arrow keys to jump forward or back (or right/left on the MCE Remote)
- Press P to pause timing and stay at the current page, press P again to continue
- Press Q or ESC to exit (or Clear on the MCE Remote)

To suppress popup messages about script errors from within the viewIT window disable script debugging, see: http://jpelectron.com/sample/WWW%20and%20HTML/disable%20script%20debugging.txt
