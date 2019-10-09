# Window Opacity Tool
A tool written in QB64 utilizing PowerShell scripts and the Windows API to set opacity levels of Desktop applications (as well as some store apps). A Windows 10 must-have for those who love to customize.

What this program does:
Allows the user to select a window to manipulate and set opacity level.

What this program can't do:
It can't filter out programs that are not compatible and can't force a window to accept opacity changes. The window must already be compatible through the Windows API. It is also NOT a permanent change. It lasts only for as long as the window remains open. You will need to change the opacity again upon reopening the window you have changed.

# NOTICE:
----------------------------------
The program requires PowerShell in order to run correctly. On first launch, the program will check to make sure your computer has been set up to run scripts. From there, you should be able to run the program normally. Sometimes PowerShell will truncate the remaining letters of a window title. Click the refresh button a few times until it is resolved. This is not a fault of my program, but of PowerShell. If you can't find a window in the list, try entering the title of the window EXACTLY as it is in the title bar and it might work. File Explorer is one of the programs that has to be done this way.

This program comes with no warranties of any sort, express or implied.

To see my video explaining the program, click here: https://www.youtube.com/watch?v=1dNukGA9e6k

For support, please post your issues on the board on https://github.com/SpriggsySpriggs/Window-Opacity-Tool or email me at zspriggs14@gmail.com

Created using Inform Beta 8 found here at https://github.com/FellippeHeitor/InForm

Source code will be available in the repository. Requires QB64 to compile.


# Windows 10 Apps that are supported by this program:
This list will expand as more apps are discovered to work. If you find some that are compatible, email me or post them in the board and I will add them here.

Groove Music

Spotify

MyTube Beta

People

Movies & TV

Photos

Video Editor

Windows Terminal

Microsoft To Do

Sticky Notes

Snip & Sketch

Weather

Windows Security (previously Windows Defender)

Microsoft Edge Dev (Chromium, not native Edge)

Translator

Microsoft Remote Desktop

(Most other Windows 10 Apps that have titles shown in either taskbar or titlebar, including File Explorer)
