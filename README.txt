Excel Image Assistant - Mohamed v1.7
====================================
Independent clean-room add-in build. No QLM / activation DLLs are used.
Info name: Mohamed Abd Elnasser - IT Manager

IMPORTANT INSTALL ORDER
1) Close ALL Excel windows.
2) Run "Build Add-in.cmd".
3) Confirm the build ends with:
   XLAM package + Ribbon relationship validation: PASS
   DONE - v1.7 validated build
4) Run "Install Add-in.cmd".
5) Close Excel completely if it opened in the background, then reopen Excel.

Visible Ribbon tab: Excel Image Assistant
Internal Ribbon IDs are unique in v1.7 to avoid conflicts/cache collisions with the original add-in or older clone builds.

Image behavior
- Set your image folder in Settings.
- Select the cells/column containing codes such as 2611029.
- The add-in searches the chosen folder for 2611029.jpg / .jpeg / .png / etc.
- Settings > Insert image beside code = Left or Right.
- The picture is fitted inside the adjacent destination cell; it is never placed over the code cell.

If RibbonX is blocked by Office, v1.7 also creates an Excel Image Assistant fallback menu under Excel's Add-ins UI when macros are allowed.
