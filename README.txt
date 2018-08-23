# CODE REVIEW TOOLS INTERFACE

VERSION: 1.0 (ACTIVE FOLDER)
DATE: 07/18/2018
Author: Zachary Zhao

Developed to improve efficiency in code review process using specific tools not included in this repo.

======IMPORTANT NOTES========
    In the case of multiple FHX files, the tool selected will act on each file sequentially, with "Operation complete."
	appearing after the last file is processed. 
    When selecting multiple files, only select UP TO 10 files at once. Please do not select any more files, as CRTI 
	may not behave properly. 
    If CRTI becomes unresponsive, please end the process or close out and reopen. If an error occurs, please check if 
        the file(s) you chose can be used with the tool that you selected.
    Try to not have the mouse near the Excel windows that open, as it may cause an error in the file format conversion. 
    Before a tool begins to execute, any CMD or Excel windows that you have open will be closed automatically.
        The Excel files should be saved before closing, but please do this manually to make sure.
    Each button has a tooltip that describes its purpose when hovered over.

=====SETUP=====
Before using the Code Review Tools Interface (CRTI), move the CRTI folder into your C: drive (such that it becomes C:/CRTI).
Any FHX files you want to use CRTI with should be stored inside the FHXFiles folder, 
    located within the CRTI Folder.
To access the interface, run the CRTI executable.


=====FOLDERS=====
The two buttons on the side (FHX Files and Results) each open their respective folders.
  - The FHXFiles folder will hold all FHX files as described in the SETUP section.
  - The Results folder will hold all refined files obtained from processing the FHX files.
      The results will be divided into separate folders based on what tool was used to obtain it, i.e. all results of the 
      FHXSFCCheck tool will be in the FHXSFCCheck folder.
      The Results folder is stored on the OCN-PCSDEVDOC01 server. 
The third button (README) opens this file.

Separate from the FHX and Results folders is the Tools folder, which holds all tools that can be used with CRTI.
    The Tools folder is divided into different sub-sections, which are based on the types of tools they contain.
    The different sub-sections are CLI Tools (command line tools), Drag and Drop Tools, and Excel Macros.


=====INSTRUCTIONS=====
To use CRTI:
    1. Choose either a single or multiple FHX files from the FHXFiles folder where designated.
    2. Click the button of the tool that you want to use.
    3. Click Ok when prompted with "Start operation" to run the tool.
           Note: Clicking cancel will stop the tool from executing. 
    4. Don't click anywhere until a box saying "Operation complete." appears.
    5. Boxes may appear out of nowhere. Do not click in these boxes, as they are part of the interface and are automated. 
 

=====CREDITS=====
Credit for the tools go to Carl Lemp (carl.lemp@snet.net), Marc Colello, and Hector.
Developed for the Automation Team of Genentech, Inc. in Oceanside, California, by Zachary Zhao (zachary.m.zhao@gmail.com).
