#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         Zachary Zhao

 Script Function:
   User-friendly interface for code review tools
   Makes tool use more efficient, encouraging users to utilize these useful tools
   Tools in question developed by Carl Lemp, Marc Colello, and Hector

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
; includes
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <WinAPIProc.au3>
#include <FontConstants.au3>
#include <ColorConstants.au3>
#include <EditConstants.au3>
#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <StaticConstants.au3>
#include <AutoItConstants.au3>
#include <Constants.au3>
#include <Array.au3>
#include <Excel.au3>
#include <File.au3>
#include <GuiListView.au3>
#include <ListViewConstants.au3>

Opt('MustDeclareVars', 1)

; Execution
MainApp()

Func MainApp()
   ; Variable declarations
   Local $GUI, $msg
   Global $fileChosen = ""
   Global Const $ARR_SIZE = 10
   Global Const $COLOR_SIDEBAR = 0x254566
   Global Const $TOOL_COLOR = 0x19D1AC
   Global Const $INSTR_COLOR = 0x3899DB
   Global Const $FILE_COLOR = 0x29B5C4
   Global $multiFileArray[$ARR_SIZE]
   Global $size

   Global $path = "C:"   ; Path for running on user machine
   Global $server = "\\OCN-PCSDEVDOC01\Reports"  ; Server path where results are stored


   ; Create child window for use cases/tool descriptions ======================================================================
   Local $useCaseGUI
   ;$useCaseGUI = GUICreate("Tool Use Cases", 800, 600)
   $useCaseGUI = GUICreate("Tool Descriptions", 1400, 600)
   GUISetState(@SW_HIDE)

   ; Use Case UI elements
   ;Local $useCaseLabel = GUICtrlCreateLabel("Click a use case to see which tools are applicable for it.", 10, 10, 500, 20)
   ;GUICtrlSetFont($useCaseLabel, 14, $FW_BOLD)
   ;Local $testButton = GUICtrlCreateButton("Recipe Parameter Deferral", 10, 50)
   ;GUICtrlSetCursor($testButton, 0)

   ;Local $applicableToolsLabel = GUICtrlCreateLabel("Applicable Tools: None", 10, 450, 790, 150)
   ;GUICtrlSetFont($applicableToolsLabel, 14, $FW_BOLD)

   ; Tool Description UI Elements
   Local $tableListview = GUICtrlCreateListView("Tool Name|Tool Description|Reason", 0, 0, 1390, 600, Default, BitOR($LVS_EX_GRIDLINES, $LVS_EX_INFOTIP))
   _GUICtrllistView_SetColumnWidth($tableListview, 0, 140)
   _GUICtrlListView_SetColumnWidth($tableListview, 1, 900)
   _GUICtrlListView_SetColumnWidth($tableListview, 2, 321)

   ; Tool Table Items ============================================================
   GUICtrlCreateListViewItem("AlarmConfig|Lists all configurable information for the alarms in the FHX: Module, Alarm Name, Alarm Type, Priority, Enable, Inverted, Mon_Attr, Alarm_Attr, Limit_Attr, Param 1,|Helps check the alarm configuration for a set of CMI's.", $tableListview)
   GUICtrlCreateListViewItem("|Param 2.|", $tableListview)
   GUICtrlCreateListViewItem("", $tableListview)

   GUICtrlCreateListViewItem("AliasByUnitClass|Verify alias assignment across units. Check alias ignores. Check for wrong class assigned to an alias. Check for unused aliases. Check for redundant aliases in the class.|Helps with bulk unit verification.", $tableListview)
   GUICtrlCreateListViewItem("", $tableListview)

   GUICtrlCreateListViewItem("AnalogInfo|List configuration information for analog modules in FHX: Object, Type, Description, Controller, Primary Display, DV Type, DV Sub Type, OUT_EU0, OUT_EU100,|Helps verify configuration prior to loop checks.", $tableListview)
   GUICtrlCreateListViewItem("|OUT_UNITS, PV_EU0, PV_EU100, PV_UNITS, IO_IN, IO_OUT, L_TYPE.|", $tableListview)
   GUICtrlCreateListViewItem("", $tableListview)

   GUICtrlCreateListViewItem("AreaLibObjList|Creates a list of all objects in the FHX where each object is tagged with the Area, Process Cell, and Unit. The list includes the library objects (classes, composites) that|By creating Excel Pivot Tables from the data in the", $tableListview)
   GUICtrlCreateListViewItem("|are used by the object.|resulting file, you can get a list of the shared objects that", $tableListview)
   GUICtrlCreateListViewItem("||are used in each area or process cell.", $tableListview)
   GUICtrlCreateListViewItem("", $tableListview)

   GUICtrlCreateListViewItem("BulkSearchReplace|Performs search/replace within a file using a set of search/replace defined in a separate file. It can replace any number of text patterns with any number of replacement|Can be used as an engineering tool to make updates in", $tableListview)
   GUICtrlCreateListViewItem("|OUT_UNITS, PV_EU0, PV_EU100, PV_UNITS, IO_IN, IO_OUT, L_TYPE.|an export file.", $tableListview)
   GUICtrlCreateListViewItem("", $tableListview)

   GUICtrlCreateListViewItem("DiffTool|Creates a version of the FHX files that are much easier to compare by stripping out lines that are always different but not significant and adding context to every line in|Saves time when comparing FHX files by allowing the ", $tableListview)
   GUICtrlCreateListViewItem("|the file so, when differences are marked, it is easy to see exactly which DeltaV object they go with.|files to be selected from Windows Explorer and making it", $tableListview)
   GUICtrlCreateListViewItem("||easy to find exactly where the changes are in the code.", $tableListview)
   GUICtrlCreateListViewItem("", $tableListview)

   GUICtrlCreateListViewItem("HistoryParams|List all parameters set up for History collection. All the attributes of the history record are reported.|Helps check history collection configuration.", $tableListview)
   GUICtrlCreateListViewItem("", $tableListview)

   GUICtrlCreateListViewItem("ModInstanceParams|Listing all module instance parameters and the values set in the instances. Outputs: Module, Module Class, Parameter Name/Path, Parameter Value, Value Set By.|Helps with bulk CMI verification and testing.", $tableListview)
   GUICtrlCreateListViewItem("", $tableListview)

   GUICtrlCreateListViewItem("ModParamsBuildOPC|List all isntance configurable parameters and the values set in the instances. It adds a formula for getting the realtime value using OPC. Outputs: Module,|Use it to compare instance configured value in the", $tableListview)
   GUICtrlCreateListViewItem("|Module Class, Parameter Name/Path, Parameter Type, Parameter Value, OPC Formula for DeltaV Addin.|database to the realtime value in the controller. Realtime", $tableListview)
   GUICtrlCreateListViewItem("||value is read through OPC using the DeltaV Excel Addin.", $tableListview)
   GUICtrlCreateListViewItem("", $tableListview)

   GUICtrlCreateListViewItem("ModTagList|List component module details: Object, Type, Description, Location/Class, Controller, Primary Display, Instrument Display, Detail Display, DV Type, DV Sub Type,|Offer quick way to check multiple modules on system", $tableListview)
   GUICtrlCreateListViewItem("|Unit Display, Unit Area, Simulate, Bypass, Alarm Delay On, and Alarm Enable Delay parameters.|prior to testing to verify the properties and key", $tableListview)
   GUICtrlCreateListViewItem("||parameters are set correctly", $tableListview)
   GUICtrlCreateListViewItem("", $tableListview)

   GUICtrlCreateListViewItem("ObjDateAndCheckout|Searches an FHX, for every component that is found in the file, with the last modification date, user who modified it, and the name of the person who has it checked out.|Helps to build DQ record.", $tableListview)
   GUICtrlCreateListViewItem("", $tableListview)

   GUICtrlCreateListViewItem("ObjDateAndVer|Searches FHX with the last modification date, and the user who modified it for every component in the file.|Helps to see what is checked out prior to trying an import.", $tableListview)
   GUICtrlCreateListViewItem("", $tableListview)

   GUICtrlCreateListViewItem("ParamList|List all controller parameters(unit, USM) in FHX.|Helps with unit verification. Verify unit and USM", $tableListview)
   GUICtrlCreateListViewItem("||parameter values.", $tableListview)
   GUICtrlCreateListViewItem("", $tableListview)

   GUICtrlCreateListViewItem("ParamMap|Creates a map of the chain of recipe parameters showing how each top level parameter is passed to the levels below.|Helps with reviewing parameter promotion through the", $tableListview)
   GUICtrlCreateListViewItem("||recipes.", $tableListview)
   GUICtrlCreateListViewItem("", $tableListview)

   GUICtrlCreateListViewItem("PhaseAliasUsageDetail|List of all expressions within a phase where aliases are used. It reports the location where the alias is used and the type of use (read/write).|Identifies exactly where in the phase the alias is referenced.", $tableListview)
   GUICtrlCreateListViewItem("", $tableListview)

   GUICtrlCreateListViewItem("PhaseParamList|List of recipe and report parameters within a phase. Along with the name, it reports the ID, range, and Eng Units.|Helps with reviewing phases.", $tableListview)
   GUICtrlCreateListViewItem("", $tableListview)

   GUICtrlCreateListViewItem("PhaseRecParamUsage|List of all expressions within a phase where the batch parameters are used. It reports the location where the parameter is used and the type of use (read/write).|Allows one to see where a recipe parameter is used to", $tableListview)
   GUICtrlCreateListViewItem("||check if it is being set to the right value.", $tableListview)
   GUICtrlCreateListViewItem("", $tableListview)

   GUICtrlCreateListViewItem("PrepDiff1Line|Look for differences between two FHX files.|Easier to find differences between two versions of an", $tableListview)
   GUICtrlCreateListViewItem("||FHX file.", $tableListview)
   GUICtrlCreateListViewItem("", $tableListview)

   GUICtrlCreateListViewItem("RecParam|Searches an FHX file for recipe parameters and reports them in a comma delimited list that can be viewed or imported into Excel.|Can be used to see which phase parameters are not", $tableListview)
   GUICtrlCreateListViewItem("||deferred to any recipe parameter.", $tableListview)
   GUICtrlCreateListViewItem("", $tableListview)

   GUICtrlCreateListViewItem("RecTree|The recipe tree is output in a text file with this structure: Procedure Unit_Proc_1|Identify all the recipe components of the recipe for DQ,", $tableListview)
   GUICtrlCreateListViewItem("|Operation_1 Phase_1|import, etc.", $tableListview)
   GUICtrlCreateListViewItem("|Phase_2 Operation_2|", $tableListview)
   GUICtrlCreateListViewItem("|Phase_1 Operation_3|", $tableListview)
   GUICtrlCreateListViewItem("|Phase_1 Unit_Proc_2. It outputs the trees for all the highest level recipe components found.|", $tableListview)
   GUICtrlCreateListViewItem("", $tableListview)

   GUICtrlCreateListViewItem("Search|Searches an FHX file for text or a text pattern and prints the full context of every place the text is found.|Find code that is writing to a module or parameter.", $tableListview)
   GUICtrlCreateListViewItem("", $tableListview)

   GUICtrlCreateListViewItem("SFCCheck|Check SFC logic in phases and/or Equipment Modules for two dozen different routine errors and guideline conformance.|Should be routinely used after editing a phase to check", $tableListview)
   GUICtrlCreateListViewItem("||for typical errors in DeltaV coding.", $tableListview)
   GUICtrlCreateListViewItem("", $tableListview)

   GUICtrlCreateListViewItem("UnlinkInstConfig|Produces a new FHX file with all instance configurable parameters unlinked from the class. NOTE: requires importing the resulting file to the deltaV database so|Ensures instances are not getting values from the class", $tableListview)
   GUICtrlCreateListViewItem("|MAKE A BACKUP FIRST.|so the default value in the class can be changed without", $tableListview)
   GUICtrlCreateListViewItem("||affecting the instances.", $tableListview)
   GUICtrlCreateListViewItem("", $tableListview)

   GUICtrlCreateListViewItem("CMSummary|A list of CM Classes and their properties. A list of module instances with all the properties that are common to all classes. A list of CM instances for each class in the |Helps with review of CM Instance Configuration.", $tableListview)
   GUICtrlCreateListViewItem("|FHX file. Includes columns for instance configurable parameters of the class. Cells contain the value of the parameter for the instance. Shaded cells are linked to the|", $tableListview)
   GUICtrlCreateListViewItem("|default value in the class.|", $tableListview)
   GUICtrlCreateListViewItem("", $tableListview)

   GUICtrlCreateListViewItem("RecipeSummary|Phase - A list of phases referenced by recipe objects. A set of columns to uniquely identify Phase/Op/UP/Proc/Formula. Columns are added for each Recipe/Report |Helps with review of recipes and where phase input", $tableListview)
   GUICtrlCreateListViewItem("|parameter/value in the phase/OP/UP/Proc/Formula. Color of each cell indicates which level the value is from. |parameter is used. Indicates which level the parameters", $tableListview)
   GUICtrlCreateListViewItem("|Recipe - Includes a column for recipe objects. Count the use of each phase in each object. A set of columns to uniquely identify Op/UP/Proc/Formula. Columns are| are being set from.", $tableListview)
   GUICtrlCreateListViewItem("|added for each Recipe parameter/value in the Op/UP/Proc/Formula. Color of each cell indicates which level the value is set from.|", $tableListview)
   GUICtrlCreateListViewItem("", $tableListview)

   GUICtrlCreateListviewItem("UnitSummary|A list of unit instances in the FHX file. List of phases assigned to these units along with their descriptions. List of aliases configured for the unit classes. A column for|Helps with review of Unit Instance Configuration.", $tableListview)
   GUICtrlCreateListviewItem("|each unit instance. Cells contain the name of the module assigned to the alias/or ignored on that unit. List of unit parameters and values configured for the unit class.|", $tableListview)
   GUICtrlCreateListviewItem("|A list of phases and the properties of the phase and their parameters on each unit class.|", $tableListview)
   GUICtrlCreateListViewItem("", $tableListview)

   GUICtrlCreateListViewItem("RecipeParamExtract|Extracts recipe parameters and puts them in a table format with values for each formula shown side by side.|Makes review of recipe parameters easier.", $tableListview)
   ; ==========================================================================================================================

   ; Create normal window 1366 x 768 resolution
   $GUI = GUICreate("Code Review Tools Interface", 1366, 768)
   GUISetState(@SW_SHOW)

   ; Topbar ==================================================================================================
   Local $topLabelPaddingSide = GUICtrlCreateLabel("", 0, 0, 10, 43)
   GUICtrlSetBkColor($topLabelPaddingSide, $COLOR_BLACK)
   GUISetFont(20, $FW_BOLD)

   Local $topLabelPaddingTop = GUICtrlCreateLabel("", 10, 0, 1356, 5)
   GUICtrlSetBkColor($topLabelpaddingTop, $COLOR_BLACK)

   Local $topBar = GUICtrlCreateLabel("Code Review Tools Interface", 10, 5, 1356, 38)
   GUICtrlSetColor($topBar, $COLOR_WHITE)
   GUICtrlSetBkColor($topBar, $COLOR_BLACK)

   ; Sidebar =================================================================================================
   Local $sideBar = GUICtrlCreateLabel("", 0, 43, 150, 730)
   GUICtrlSetBkColor($sideBar, $COLOR_SIDEBAR)
   GUICtrlSetStyle(-1, $SS_LEFT)

   GUISetFont(16, $FW_BOLD)
   Local $sideBarText = GUICtrlCreateLabel("Folders", 38, 64)
   GUICtrlSetColor($sideBarText, $COLOR_WHITE)
   GUICtrlSetBkColor($sideBarText, $COLOR_SIDEBAR)

   GUISetFont(12, $FW_NORMAL)
   Local $fhxButton = GUICtrlCreateButton("FHX Files", 16, 106, 118, 30)
   GUICtrlSetState($fhxButton, $GUI_SHOW)
   GUICtrlSetColor($fhxButton, $COLOR_WHITE)
   GUICtrlSetBkColor($fhxButton, $COLOR_SIDEBAR)
   GUICtrlSetCursor($fhxButton, 0)

   Local $resultsButton = GUICtrlCreateButton("Results", 31, 156, 94, 30)
   GUICtrlSetState($resultsButton, $GUI_SHOW)
   GUICtrlSetColor($resultsButton, $COLOR_WHITE)
   GUICtrlSetBkColor($resultsButton, $COLOR_SIDEBAR)
   GUICtrlSetCursor($resultsButton, 0)

   Local $readmeButton = GUICtrlCreateButton("README", 31, 206, 94, 30)
   GUICtrlSetState($readmeButton, $GUI_SHOW)
   GUICtrlSetColor($readmeButton, $COLOR_WHITE)
   GUICtrlSetBkColor($readmeButton, $COLOR_SIDEBAR)
   GUICtrlSetCursor($readmeButton, 0)

   Local $useCasesButton = GUICtrlCreateButton("Use Cases", 35, 256, 86, 30)
   GUICtrlSetState($useCasesButton, $GUI_SHOW)
   GUICtrlSetColor($useCasesButton, $COLOR_WHITE)
   GUICtrlSetBkColor($useCasesButton, $COLOR_SIDEBAR)
   GUICtrlSetCursor($useCasesButton, 0)

   ; Main window content =======================================================================================
   ; Instruction section
   GUISetFont(16, $FW_BOLD)
   Local $instrAreaLabel = GUICtrlCreateLabel("", 160, 63, 1196, 80, $SS_SUNKEN) ; Note: might change style from $SS_SUNKEN

   Local $instrLabelPadding = GUICtrlCreateLabel("", 160, 63, 15, 32)
   GUICtrlSetBkColor($instrLabelPadding, $INSTR_COLOR)

   Local $instrLabelPadTop = GUICtrlCreateLabel("",160, 53, 1196, 10)
   GUICtrlSetBkColor($instrLabelPadTop, $INSTR_COLOR)

   Local $instrLabel = GUICtrlCreateLabel("Instructions", 175, 63, 1181)
   GUICtrlSetBkColor($instrLabel, $INSTR_COLOR)
   GUICtrlSetColor($instrLabel, $COLOR_WHITE)
   GUISetFont(14, $FW_NORMAL)
   Local $instructions = GUICtrlCreateLabel("Before using this tool, read the accompanying README to ensure proper use. (NOTE: Not all tools work on all files.)", 175, 105)

   ; File selection section
   GUISetFont(16, $FW_BOLD)
   Local $fileAreaLabel = GUICtrlCreateLabel("", 160, 170, 1196, 200, $SS_SUNKEN) ; Note: might change style from $SS_SUNKEN ; Originally 160, 243, 1195, 500
   GUICtrlSetState($fileAreaLabel, $GUI_DISABLE)

   Local $fileTitleLabelPadTop = GUICtrlCreateLabel("", 160, 160, 1196, 10)
   GUICtrlSetBkColor($fileTitleLabelPadTop, $FILE_COLOR)

   Local $fileTitleLabelPadding = GUICtrlCreateLabel("", 160, 170, 15, 32)
   GUICtrlSetBkColor($fileTitleLabelPadding, $FILE_COLOR)

   Local $fileTitleLabel = GUICtrlCreateLabel("1. File Select", 175, 170, 1181, 32)
   GUICtrlSetBkColor($fileTitleLabel, $FILE_COLOR)
   GUICtrlSetColor($fileTitleLabel, $COLOR_WHITE)

   Local $fileButton = GUICtrlCreateButton("Choose File(s)", 170, 212, 157, 36) ; Originally 155, 320, 127, 36
   GUICtrlSetCursor($fileButton, 0)

   Local $fileNameDisplay
   Global $fileDesc = ""
   GUISetFont(14)
   $fileNameDisplay = GUICtrlCreateLabel("File(s) Chosen: None", 340, 218, 1000, 140)

   ; Tools section
   GUISetFont(16, $FW_BOLD)
   Local $toolAreaLabel = GUICtrlCreateLabel("", 160, 400, 1196, 353, $SS_SUNKEN) ; Note: might change style from $SS_SUNKEN ; Originally 160, 243, 1195, 500
   GUICtrlSetState($toolAreaLabel, $GUI_DISABLE)

   Local $titleLabelPadTop = GUICtrlCreateLabel("", 160, 390, 1196, 10)
   GUICtrlSetBkColor($titleLabelPadTop, $TOOL_COLOR)

   Local $titleLabelPadding = GUICtrlCreateLabel("", 160, 400, 15, 32)
   GUICtrlSetBkColor($titleLabelPadding, $TOOL_COLOR)

   Local $titleLabel = GUICtrlCreateLabel("2. Tools", 175, 400, 1181, 32)
   GUICtrlSetBkColor($titleLabel, $TOOL_COLOR)
   GUICtrlSetColor($titleLabel, $COLOR_WHITE)

   ; Tool buttons=========================================
   ; CLI Section ==============================
   GUISetFont(12, $FW_NORMAL)
   Local $FHXAlarmConfigButton = GUICtrlCreateButton("AlarmConfig", 220, 442, 171, 30)
   GUICtrlSetCursor($FHXAlarmConfigButton, 0)

   Local $FHXAliasByUnitClassButton = GUICtrlCreateButton("AliasByUnitClass", 401, 442, 171, 30)
   GUICtrlSetCursor($FHXAliasByUnitClassButton, 0)

   Local $FHXAnalogInfoButton = GUICtrlCreateButton("AnalogInfo", 582, 442, 171, 30)
   GUICtrlSetCursor($FHXAnalogInfoButton, 0)

   Local $FHXAreaLibObjListButton = GUICtrlCreateButton("AreaLibObjList", 763, 442, 171, 30)
   GUICtrlSetCursor($FHXAreaLibObjListButton, 0)

   Local $FHXBulkTextProcButton = GUICtrlCreateButton("BulkSearchReplace", 944, 442, 171, 30)
   GUICtrlSetCursor($FHXBulkTextProcButton, 0)

   Local $FHXDiffToolButton = GUICtrlCreateButton("DiffTool", 1125, 442, 171, 30)
   GUICtrlSetCursor($FHXDiffToolButton, 0)

   Local $FHXHistoryParamsButton = GUICtrlCreateButton("HistoryParams", 220, 482, 171, 30)
   GUICtrlSetCursor($FHXHistoryParamsButton, 0)

   Local $FHXModuleInstanceParamsButton = GUICtrlCreateButton("ModInstanceParams", 401, 482, 171, 30)
   GUICtrlSetCursor($FHXModuleInstanceParamsButton, 0)

   Local $FHXModParamsBuildOPCButton = GUICtrlCreateButton("ModParamsBuildOPC", 582, 482, 171, 30)
   GUICtrlSetCursor($FHXModParamsBuildOPCButton, 0)

   Local $FHXModTagListButton = GUICtrlCreateButton("ModTagList", 763, 482, 171, 30)
   GUICtrlSetCursor($FHXModTagListButton, 0)

   Local $FHXObjDateAndCheckOutButton = GUICtrlCreateButton("ObjDateAndCheckOut", 944, 482, 171, 30)
   GUICtrlSetCursor($FHXObjDateAndCheckOutButton, 0)

   Local $FHXObjDateAndVerButton = GUICtrlCreateButton("ObjDateAndVer", 1125, 482, 171, 30)
   GUICtrlSetCursor($FHXObjDateAndVerButton, 0)

   Local $FHXParamListButton = GUICtrlCreateButton("ParamList", 220, 522, 171, 30)
   GUICtrlSetCursor($FHXParamListButton, 0)

   Local $FHXParamMapButton = GUICtrlCreateButton("ParamMap", 401, 522, 171, 30)
   GUICtrlSetCursor($FHXParamMapButton, 0)

   Local $FHXPhaseAliasUsageDetailButton = GUICtrlCreateButton("PhaseAlsUsageDetail", 582, 522, 171, 30)
   GUICtrlSetCursor($FHXPhaseAliasUsageDetailButton, 0)

   Local $FHXPhaseParmListButton = GUICtrlCreateButton("PhaseParmList", 763, 522, 171, 30)
   GUICtrlSetCursor($FHXPhaseParmListButton, 0)

   Local $FHXPhaseRecParmUsageButton = GUICtrlCreateButton("PhaseRecParmUsage", 944, 522, 171, 30)
   GUICtrlSetCursor($FHXPhaseRecParmUsageButton, 0)

   Local $FHXPrepDiff1LineButton = GUICtrlCreateButton("PrepDiff1Line", 1125, 522, 171, 30)
   GUICtrlSetCursor($FHXPrepDiff1LineButton, 0)

   Local $FHXRecParamButton = GUICtrlCreateButton("RecParam", 220, 562, 171, 30)
   GUICtrlSetCursor($FHXRecParamButton, 0)

   Local $FHXRecTreeButton = GUICtrlCreateButton("RecTree", 401, 562, 171, 30)
   GUICtrlSetCursor($FHXRecTreeButton, 0)

   Local $FHXSearchButton = GUICtrlCreateButton("Search", 582, 562, 171, 30)
   GUICtrlSetCursor($FHXSearchButton, 0)

   Local $FHXSFCCheckButton = GUICtrlCreateButton("SFCCheck", 763, 562, 171, 30)
   GUICtrlSetCursor($FHXSFCCheckButton, 0)

   Local $FHXUnlinkInstConfigButton = GUICtrlCreateButton("UnlinkInstConfig", 944, 562, 171, 30)
   GUICtrlSetCursor($FHXUnlinkInstConfigButton, 0)

  ; VBScript Tools
   Local $FHXCMSummaryButton = GUICtrlCreateButton("CMSummary", 1125, 562, 171, 30)
   GUICtrlSetCursor($FHXCMSummaryButton, 0)

   Local $FHXRecipeSummaryButton = GUICtrlCreateButton("RecipeSummary", 220, 602, 171, 30)
   GUICtrlSetCursor($FHXRecipeSummaryButton, 0)

   Local $FHXUnitSummaryButton = GUICtrlCreateButton("UnitSummary", 401, 602, 171, 30)
   GUICtrlSetCursor($FHXUnitSummaryButton, 0)

   ; Excel Macros
   Local $FHXRecipeParamExtractButton = GUICtrlCreateButton("RecipeParamExtract", 582, 602, 171, 30)
   GUICtrlSetCursor($FHXRecipeParamExtractButton, 0)


   ; Tooltips===============================================================================================
   GUICtrlSetTip($fhxButton, "Open the FHXFiles folder.")
   GUICtrlSetTip($resultsButton, "Open the Results folder.")
   GUICtrlSetTip($fileButton, "Select FHX files to process.")

   ; Primary functionality =====================================================================================
   Local $tempDrive, $tempDir, $tempExtension
   Global $fileName

   ; Listen for button presses
   While 1
	  $msg = GUIGetMsg()
	  Switch $msg
		 Case $GUI_EVENT_CLOSE
			ExitLoop

		 ; Sidebar buttons
		 Case $fhxButton
			Run("explorer.exe " & $path & "\crti\fhxfiles")
		 Case $resultsButton
			OpenResults()
		 Case $readmeButton
			Run("explorer.exe " & $path & "\crti\readme.txt")

		 ; Child window functionality ===================================================================
		 Case $useCasesButton
			;MsgBox("", "CRTI", "This feature has not yet been implemented")
			; Use case function goes here
			GUISetState($GUI_DISABLE, $GUI)
            GUISetState(@SW_SHOW, $useCaseGUI)
            While 1
			   Switch GUIGetMsg()
				  Case $GUI_EVENT_CLOSE
					 ;GUICtrlSetData($applicableToolsLabel, "Applicable Tools: None")
					 GUISetState(@SW_HIDE, $useCaseGUI)
					 GUISetState($GUI_ENABLE, $GUI)
					 ExitLoop

				  ;Case $testButton
					 ;GUICtrlSetData($applicableToolsLabel, "Applicable Tools: Test 1")
			   EndSwitch
            WEnd
		 ; =============================================================================================

		 ; Main content buttons
		 Case $fileButton
			ChooseFile()
			; Update text to reflect file chosen
			If $fileChosen = "" Then
			   $fileDesc = "None"
			   $fileName = ""
			Else
			   ; Multiple files chosen
			   If StringInStr($fileChosen, ",") Then
				  $fileDesc = ""
				  Local $array[$ARR_SIZE] = StringSplit($fileChosen, ",")
				  Local $uniqueArray = _ArrayUnique($array)
				  $size = 0
				  Local $tempSize = UBound($uniqueArray) - 2

				  If $tempSize > 0 Then
					 $size = $tempSize
				  EndIf

				  For $i = 1 To $array[0]
					 _PathSplit($array[$i], $tempDrive, $tempDir, $fileName, $tempExtension)
					 $fileDesc &= $fileName & $tempExtension
					 If Not ($i = $size) Then
						$fileDesc &= ", "
					 EndIf
				  Next
			   ; Single file chosen
			   Else
				  _PathSplit($fileChosen, $tempDrive, $tempDir, $fileName, $tempExtension)
				  $fileDesc = $fileName & $tempExtension
			   EndIf
			EndIf

			If ($size = 1) Then
			   GUICtrlSetData($fileNameDisplay, "File Chosen: " & $fileDesc)
			ElseIf ($size = 0) Then
			   $fileDesc = "None"
			   GUICtrlSetData($fileNameDisplay, "File Chosen: " & $fileDesc)
			Else
			   GUICtrlSetData($fileNameDisplay, "Files Chosen: " & $fileDesc)
			EndIf

; CLI Tools=========================================================================================
		 Case $FHXAlarmConfigButton
			BaseCLIWrapper("FHXAlarmConfig", $fileChosen)
			;TestReturnVal()

		 Case $FHXAliasByUnitClassButton
			BaseCLIWrapper("FHXAliasByUnitClass", $fileChosen)

		 Case $FHXAnalogInfoButton
			BaseCLIWrapper("FHXAnalogInfo", $fileChosen)

		 Case $FHXAreaLibObjListButton
			BaseCLIWrapper("FHXAreaLibObjList", $fileChosen)

		 Case $FHXBulkTextProcButton
			BulkTextProcWrapper($fileChosen)

		 Case $FHXDiffToolButton
			DiffToolWrapper($fileChosen)

		 Case $FHXHistoryParamsButton
			BaseCLIWrapper("FHXHistoryParams", $fileChosen)

		 Case $FHXModuleInstanceParamsButton
			BaseCLIWrapper("FHXModuleInstanceParams", $fileChosen)

 		 Case $FHXModParamsBuildOPCButton
			ModParamsBuildOPCWrapper($fileChosen)

 		 Case $FHXModTagListButton
			BaseCLIWrapper("FHXModTagList", $fileChosen)

		 Case $FHXObjDateAndCheckOutButton
			BaseCLIWrapper("FHXObjDateAndCheckOut", $fileChosen)

		 Case $FHXObjDateAndVerButton
			BaseCLIWrapper("FHXObjDateAndVer", $fileChosen)

		 Case $FHXParamListButton
			BaseCLIWrapper("FHXParamList", $fileChosen)

		 Case $FHXParamMapButton
			BaseCLIWrapper("FHXParamMap", $fileChosen)

		 Case $FHXPhaseAliasUsageDetailButton
			BaseCLIWrapper("FHXPhaseAliasUsageDetail", $fileChosen)

		 Case $FHXPhaseParmListButton
			BaseCLIWrapper("FHXPhaseParmList", $fileChosen)

		 Case $FHXPhaseRecParmUsageButton
			BaseCLIWrapper("FHXPhaseRecParmUsage", $fileChosen)

		 Case $FHXPrepDiff1LineButton
			BaseCLIWrapper("FHXPrepDiff1Line", $fileChosen)

		 Case $FHXRecParamButton
			BaseCLIWrapper("FHXRecParam", $fileChosen)

		 Case $FHXRecTreeButton
			RecTreeWrapper($fileChosen)

		 Case $FHXSFCCheckButton
			BaseCLIWrapper("FHXSFCCheck", $fileChosen)

		 Case $FHXSearchButton
			FHXSearchWrapper($fileChosen)

		 Case $FHXUnlinkInstConfigButton
			FHXUnlinkInstConfigWrapper($fileChosen)

; DragNDrop Tools=============================================================
		 Case $FHXCMSummaryButton
			DragNDropToolWrapper("CMSummary", $fileChosen)

		 Case $FHXRecipeSummaryButton
			DragNDropToolWrapper("RecipeSummary", $fileChosen)

		 Case $FHXUnitSummaryButton
			DragNDropToolWrapper("UnitSummary", $fileChosen)

; Excel Macros===========================================================
		 Case $FHXRecipeParamExtractButton
			RecParamExtractWrapper($fileChosen)

	  EndSwitch
   WEnd

EndFunc

; Supplementary Functions ====================================================

; Opens results folder on \\OCN-PCSDEVDOC01\Reports\CRTI\Results
Func OpenResults()

   ; NOTE: DELETE THESE 4 LINES LATER
   ; Local $osv = @OSVersion
   ; Local $osBit = @OSArch
   ; MsgBox(0, "OS Details", "OS: " & $osv & @CRLF & "Bit: " & $osBit)
   ; Return

   ;==========================================================================================================

   CleanUp()
   Local $file = FileOpen($path & "\CRTI\Tools\openResults.bat", 2)
   FileWriteLine($file, $path & "\windows\system32\rundll32.exe shell32.dll,#61")
   FileClose($file)
   ShellExecute($path & "\CRTI\Tools\openResults.bat")

   WinActivate("Run")
   WinWaitActive("Run", "", 5)
   If Not WinActive("Run") Then
	  MsgBox("", "CRTI", "Something went wrong when opening the results folder.")
	  Return
   EndIf
   Send($server & "\CRTI\Results{ENTER}")

   ProcessWaitClose("cmd.exe")
   FileDelete($path & "\CRTI\Tools\openResults.bat")

EndFunc

; Save and close all cmd and Excel windows before continuing
Func CleanUp()

   ; Close all cmd windows
   If ProcessExists("cmd.exe") Then
	  While ProcessExists("cmd.exe")
		 ProcessClose("cmd.exe")
	  WEnd
   EndIf

   ; Save and close all Excel windows
   If ProcessExists("excel.exe") Then
	  Local $list = ProcessList("excel.exe")
	  Local $excelPID = $list[1][1]

	  Local $windowArray = _WinAPI_EnumProcessWindows($excelPID, False)
	  Local $arraySize = $windowArray[0][0]
	  Local $arrayStart = 2

	  For $i = $arrayStart To $arraySize - 17
		 Local $winHandle = $windowArray[$i][0]
		 Local $winTitle = WinGetTitle($winHandle)

		 Local $stringFound = StringInStr($winTitle, "Excel")
		 If $stringFound Then
			Local $oBook = _Excel_BookAttach($winTitle, "Title")
			If IsObj($oBook) Then
			   _Excel_BookClose($oBook)
			EndIf
		 EndIf
	  Next
	  ProcessClose("excel.exe")
   EndIf

EndFunc

; Fetches return value from text file
Func getReturnVal(ByRef $returnVal)

   $returnVal = FileRead($path & "\CRTI\returnVal.txt")
   FileDelete($path & "\CRTI\returnVal.txt")

EndFunc

; Modify error message based on return value
Func MakeErrorMsg($returnVal, ByRef $errorMsg)

   Switch $returnVal

	  Case 0
		 $errorMsg = "No error."
	  Case 1
		 $errorMsg = "The specified input file can't be opened. Check if it exists."
	  Case 2
		 $errorMsg = "Unspecified fatal error from the operating system"
	  Case 5
		 $errorMsg = "The specified output file can't be opened. Make sure it isn't open in another application."
	  Case 255
		 $errorMsg = "Interrupt detected (Ctrl-C or Ctrl-Break from the keyboard.)"
	  Case Else
		 $errorMsg = "Unknown error."

   EndSwitch

EndFunc

; Select the fhx file(s) to be processed
Func ChooseFile()
   $fileChosen = ""
   $fileChosen = FileOpenDialog("Select a file", $path & "\CRTI\FHXFiles", "FHX files (*.fhx)", BitOR($FD_FILEMUSTEXIST, $FD_PATHMUSTEXIST, $FD_MULTISELECT))
   If @error Then
	  ; Display error message
	  FileChangeDir(@ScriptDir)
   Else
	  Local $tempArray[$ARR_SIZE]
	  FileChangeDir(@ScriptDir)
	  $tempArray = StringSplit($fileChosen, "|")

	  Local $uniqueArray = _ArrayUnique($tempArray)
	  $size = 0
	  Local $tempSize = UBound($uniqueArray) - 2

	  If $tempSize > 0 Then
		 $size = $tempSize
	  EndIf

	  If $size > $ARR_SIZE + 1 Then
		 MsgBox($MB_OK, "CRTI", "Please choose a maximum of " & $ARR_SIZE & " files.")
		 $fileChosen = ""
		 Return
	  EndIf

	  ; Populate array with files if multiple files are chosen
	  If $size > 1 Then
		 $fileChosen = ""
		 $multiFileArray = $tempArray

		 For $i = 2 To $multiFileArray[0]
			$fileChosen &= $multiFileArray[1] & "\" & $multiFileArray[$i]
			If Not($i = $size) Then
			   $fileChosen &= ", "
			EndIf
		 Next
	  EndIf
   EndIf
EndFunc

; Creates and runs a batch file for normal tool format
Func CreateBatchFile($paramTool, $paramFilename)

   Local $drive, $dir, $name, $extension
   _PathSplit($paramFilename, $drive, $dir, $name, $extension)

   Local $file = FileOpen($path & "\CRTI\Tools\crtiFhxToRaw.bat", 2)

   ;FileWriteLine($file, "@echo off")
   FileWriteLine($file, "cls")
   FileWriteLine($file, "set tool=" & $paramTool)
   FileWriteLine($file, "set file=" & $paramFilename)
   FileWriteLine($file, "set name=" & $name)
   FileWriteLine($file, "set path=" & $path)
   FileWriteLine($file, 'call %path%\CRTI\Tools\CLITools\%tool%.exe "%file%" > "%path%\CRTI\TempResults\%tool%\%name%.csv"')
   FileWriteLine($file, "echo %ERRORLEVEL% > %path%\CRTI\returnVal.txt")

   FileClose($file)
   ShellExecute($path & "\CRTI\Tools\crtiFhxToRaw.bat")
   ProcessWaitClose("cmd.exe")

   FileDelete($path & "\CRTI\Tools\crtiFhxToRaw.bat")

EndFunc

; Convert csv files to xlsx, delete csv files, move xlsx results to server
Func FormatExcel ($paramTool, $paramFilename)

   Local $drive, $dir, $name, $extension
   _PathSplit($paramFilename, $drive, $dir, $name, $extension)

   Local $datetime = FileGetTime($paramFilename, 0, 1)
   $datetime = StringMid($datetime, 5, 2) & "_" & StringMid($datetime, 7, 2) & "_" & StringMid($datetime, 1, 4) & "_" & StringMid($datetime, 9, 2) & "_" & StringMid($datetime, 11, 2) & "_" & StringMid($datetime, 13, 2)

   Local $windowTitle = $name & " - Excel"
   Local $oExcel = _Excel_Open(True, False, True, True)  ; Works: FFTT
   Local $oBook = _Excel_BookOpen($oExcel, $path & '\CRTI\TempResults\' & $paramTool & '\' & $name & '.csv')

   ProcessWait("excel.exe")
   WinActivate($windowTitle)

   ; Save as xlsx instead of csv
   If isObj($oBook) Then
	  Send("{F12}")
	  Send("{TAB}")
	  Send("{DOWN}")
	  Send("o")
	  Send("e")
	  Send("{ENTER}")
	  Send("+{TAB}")
	  Send($name & "(" & $datetime & ")")
	  Send("{ENTER}")
	  Send("!y")

	  Local $delay = 2000 ;

	  Send("!{F4}")
	  Sleep($delay)
   EndIf
   ProcessClose("excel.exe")

   ;Properly size columns
   Local $oExcel2 = _Excel_Open(True, False, True, True) ; Works: FFTT
   Local $oBook2 = _Excel_BookOpen($oExcel2, $path & "\CRTI\TempResults\" & $paramTool & "\" & $name & "(" & $datetime & ").xlsx")

   ProcessWait("excel.exe")
   WinActivate($windowTitle)

   If IsObj($oBook2) Then
	  $oBook2.ActiveSheet.Columns("A:ZZ").AutoFit
   Else
	  MsgBox($MB_SYSTEMMODAL, "CRTI", "ERROR. Check that this tool can be used on the files you chose.")
   EndIf

   Send("^s")
   _Excel_Close($oExcel2)
   ProcessWaitClose("excel.exe")

   ; Move xlsx file to server and clean up files
   Local $file = FileOpen($path & "\CRTI\Tools\delFile.bat", 2)
   ;FileWriteLine($file, "@echo off")
   FileWriteLine($file, "cls")
   FileWriteLine($file, "set path=" & $path)
   FileWriteLine($file, "set server=" & $server)
   FileWriteLine($file, "set tool=" & $paramTool)
   FileWriteLine($file, "set name=" & $name)
   FileWriteLine($file, "set datetime=" & $datetime)
   FileWriteLine($file, 'C:\Windows\System32\robocopy %path%\CRTI\TempResults\%tool%\ %server%\CRTI\Results\%tool%\ "%name%(%datetime%).xlsx"')
   FileWriteLine($file, 'del "%path%\CRTI\TempResults\%tool%\%name%.csv"')
   FileWriteLine($file, 'del "%path%\CRTI\TempResults\%tool%\%name%(%datetime%).xlsx"')

   FileClose($file)
   ShellExecute($path & "\CRTI\Tools\delFile.bat")
   ProcessWaitClose("cmd.exe")
   FileDelete($path & "\CRTI\Tools\delFile.bat")

EndFunc

; Creates batch file and formats Excel spreadsheet for regular CLI Tools
Func BaseCLIWrapper($paramTool, $paramFilename)

   If $paramFilename = "" Then
	  MsgBox($MB_SYSTEMMODAL, "CRTI", "Please choose a file")
   Else
	  Local $msg = MsgBox($MB_OKCANCEL, "CRTI", "Start operation.")
	  If $msg = $IDCANCEL Then
		 MsgBox($MB_SYSTEMMODAL, "CRTI", "Operation cancelled.")
	  Else
		 CleanUp()
		 ; Single file case
		 If NOT StringInStr($paramFilename, ", ") Then
			CreateBatchFile($paramTool, $paramFilename)

			; Check for tool error
			Local $returnVal, $errorMsg
			getReturnVal($returnVal)
			If $returnVal = 0 Then
			   FormatExcel($paramTool, $paramFilename)
			   MsgBox($MB_SYSTEMMODAL, "CRTI", "Operation complete.")
			Else
			   MakeErrorMsg($returnVal, $errorMsg)
			   MsgBox($MB_SYSTEMMODAL, "CRTI", $errorMsg)
			   Return
			EndIf

		 Else ; Multiple files case
			Local $array[$ARR_SIZE] = StringSplit($paramFilename, ",")
			For $i = 1 To $array[0]

			   If StringLeft($array[$i], 1) = " " Then
				  $array[$i] = StringTrimLeft($array[$i], 1)
			   EndIf

			   ;MsgBox("", "", $i & ": " & $array[$i]) ; Displays each file chosen

			   CreateBatchFile($paramTool, $array[$i])
			   Local $returnVal, $errorMsg
			   getReturnVal($returnVal)
			   If $returnVal = 0 Then
				  FormatExcel($paramTool, $array[$i])
				  If $i = $array[0] Then
					 MsgBox($MB_SYSTEMMODAL, "CRTI", "Operation complete.")
				  EndIf
			   Else
				  MakeErrorMsg($returnVal, $errorMsg)
				  MsgBox($MB_SYSTEMMODAL, "CRTI", $errorMsg)
			   EndIf
			Next
		 EndIf
	  EndIf
   EndIf

EndFunc

; Run tool
Func RunRecTree($paramFilename)

   Local $drive, $dir, $name, $extension
   _PathSplit($paramFilename, $drive, $dir, $name, $extension)

   Local $datetime = FileGetTime($paramFilename, 0, 1)
   $datetime = StringMid($datetime, 5, 2) & "_" & StringMid($datetime, 7, 2) & "_" & StringMid($datetime, 1, 4) & "_" & StringMid($datetime, 9, 2) & "_" & StringMid($datetime, 11, 2) & "_" & StringMid($datetime, 13, 2)

   Local $file = FileOpen($path & "\CRTI\Tools\crtiFhxToRaw.bat", 2)

   FileWriteLine($file, "@echo off")
   FileWriteLine($file, "cls")
   FileWriteLine($file, "set file=" & $paramFilename)
   FileWriteLine($file, "set name=" & $name)
   FileWriteLine($file, "set path=" & $path)
   FileWriteLine($file, 'call %path%\CRTI\Tools\CLITools\FHXRecTree "%file%" > "%path%\CRTI\TempResults\FHXRecTree\%name%.txt"')
   FileWriteLine($file, "echo %ERRORLEVEL% > %path%\CRTI\returnVal.txt")

   FileClose($file)
   ShellExecute($path & "\CRTI\Tools\crtiFhxToRaw.bat")
   ProcessWaitClose("cmd.exe")
   FileDelete($path & "\CRTI\Tools\crtiFhxToRaw.bat")

   ; Check for tool error
   Local $returnVal, $errorMsg
   getReturnVal($returnVal)
   If Not ($returnVal = 0) Then
	  MakeErrorMsg($returnVal, $errorMsg)
	  MsgBox($MB_SYSTEMMODAL, "CRTI", $errorMsg)
	  Return
   EndIf

   ; Convert to xlsx from csv
   Local $windowTitle = $name & " - Excel"
   Local $oExcel = _Excel_Open(True, False, True, True)  ; Works: FFTT
   Local $oBook = _Excel_BookOpen($oExcel, $path & "\CRTI\TempResults\FHXRecTree\" & $name & ".txt")

   ProcessWait("excel.exe")
   WinActivate($windowTitle)

   If isObj($oBook) Then
	  Send("{F12}")
	  Send("{TAB}")
	  Send("{DOWN}")
	  Send("o")
	  Send("e")
	  Send("{ENTER}")
	  Send("+{TAB}")
	  Send($name & "(" & $datetime & ")")
	  Send("{ENTER}")
	  Send("!y")

	  Local $delay = 2000 ;

	  Send("!{F4}")
	  Sleep($delay)
   EndIf
   ProcessClose("excel.exe")

   ;Properly size columns
   Local $oExcel2 = _Excel_Open(True, False, True, True)
   Local $oBook2 = _Excel_BookOpen($oExcel2, $path & "\CRTI\TempResults\FHXRecTree\" & $name & "(" & $datetime & ").xlsx")

   ProcessWait("excel.exe")
   WinActivate($windowTitle)

   If IsObj($oBook2) Then
	  $oBook2.ActiveSheet.Columns("A:ZZ").AutoFit
   Else
	  MsgBox($MB_SYSTEMMODAL, "CRTI", "ERROR. Check that this tool can be used on the files you chose.")
   EndIf

   Send("^s")
   _Excel_Close($oExcel2)
   ProcessWaitClose("excel.exe")

   ; Move xlsx file to server and clean up files
   Local $file = FileOpen($path & "\CRTI\Tools\delFile.bat", 2)
   FileWriteLine($file, "@echo off")
   FileWriteLine($file, "cls")
   FileWriteLine($file, "set path=" & $path)
   FileWriteLine($file, "set server=" & $server)
   FileWriteLine($file, "set name=" & $name)
   FileWriteLine($file, "set datetime=" & $datetime)
   FileWriteLine($file, 'C:\Windows\System32\robocopy %path%\CRTI\TempResults\FHXRecTree\ %server%\CRTI\Results\FHXRecTree\ "%name%(%datetime%).xlsx"')
   FileWriteLine($file, 'del "%path%\CRTI\TempResults\FHXRecTree\%name%.txt"')
   FileWriteLine($file, 'del "%path%\CRTI\TempResults\FHXRecTree\%name%(%datetime%).xlsx"')
   FileClose($file)
   ShellExecute($path & "\CRTI\Tools\delFile.bat")
   ProcessWaitClose("cmd.exe")
   FileDelete($path & "\CRTI\Tools\delFile.bat")

EndFunc

Func RecTreeWrapper($paramFilename)

   If $fileName = "" Then
	  MsgBox($MB_SYSTEMMODAL, "CRTI", "Please choose a file")
   Else
	  Local $msg = MsgBox($MB_OKCANCEL, "CRTI", "Start operation.")
	  If $msg = $IDCANCEL Then
		 MsgBox($MB_SYSTEMMODAL, "CRTI", "Operation cancelled.")
	  Else
		 CleanUp()
		 If NOT StringInStr($paramFilename, ", ") Then
			RunRecTree($paramFilename)
			MsgBox($MB_SYSTEMMODAL, "CRTI", "Operation complete.")
		 Else
			Local $array[$ARR_SIZE] = StringSplit($paramFilename, ",")
			For $i = 1 To $array[0]
			   ;MsgBox("", "", $i & ": " & $array[$i]) ; Displays each file chosen

			   If StringLeft($array[$i], 1) = " " Then
				  $array[$i] = StringTrimLeft($array[$i], 1)
			   EndIf

			   RunRecTree($array[$i])
			   If $i = $array[0] Then
				  MsgBox($MB_SYSTEMMODAL, "CRTI", "Operation complete.")
			   EndIf
			Next
		 EndIf
	  EndIf
   EndIf

EndFunc

; Runs FHXBulkTextProc on specified file
Func RunBulkTextProc($paramFilename, $patternFile)

   Local $drive, $dir, $name, $extension
   _PathSplit($paramFilename, $drive, $dir, $name, $extension)

   Local $datetime = FileGetTime($paramFilename, 0, 1)
   $datetime = StringMid($datetime, 5, 2) & "_" & StringMid($datetime, 7, 2) & "_" & StringMid($datetime, 1, 4) & "_" & StringMid($datetime, 9, 2) & "_" & StringMid($datetime, 11, 2) & "_" & StringMid($datetime, 13, 2)

   Run("cmd.exe")
   WinWaitActive("C:\WINDOWS\SYSTEM32\cmd.exe", "", 10)
   If Not WinActive("C:\WINDOWS\SYSTEM32\cmd.exe") Then
	  MsgBox("", "CRTI", "Something went wrong when opening cmd")
	  Return
   EndIf

   Send('call ' & $path & '\CRTI\Tools\CLITools\FHXBulkTextProc.exe "' & $patternFile & '" "' & $paramFilename & '" > "' & $path & '\CRTI\TempResults\FHXBulkTextProc\' & $name & '(' & $datetime & ').fhx"')
   Send("{ENTER}")
   Sleep(1000)
   Send("n")
   Send("{ENTER}")
   Send("echo %ERRORLEVEL% > " & $path & "\CRTI\returnVal.txt")
   Send("{ENTER}")
   Send("exit")
   Send("{ENTER}")

   ProcessWaitClose("cmd.exe")

   ; Check for tool error
   Local $returnVal, $errorMsg
   getReturnVal($returnVal)
   If Not ($returnVal = 0) Then
	  MakeErrorMsg($returnVal, $errorMsg)
	  MsgBox($MB_SYSTEMMODAL, "CRTI", $errorMsg)
	  Return
   EndIf

   Local $file2 = FileOpen($path & "\CRTI\Tools\copyFile.bat", 2)
   FileWriteLine($file2, "set path=" & $path)
   FileWriteLine($file2, "set server=" & $server)
   FileWriteLine($file2, "set name=" & $name)
   FileWriteLine($file2, "set datetime=" & $datetime)
   FileWriteLine($file2, 'C:\Windows\System32\robocopy %path%\CRTI\TempResults\FHXBulkTextProc\ %server%\CRTI\Results\FHXBulkTextProc\ "%name%(%datetime%).fhx"')
   FileWriteLine($file2, 'del "%path%\CRTI\TempResults\FHXBulkTextProc\%name%(%datetime%).fhx"')

   FileClose($file2)
   ShellExecute($path & "\CRTI\Tools\copyFile.bat")
   ProcessWaitClose("cmd.exe")

   FileDelete($path & "\CRTI\Tools\runBulkTextProc.bat")
   FileDelete($path & "\CRTI\Tools\copyFile.bat")

EndFunc

; Wrapper that handles single and multiple file cases for BulkTextProc
Func BulkTextProcWrapper($paramFilename)

   If $fileName = "" Then
	  MsgBox($MB_SYSTEMMODAL, "CRTI", "Please choose a file")
   Else
	  Local $patternFile

	  Local $msg = MsgBox($MB_OKCANCEL, "CRTI", "Please choose the txt file containing the patterns to be deleted/replaced.")
	  If $msg = $IDCANCEL Then
		 MsgBox($MB_SYSTEMMODAL, "CRTI", "Operation cancelled.")
	  Else
		 CleanUp()
		 $patternFile = FileOpenDialog("Select a file", $path & "\CRTI\", "Text files (*.txt)", BitOR($FD_FILEMUSTEXIST, $FD_PATHMUSTEXIST))

		 If $patternFile = "" Then
			MsgBox($MB_SYSTEMMODAL, "CRTI", "User did not select a pattern file.")
			Return
		 EndIf

		 MsgBox($MB_OK, "CRTI", "Start operation.")
		 ;If $size = 1 Then
		 If NOT StringInStr($paramFilename, ", ") Then
			RunBulkTextProc($paramFilename, $patternFile)
			MsgBox($MB_SYSTEMMODAL, "CRTI", "Operation complete.")
		 Else
			Local $array[$ARR_SIZE] = StringSplit($paramFilename, ",")
			For $i = 1 To $array[0]
			   If StringLeft($array[$i], 1) = " " Then
				  $array[$i] = StringTrimLeft($array[$i], 1)
			   EndIf

			   RunBulkTextProc($array[$i], $patternFile)
			   If $i = $array[0] Then
				  MsgBox($MB_SYSTEMMODAL, "CRTI", "Operation complete.")
			   EndIf
			Next
		 EndIf
	  EndIf
   EndIf

EndFunc

; Run tool
Func RunFHXSearch($paramFilename, $paramText)

   Local $drive, $dir, $name, $extension
   _PathSplit($paramFilename, $drive, $dir, $name, $extension)

   Local $datetime = FileGetTime($paramFilename, 0, 1)
   $datetime = StringMid($datetime, 5, 2) & "_" & StringMid($datetime, 7, 2) & "_" & StringMid($datetime, 1, 4) & "_" & StringMid($datetime, 9, 2) & "_" & StringMid($datetime, 11, 2) & "_" & StringMid($datetime, 13, 2)

   Local $file = FileOpen($path & "\CRTI\Tools\runFHXSearch.bat", 2)

   FileWriteLine($file, "@echo off")
   FileWriteLine($file, "cls")
   FileWriteLine($file, "set file=" & $paramFilename)
   FileWriteLine($file, "set name=" & $name)
   FileWriteLine($file, "set text=" & $paramText)
   FileWriteLine($file, "set path=" & $path)
   FileWriteLine($file, "set datetime=" & $datetime)
   FileWriteLine($file, 'call %path%\CRTI\Tools\CLITools\FHXSearch "%file%" %text% > "%path%\CRTI\TempResults\FHXSearch\%name%(%datetime%).txt"')
   FileWriteLine($file, "echo %ERRORLEVEL% > %path%\CRTI\returnVal.txt")

   FileClose($file)
   ShellExecute($path & "\CRTI\Tools\runFHXSearch.bat")
   ProcessWaitClose("cmd.exe")

   ; Check for tool error
   Local $returnVal, $errorMsg
   getReturnVal($returnVal)
   If Not ($returnVal = 0) Then
	  MakeErrorMsg($returnVal, $errorMsg)
	  MsgBox($MB_SYSTEMMODAL, "CRTI", $errorMsg)
	  Return
   EndIf

   Local $file2 = FileOpen($path & "\CRTI\Tools\copyFile.bat", 2)
   FileWriteLine($file2, "set path=" & $path)
   FileWriteLine($file2, "set server=" & $server)
   FileWriteLine($file2, "set name=" & $name)
   FileWriteLine($file2, "set datetime=" & $datetime)
   FileWriteLine($file2, 'C:\Windows\System32\robocopy %path%\CRTI\TempResults\FHXSearch\ %server%\CRTI\Results\FHXSearch\ "%name%(%datetime%).txt"')
   FileWriteLine($file2, 'del "%path%\CRTI\TempResults\FHXSearch\%name%(%datetime%).txt"')
   FileClose($file2)
   ShellExecute($path & "\CRTI\Tools\copyFile.bat")
   ProcessWaitClose("cmd.exe")

   FileDelete($path & "\CRTI\Tools\runFHXSearch.bat")
   FileDelete($path & "\CRTI\Tools\copyFile.bat")

EndFunc

Func FHXSearchWrapper($paramFilename)

   If $fileName = "" Then
	  MsgBox($MB_SYSTEMMODAL, "CRTI", "Please choose a file")
   Else
	  Local $msg = MsgBox($MB_OKCANCEL, "CRTI", "Start operation.")

	  If $msg = $IDCANCEL Then
		 MsgBox($MB_SYSTEMMODAL, "CRTI", "Operation cancelled.")
	  Else
		 CleanUp()
		 Local $textPattern = InputBox("CRTI", "Enter the text or pattern to search for: ", "")
		 If $textPattern = "" Then
			MsgBox($MB_SYSTEMMODAL, "CRTI", "User did not enter a text pattern to search for.")
			Return
		 EndIf

		 ; Make sure text is enclosed within quotes
		 Local $firstChar = StringLeft($textPattern, 1)
		 Local $lastChar = StringRight($textPattern, 1)

		 If Not ($firstChar = '"') Then
			$textPattern = '"' & $textPattern
		 EndIf
		 If Not ($lastChar = '"') Then
			$textPattern = $textPattern & '"'
		 EndIf

		 ;If $size = 1 Then
		 If NOT StringInStr($paramFilename, ", ") Then
			RunFHXSearch($paramFilename, $textPattern)
			MsgBox($MB_SYSTEMMODAL, "CRTI", "Operation complete.")
		 Else
			Local $array[$ARR_SIZE] = StringSplit($paramFilename, ",")
			For $i = 1 To $array[0]
			   If StringLeft($array[$i], 1) = " " Then
				  $array[$i] = StringTrimLeft($array[$i], 1)
			   EndIf

			   RunFHXSearch($array[$i], $textPattern)
			   If $i = $array[0] Then
				  MsgBox($MB_SYSTEMMODAL, "CRTI", "Operation complete.")
			   EndIf
			Next
		 EndIf
	  EndIf
   EndIf

EndFunc

; Run tool
Func RunFHXUnlinkInstConfig($paramFilename)

   Local $drive, $dir, $name, $extension
   _PathSplit($paramFilename, $drive, $dir, $name, $extension)

   Local $datetime = FileGetTime($paramFilename, 0, 1)
   $datetime = StringMid($datetime, 5, 2) & "_" & StringMid($datetime, 7, 2) & "_" & StringMid($datetime, 1, 4) & "_" & StringMid($datetime, 9, 2) & "_" & StringMid($datetime, 11, 2) & "_" & StringMid($datetime, 13, 2)

   Local $file = FileOpen($path & "\CRTI\Tools\runUnlinkInstConfig.bat", 2)

   FileWriteLine($file, "@echo off")
   FileWriteLine($file, "cls")
   FileWriteLine($file, "set file=" & $paramFilename)
   FileWriteLine($file, "set name=" & $name)
   FileWriteLine($file, "set path=" & $path)
   FileWriteLine($file, "set datetime=" & $datetime)
   FileWriteLine($file, 'call %path%\CRTI\Tools\CLITools\FHXUnlinkInstConfig "%file%" > "%path%\CRTI\TempResults\FHXUnlinkInstConfig\%name%(%datetime%).fhx"')
   FileWriteLine($file, "echo %ERRORLEVEL% > %path%\CRTI\returnVal.txt")

   FileClose($file)
   ShellExecute($path & "\CRTI\Tools\runUnlinkInstConfig.bat")
   ProcessWaitClose("cmd.exe")

   ; Check for tool error
   Local $returnVal, $errorMsg
   getReturnVal($returnVal)
   If Not ($returnVal = 0) Then
	  MakeErrorMsg($returnVal, $errorMsg)
	  MsgBox($MB_SYSTEMMODAL, "CRTI", $errorMsg)
	  Return
   EndIf

   Local $file2 = FileOpen($path & "\CRTI\Tools\copyFile.bat", 2)
   FileWriteLine($file2, "set path=" & $path)
   FileWriteLine($file2, "set server=" & $server)
   FileWriteLine($file2, "set name=" & $name)
   FileWriteLine($file2, "set datetime=" & $datetime)
   FileWriteLine($file2, 'C:\Windows\System32\robocopy %path%\CRTI\TempResults\FHXUnlinkInstConfig\ %server%\CRTI\Results\FHXUnlinkInstConfig\ "%name%(%datetime%).fhx"')
   FileWriteLine($file2, 'del "%path%\CRTI\TempResults\FHXUnlinkInstConfig\%name%(%datetime%).fhx"')
   FileClose($file2)
   ShellExecute($path & "\CRTI\Tools\copyFile.bat")
   ProcessWaitClose("cmd.exe")

   FileDelete($path & "\CRTI\Tools\runUnlinkInstConfig.bat")
   FileDelete($path & "\CRTI\Tools\copyFile.bat")

EndFunc

Func FHXUnlinkInstConfigWrapper($paramFilename)

   If $fileName = "" Then
	  MsgBox($MB_SYSTEMMODAL, "CRTI", "Please choose a file")
   Else
	  Local $msg = MsgBox($MB_OKCANCEL, "CRTI", "Start operation.")

	  If $msg = $IDCANCEL Then
		 MsgBox($MB_SYSTEMMODAL, "CRTI", "Operation cancelled.")
	  Else
		 CleanUp()
		 If NOT StringInStr($paramFilename, ", ") Then
			RunFHXUnlinkInstConfig($paramFilename)
			MsgBox($MB_SYSTEMMODAL, "CRTI", "Operation complete.")
		 Else
			Local $array[$ARR_SIZE] = StringSplit($paramFilename, ",")
			For $i = 1 To $array[0]
			   If StringLeft($array[$i], 1) = " " Then
				  $array[$i] = StringTrimLeft($array[$i], 1)
			   EndIf

			   RunFHXUnlinkInstConfig($array[$i])
			   If $i = $array[0] Then
				  MsgBox($MB_SYSTEMMODAL, "CRTI", "Operation complete.")
			   EndIf
			Next
		 EndIf
	  EndIf
   EndIf

EndFunc

; Run tool
Func RunModParamsBuildOPC($paramFilename)

   Local $drive, $dir, $name, $extension
   _PathSplit($paramFilename, $drive, $dir, $name, $extension)

   Local $datetime = FileGetTime($paramFilename, 0, 1)
   $datetime = StringMid($datetime, 5, 2) & "_" & StringMid($datetime, 7, 2) & "_" & StringMid($datetime, 1, 4) & "_" & StringMid($datetime, 9, 2) & "_" & StringMid($datetime, 11, 2) & "_" & StringMid($datetime, 13, 2)

   Local $file = FileOpen($path & "\CRTI\Tools\runModParamsBuildOPC.bat", 2)

   FileWriteLine($file, "@echo off")
   FileWriteLine($file, "cls")
   FileWriteLine($file, "set file=" & $paramFilename)
   FileWriteLine($file, "set path=" & $path)
   FileWriteLine($file, 'call %path%\CRTI\Tools\CLITools\FHXModParamsBuildOPC "%file%"')
   FileWriteLine($file, "echo %ERRORLEVEL% > %path%\CRTI\returnVal.txt")

   FileClose($file)
   ShellExecute($path & "\CRTI\Tools\runModParamsBuildOPC.bat")
   ProcessWaitClose("cmd.exe")
   FileDelete($path & "\CRTI\Tools\runModParamsBuildOPC.bat")

   ; Check for tool error
   Local $returnVal, $errorMsg
   getReturnVal($returnVal)
   If Not ($returnVal = 0) Then
	  MakeErrorMsg($returnVal, $errorMsg)
	  MsgBox($MB_SYSTEMMODAL, "CRTI", $errorMsg)
	  Return
   EndIf

   ; Save as xlsx instead of csv
   Local $windowTitle = $name & " - Excel"
   Local $oExcel = _Excel_Open(True, False, True, True)  ; Works: FFTT
   Local $oBook = _Excel_BookOpen($oExcel, $path & "\CRTI\FHXFiles\" & $name & ".txt")

   ProcessWait("excel.exe")
   WinActivate($windowTitle)

   If isObj($oBook) Then
	  Send("{F12}")
	  Send("{TAB}")
	  Send("{DOWN}")
	  Send("o")
	  Send("e")
	  Send("{ENTER}")
	  Send("+{TAB}")
	  Send($name & "(" & $datetime & ")")
	  Send("{ENTER}")
	  Send("!y")

	  Local $delay = 2000 ;

	  Send("!{F4}")
	  Sleep($delay)
   EndIf
   ProcessClose("excel.exe")

   ; Copy temp result to user folder
   Local $file2 = FileOpen($path & "\CRTI\Tools\delFile.bat", 2)

   ;FileWriteLine($file2, "@echo off")
   FileWriteLine($file2, "cls")
   FileWriteLine($file2, "set path=" & $path)
   FileWriteLine($file2, "set name=" & $name)
   FileWriteLine($file2, "set datetime=" & $datetime)
   FileWriteLine($file2, 'C:\Windows\System32\robocopy %path%\CRTI\FHXFiles\ %path%\CRTI\TempResults\FHXModParamsBuildOPC\ "%name%(%datetime%).xlsx"')

   FileClose($file2)
   ShellExecute($path & "\CRTI\Tools\delFile.bat")
   ProcessWaitClose("cmd.exe")
   FileDelete($path & "\CRTI\Tools\delFile.bat")

   Local $oExcel2 = _Excel_Open(True, False, True, True) ; Works: FFTT
   Local $oBook2 = _Excel_BookOpen($oExcel2, $path & "\CRTI\TempResults\FHXModParamsBuildOPC\" & $name & "(" & $datetime & ").xlsx")

   ProcessWait("excel.exe")
   WinActivate($windowTitle)

   If IsObj($oBook2) Then
	  ; Properly size columns
	  $oBook2.ActiveSheet.Columns("A:ZZ").AutoFit
   Else
	  MsgBox($MB_SYSTEMMODAL, "CRTI", "ERROR. Check that this tool can be used on the files you chose.")
   EndIf

   Send("^s")
   _Excel_Close($oExcel2)
   ProcessWaitClose("excel.exe")

   ; Copy result to server and delete extra files
   Local $file3 = FileOpen($path & "\CRTI\Tools\movFile.bat", 2)
   FileWriteLine($file3, "set path=" & $path)
   FileWriteLine($file3, "set server=" & $server)
   FileWriteLine($file3, "set name=" & $name)
   FileWriteLine($file3, "set datetime=" & $datetime)
   FileWriteLine($file3, 'C:\Windows\System32\robocopy %path%\CRTI\TempResults\FHXModParamsBuildOPC\ %server%\CRTI\Results\FHXModParamsBuildOPC\ "%name%(%datetime%).xlsx"')
   FileWriteLine($file3, 'del "%path%\CRTI\TempResults\FHXModParamsBuildOPC\%name%(%datetime%).xlsx"')
   FileWriteLine($file3, 'del "%path%\CRTI\FHXFiles\%name%.txt"')
   FileWriteLine($file3, 'del "%path%\CRTI\FHXFiles\%name%(%datetime%).xlsx"')
   FileClose($file3)
   ShellExecute($path & "\CRTI\Tools\movFile.bat")
   ProcessWaitClose("cmd.exe")
   FileDelete($path & "\CRTI\Tools\movFile.bat")

EndFunc

Func ModParamsBuildOPCWrapper($paramFilename)

   If $fileName = "" Then
	  MsgBox($MB_SYSTEMMODAL, "CRTI", "Please choose a file")
   Else
	  Local $msg = MsgBox($MB_OKCANCEL, "CRTI", "Start operation.")

	  If $msg = $IDCANCEL Then
		 MsgBox($MB_SYSTEMMODAL, "CRTI", "Operation cancelled.")
	  Else
		 CleanUp()
		 If NOT StringInStr($paramFilename, ", ") Then
			RunModParamsBuildOPC($paramFilename)
			MsgBox($MB_SYSTEMMODAL, "CRTI", "Operation complete.")
		 Else
			Local $array[$ARR_SIZE] = StringSplit($paramFilename, ",")
			For $i = 1 To $array[0]
			   If StringLeft($array[$i], 1) = " " Then
				  $array[$i] = StringTrimLeft($array[$i], 1)
			   EndIf

			   RunModParamsBuildOPC($array[$i])
			   If $i = $array[0] Then
				  MsgBox($MB_SYSTEMMODAL, "CRTI", "Operation complete.")
			   EndIf
			Next
		 EndIf
	  EndIf
   EndIf

EndFunc

; Run tool
Func RunDragNDropTool($paramTool, $paramFilename) ; NOTE ================= NEED TO INCLUDE DATETIME SOMEHOW

   Local $drive, $dir, $name, $extension
   _PathSplit($paramFilename, $drive, $dir, $name, $extension)

   Local $datetime = FileGetTime($paramFilename, 0, 1)
   $datetime = StringMid($datetime, 5, 2) & "_" & StringMid($datetime, 7, 2) & "_" & StringMid($datetime, 1, 4) & "_" & StringMid($datetime, 9, 2) & "_" & StringMid($datetime, 11, 2) & "_" & StringMid($datetime, 13, 2)

   MsgBox("", "CRTI", "If a security warning to open the file appears, please open it.")

   Local $file1 = FileOpen($path & "\CRTI\Tools\rundnd.bat", 2)
   FileWriteLine($file1, "@echo off")
   FileWriteLine($file1, "cls")
   FileWriteLine($file1, "set path=" & $path)
   FileWriteLine($file1, "set tool=" & $paramTool)
   FileWriteLine($file1, "set name=" & $name)
   FileWriteLine($file1, 'C:\Windows\System32\robocopy %path%\CRTI\FHXFiles\ %path%\CRTI\Tools\DragNDropTools\%tool%\ "%name%.fhx"')
   FileWriteLine($file1, '%path%\CRTI\Tools\DragNDropTools\%tool%\tool.vbs "%path%\CRTI\Tools\DragNDropTools\%tool%\%name%.fhx"')
   FileClose($file1)
   ShellExecute($path & "\CRTI\Tools\rundnd.bat")

   ProcessWaitClose("cmd.exe")
   FileDelete($path & "\CRTI\Tools\rundnd.bat")

   Sleep(1500)

   WinWaitActive("Open File - Security Warning", "Do you want to open this file?", 10)
   Send("o")

   Sleep(2000)

   ; Loop until process is done (signaled by doneProcessing.txt)
   Local $file = FileOpen($path & "\CRTI\Tools\checkForFile.bat", 2)

   FileWriteLine($file, "@echo off")
   FileWriteLine($file, "set path=" & $path)
   FileWriteLine($file, "set lookForFile=%path%\CRTI\doneProcessing.txt")
   FileWriteLine($file, "echo Processing...")
   FileWriteLine($file, ":CheckForFile")
   FileWriteLine($file, "IF EXIST %lookForFile% GOTO FoundIt")
   FileWriteLine($file, "TIMEOUT /T 3 >nul")
   FileWriteLine($file, "GOTO CheckForFile")
   FileWriteLine($file, ":FoundIt")
   FileWriteLine($file, "ECHO Found: %lookForFile%")
   FileWriteLine($file, "exit")

   FileClose($file)

   WinWait("Blank Page - Internet Explorer", "", 5)
   If WinExists("Blank Page - Internet Explorer") Then
	  ShellExecute($path & "\CRTI\Tools\checkForFile.bat")
	  ProcessWaitClose("cmd.exe")
   Else
	  MsgBox("", "CRTI", "Something went wrong when launching the tool.")
	  Return
   EndIf
   FileDelete($path & "\CRTI\Tools\checkForFile.bat")

   Sleep(1000)
   ; Close vbs window
   WinClose("Blank Page - Internet Explorer")

   ; NOTE ================================================================ Reformat to batch script for consistency?
   ; Move files to server and delete extra files
   Run("cmd.exe", "", @SW_SHOW, $RUN_CREATE_NEW_CONSOLE)
   WinWait($path & "\WINDOWS\SYSTEM32\cmd.exe", "", 5)
   If WinExists($path & "\WINDOWS\SYSTEM32\cmd.exe") Then
	  WinActivate($path & "\WINDOWS\SYSTEM32\cmd.exe")
   Else
	  MsgBox("", "CRTI", "Something went wrong when opening the cmd prompt")
	  Return
   EndIf

   If $paramTool = "CMSummary" Or $paramTool = "UnitSummary" Then
	  Send('C:\Windows\System32\robocopy{SPACE}' & $path & '\CRTI\Tools\DragNDropTools\' & $paramTool & '\ ' & $server & '\CRTI\Results\FHX' & $paramTool & '\ "' & $name & '-out.xml"')
	  Send("{ENTER}")
	  Send('del{SPACE}' & $path & '\CRTI\Tools\DragNDropTools\' & $paramTool & '\"' & $name & '-out.xml"')
	  Send("{ENTER}")
	  Send('RENAME "' & $server & '\CRTI\Results\FHX' & $paramTool & '\' & $name & '-out.xml" "' & $name & '-out(' & $datetime & ').xml"')
	  Send("{ENTER}")
   ElseIf $paramTool = "RecipeSummary" Then
	  Send('C:\Windows\System32\robocopy{SPACE}' & $path & '\CRTI\Tools\DragNDropTools\' & $paramTool & '\ ' & $server & '\CRTI\Results\FHX' & $paramTool & '\ "' & $name & '-PHASES.xml"')
	  Send("{ENTER}")
	  Send('C:\Windows\System32\robocopy{SPACE}' & $path & '\CRTI\Tools\DragNDropTools\' & $paramTool & '\ ' & $server & '\CRTI\Results\FHX' & $paramTool & '\ "' & $name & '-RECIPES.xml"')
	  Send("{ENTER}")
	  Send('del{SPACE}"' & $path & '\CRTI\Tools\DragNDropTools\' & $paramTool & '\' & $name & '-PHASES.xml"')
	  Send("{ENTER}")
	  Send('del{SPACE}"' & $path & '\CRTI\Tools\DragNDropTools\' & $paramTool & '\' & $name & '-RECIPES.xml"')
	  Send("{ENTER}")
	  Send('RENAME "' & $server & '\CRTI\Results\FHX' & $paramTool & '\' & $name & '-PHASES.xml" "' & $name & '(' & $datetime & ')-PHASES.xml"')
	  Send("{ENTER}")
	  Send('RENAME "' & $server & '\CRTI\Results\FHX' & $paramTool & '\' & $name & '-RECIPES.xml" "' & $name & '(' & $datetime & ')-RECIPES.xml"')
	  Send("{ENTER}")
   EndIf
   Send("del{SPACE}" & $path & "\CRTI\doneProcessing.txt")
   Send("{ENTER}")
   Send('del{SPACE}"' & $path & '\CRTI\Tools\DragNDropTools\' & $paramTool & '\' & $name & '.fhx"')
   Send("{ENTER}")
   Send('del{SPACE}"' & $path & '\CRTI\Tools\DragNDropTools\' & $paramTool & '\temp\' & $name & '.xml"')
   Send("{ENTER}")
   Send("exit")
   Send("{ENTER}")

   ProcessWaitClose("cmd.exe")

EndFunc

Func DragNDropToolWrapper($paramTool, $paramFilename)

   If $fileName = "" Then
	  MsgBox($MB_SYSTEMMODAL, "CRTI", "Please choose a file")
   Else
	  Local $msg = MsgBox($MB_OKCANCEL, "CRTI", "Start operation.")

	  If $msg = $IDCANCEL Then
		 MsgBox($MB_SYSTEMMODAL, "CRTI", "Operation cancelled.")
	  Else
		 CleanUp()
		 If NOT StringInStr($paramFilename, ", ") Then
			RunDragNDropTool($paramTool, $paramFilename)
			MsgBox($MB_SYSTEMMODAL, "CRTI", "Operation complete.")
		 Else
			Local $array[$ARR_SIZE] = StringSplit($paramFilename, ",")
			For $i = 1 To $array[0]
			   If StringLeft($array[$i], 1) = " " Then
				  $array[$i] = StringTrimLeft($array[$i], 1)
			   EndIf

			   RunDragNDropTool($paramTool, $array[$i])
			   If $i = $array[0] Then
				  MsgBox($MB_SYSTEMMODAL, "CRTI", "Operation complete.")
			   EndIf
			Next
		 EndIf
	  EndIf
   EndIf

EndFunc

; Run tool
Func RunDiffTool($paramFile1, $paramFile2)

   Local $drive1, $dir1, $name1, $extension1
   _PathSplit($paramFile1, $drive1, $dir1, $name1, $extension1)
   Local $datetime1 = FileGetTime($paramFile1, 0, 1)
   $datetime1 = StringMid($datetime1, 5, 2) & "_" & StringMid($datetime1, 7, 2) & "_" & StringMid($datetime1, 1, 4) & "_" & StringMid($datetime1, 9, 2) & "_" & StringMid($datetime1, 11, 2) & "_" & StringMid($datetime1, 13, 2)

   Local $drive2, $dir2, $name2, $extension2
   _PathSplit($paramFile2, $drive2, $dir2, $name2, $extension2)
   Local $datetime2 = FileGetTime($paramFile2, 0, 1)
   $datetime2 = StringMid($datetime2, 5, 2) & "_" & StringMid($datetime2, 7, 2) & "_" & StringMid($datetime2, 1, 4) & "_" & StringMid($datetime2, 9, 2) & "_" & StringMid($datetime2, 11, 2) & "_" & StringMid($datetime2, 13, 2)

   Run("cmd.exe")
   WinWaitActive("C:\WINDOWS\SYSTEM32\cmd.exe", "", 10)
   If Not WinActive("C:\WINDOWS\SYSTEM32\cmd.exe") Then
	  MsgBox("", "CRTI", "Something went wrong when opening cmd")
	  Return
   EndIf

   Send('call ' & $path & '\CRTI\Tools\CLITools\ShowDiff.bat "' & $paramFile1 & '" "' & StringTrimLeft($paramFile2, 0) & '" > "' & $path & '\CRTI\TempResults\FHXDiffTool\' & $name1 & '(' & $datetime1 & ')-' & $name2 & '(' & $datetime2 & ')-diff.txt"')
   Send("{ENTER}")

   ProcessWaitClose("cmd.exe")

   ; NOTE ============================== currently terminates before able to record return value
   ; Check for tool error
   Local $returnVal, $errorMsg
   getReturnVal($returnVal)
   If Not ($returnVal = 0) Then
	  MakeErrorMsg($returnVal, $errorMsg)
	  MsgBox($MB_SYSTEMMODAL, "CRTI", $errorMsg)
	  Return
   EndIf

   Local $file2 = FileOpen($path & "\CRTI\Tools\copyFile.bat", 2)
   FileWriteLine($file2, "set path=" & $path)
   FileWriteLine($file2, "set server=" & $server)
   FileWriteLine($file2, "set name1=" & $name1)
   FileWriteLine($file2, "set name2=" & $name2)
   FileWriteLine($file2, "set datetime1=" & $datetime1)
   FileWriteLine($file2, "set datetime2=" & $datetime2)
   FileWriteLine($file2, 'C:\Windows\System32\robocopy %path%\CRTI\TempResults\FHXDiffTool\ %server%\CRTI\Results\FHXDiffTool\ "%name1%(%datetime1%)-%name2%(%datetime2%)-diff.txt"')
   FileWriteLine($file2, 'del "%path%\CRTI\TempResults\FHXDiffTool\%name1%(%datetime1%)-%name2%(%datetime2%)-diff.txt"')

   FileClose($file2)
   ShellExecute($path & "\CRTI\Tools\copyFile.bat")
   ProcessWaitClose("cmd.exe")

   FileDelete($path & "\CRTI\Tools\copyFile.bat")

EndFunc

; NOTE =================== Currently only takes 2 files
Func DiffToolWrapper($paramFiles)

   If $fileName = "" Then
	  MsgBox($MB_SYSTEMMODAL, "CRTI", "Please choose two files.")
   Else
	  Local $array[$ARR_SIZE] = StringSplit($paramFiles, ",")
	  If NOT ($array[0] = 2) Then
		 MsgBox($MB_SYSTEMMODAL, "CRTI", "Please choose two files.")
	  Else
		 Local $msg = MsgBox($MB_OKCANCEL, "CRTI", "Start operation.")

		 If $msg = $IDCANCEL Then
			MsgBox($MB_SYSTEMMODAL, "CRTI", "Operation cancelled.")
		 Else
			CleanUp()
			If StringLeft($array[2], 1) = " " Then
			   $array[2] = StringTrimLeft($array[2], 1)
			EndIf

			Local $paramFile1 = $array[1]
			Local $paramFile2 = $array[2]
			RunDiffTool($paramFile1, $paramFile2)
			MsgBox($MB_SYSTEMMODAL, "CRTI", "Operation complete.")
		 EndIf
	  EndIf
   EndIf

EndFunc

; Run tool
Func RunRecParamExtract($paramFilename)

   Local $drive, $dir, $name, $extension
   _PathSplit($paramFilename, $drive, $dir, $name, $extension)

   Local $datetime = FileGetTime($paramFilename, 0, 1)
   $datetime = StringMid($datetime, 5, 2) & "_" & StringMid($datetime, 7, 2) & "_" & StringMid($datetime, 1, 4) & "_" & StringMid($datetime, 9, 2) & "_" & StringMid($datetime, 11, 2) & "_" & StringMid($datetime, 13, 2)

   Local $oExcel = _Excel_Open(True, False, True, True)
   Local $oBook = _Excel_BookOpen($oExcel, $path & "\CRTI\Tools\ExcelMacros\RecipeParamExtraction.xlsm")

   ProcessWait("excel.exe")
   WinActivate("RecipeParamExtraction - Excel")
   WinWaitActive("RecipeParamExtraction - Excel", "", 10)
   If Not WinActive("RecipeParamExtraction - Excel") Then
	  MsgBox("", "CRTI", "Something went wrong when running the tool.")
	  Return
   EndIf

   ; Navigate to FHX sheet in Excel
   Send("{F5}")
   Send("fhx")
   Send("{LSHIFT down}1")
   Send("{LSHIFT up}")
   Send("a1{ENTER}")

   ; Copy contents of fhx file from notepad
   Run("notepad.exe " & $paramFilename, "", @SW_SHOW, $RUN_CREATE_NEW_CONSOLE)
   ProcessWait("notepad.exe")
   WinActivate($name & " - Notepad")
   WinWaitActive($name & " - Notepad", "", 10)
   If Not WinActive($name & " - Notepad") Then
	  MsgBox("", "CRTI", "Something went wrong when running the tool.")
	  Return
   EndIf
   Send("^a")
   Send("^c")
   ProcessClose("notepad.exe")
   ProcessWaitClose("notepad.exe", 5)

   ; Paste contents from fhx file into column A
   WinActivate("RecipeParamExtraction - Excel")
   WinWaitActive("RecipeParamExtraction - Excel", "", 10)
   If Not WinActive("RecipeParamExtraction - Excel") Then
	  MsgBox("", "CRTI", "Something went wrong when running the tool.")
	  Return
   EndIf
   Send("^g")
   Send("a1{ENTER}")
   Send("^{SPACE}")
   Send("^v")

   ; Activate macro
   Send("!wmv{TAB}{DOWN}{ENTER}")

   Local $doneFlag = 0
   Local $oSheet = $oBook.Sheets(1)

   While Not ($doneFlag)

	  _Excel_RangeWrite($oBook, $oSheet, "done", "ZZ1")
	  Local $reading = _Excel_RangeRead($oBook, $oSheet, "ZZ1")
	  If $reading = "done" Then
		 $doneFlag = 1
	  EndIf

   WEnd

   Send("^s")
   Sleep(2000)

   ; create results excel file
   Local $file = FileOpen($path & "\CRTI\Tools\createExcel.bat", 2)
   FileWriteLine($file, "@echo off")
   FileWriteLine($file, "cls")
   FileWriteLine($file, "set path=" & $path)
   FileWriteLine($file, "set name=" & $name)
   FileWriteLine($file, "set datetime=" & $datetime)
   FileWriteLine($file, 'copy %path%\CRTI\TempResults\FHXRecipeParamExtraction\blankWorksheet_(DO_NOT_DELETE).xlsx "%path%\CRTI\TempResults\FHXRecipeParamExtraction\%name%(%datetime%).xlsx"')
   FileClose($file)

   ShellExecute($path & "\CRTI\Tools\createExcel.bat", "")
   ProcessWaitClose("cmd.exe")
   FileDelete($path & "\CRTI\Tools\createExcel.bat")

   Local $oBook2 = _Excel_BookOpen($oExcel, $path & "\CRTI\TempResults\FHXRecipeParamExtraction\" & $name & "(" & $datetime & ").xlsx")
   Local $oCopiedSheet = _Excel_SheetCopyMove($oBook, $oBook.Sheets(2), $oBook2)

   $oCopiedSheet.Name = "Results"

   _Excel_Close($oExcel)

   ; Move file to server results folder
   Local $file2 = FileOpen($path & "\CRTI\Tools\movFile.bat", 2)
   FileWriteLine($file2, "set path=" & $path)
   FileWriteLine($file2, "set server=" & $server)
   FileWriteLine($file2, "set name=" & $name)
   FileWriteLine($file2, "set datetime=" & $datetime)
   FileWriteLine($file2, 'C:\Windows\System32\robocopy %path%\CRTI\TempResults\FHXRecipeParamExtraction\ %server%\CRTI\Results\FHXRecipeParamExtraction\ "%name%(%datetime%).xlsx"')
   FileWriteLine($file2, 'del "%path%\CRTI\TempResults\FHXRecipeParamExtraction\%name%(%datetime%).xlsx"')
   FileClose($file2)
   ShellExecute($path & "\CRTI\Tools\movFile.bat")
   ProcessWaitClose("cmd.exe")
   FileDelete($path & "\CRTI\Tools\movFile.bat")

EndFunc

Func RecParamExtractWrapper($paramFilename)

   If $fileName = "" Then
	  MsgBox($MB_SYSTEMMODAL, "CRTI", "Please choose a file")
   Else
	  Local $msg = MsgBox($MB_OKCANCEL, "CRTI", "Start operation.")

	  If $msg = $IDCANCEL Then
		 MsgBox($MB_SYSTEMMODAL, "CRTI", "Operation cancelled.")
	  Else
		 CleanUp()
		 If NOT StringInStr($paramFilename, ", ") Then
			RunRecParamExtract($paramFilename)
			MsgBox($MB_SYSTEMMODAL, "CRTI", "Operation complete.")
		 Else
			Local $array[$ARR_SIZE] = StringSplit($paramFilename, ",")
			For $i = 1 To $array[0]
			   If StringLeft($array[$i], 1) = " " Then
				  $array[$i] = StringTrimLeft($array[$i], 1)
			   EndIf

			   RunRecParamExtract($array[$i])
			   If $i = $array[0] Then
				  MsgBox($MB_SYSTEMMODAL, "CRTI", "Operation complete.")
			   EndIf
			Next
		 EndIf
	  EndIf
   EndIf

EndFunc
