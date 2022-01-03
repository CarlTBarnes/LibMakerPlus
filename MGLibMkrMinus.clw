     PROGRAM 
     
ReleaseInfoString EQUATE('2020.9.6 by Carl Barnes')     !Appears on Info/About Window 

![ ] Add a column with Export Forward info e.g. ImageHlp.DLL has Exports that forward to  DebugHlp.DLL. Probably do NOT want some forwards.
![ ] WinSxS Finder/Picker Window. Enter File Name. Scan subfolders named X86. Needed for ComCtrl v6. Use CWA Directory code gen
!Region Recent Changes
!Updated 2020-September By Carl Barnes
! - Copy button on toolbar to put list of symbols on clipboard so do not have to write EXP. Can do Tagged, Untagged, Filtered, Etc.
! - Search window results new Right-click Popup(Copy | Delete). CtrlC copies symbol name.
! - Exports Rt Click Popup() first line has "Symbol in Module" in bold so can tell in a long list the Module i.e. [31763(700)]Symbol "FilteredQ.symbol" 
! - Filtered List shows the Module in a column
! - Info / About window shows EXE Path. Moved Info ? button to far right. Arnor's Icetips email.
! - New ReRun button runs another instance of Libmaker
! - New SelectExportQ() replaces Select(?List1,x) so first Expands Tree level 1 then selects symbol line. Now select row when collapsed works and Vscroll shows. Prev if contracted did not work well
! - Tag Popup menu offers Tag Duplicates. Not typically needed, well there are a lot more Export Forwards
! - SkipAddExportQ() that reports Duplicates and Problem Symbol names now offers a "Tag All" so can see in List
! - DB() to OutputDebugString 
!
!Updated 2019-September By Carl Barnes
! - Right-Click on Add or subtract does POPUP of System32, Clarion 
! - Commandline "loadclawin32" will load the Clarion Win32.LIB DLLs e.g. Kernel32 User32 GDI32
! - Refactoring, Bugs fixed where code failed to check if Filter tab was selected, like Delete key
! - Subtract Button (cut icon) to Subtract Clarion's Win32.LIB so no Dups.
!     Also to Subtract a Win 7 DLL from Win 10 to see what's new
!     Dups can be Deleted or Tagged or Untagged.  See SubtractFileLIB()
! - Window tweaks formatting. Use Segoe UI and larger size. Align counts.
! - Locator on Window for All with Next/Prev button
! - Sorting preserves selected line.
! - SYSTEM{PROP:PropVScroll}=1 for better thumb
! - SYSTEM{PROP:MsgModeDefault}=MSGMODE:CANCOPY so all messages allow copy
! - Tagging individual lines for filter tab toggled by Double-Click or Enter key on All Tag
! - Delete key on Filtered tab untags on All list, delets from Filtered tab
! - Tag button on toolbar allows Tag /Untag All
! - SelectQBE Search String window enhanced for Ends with, Wildcard, RegEx, Case sensitive
!      Shows List of Symbols that match and allows Tag, Untag, or Delete
! - Right-click POPUP on list allows Tag, Copy symbol name, Delete, Expand all, Collapse all
! - Ctrl+C on Copies symbol name
! - Queue sorting is no longer case sensitive
! - Check for 64-bit DLL and warn. see IMAGE_NT_OPTIONAL_HDR64_MAGIC
! - Check for 16-bit DLL and warn.
! - Common Symbols to skip spotted: DllGetVersion, DllRegisterServer, DllUnregisterServer, DllInstall, DllCanUnloadNow, DllGetActivationFactory, DllGetClassObject
! - Save LIB now saves Filtered if box is checked Write Filtered and Found. WTF was the filtered tab for?
! - Command line READ= WRITE= /CLOSE copied from new MG
! - Can Send TO or Drop on EXE multiple file names
! - Drop onto LIST supports multiple files 
! - Add Files can pick Multiple in FileDialog 
! - Xfer Changes Mark did in August 2019 - One Generate Code Button
! - Refactor Write Filtered with WriteExpQ that has Tree 1 Libs. Ideally Tree should be in FilteredQ but too much code to check or refactor
! - Make generate Classes do only Filtered. 
! - Refactor soring so it is right for write
! - Got rid of List *Colors, changed to Y Styles, copied MG August design mostly
! - "Rx" button on Search window with Clarion MATCH Help
! - WinSxS Folder File Picker to get version 6 of ComCtl32.DLL (in System32 is v5.8 without new stuff). 
!       Right click on Add button and Popup has WinSxS, so does Subtract.
!       E.g. Add v6 of ComCtl from SxS, then subtract v5 from Sys32 (pick Tag not Delete) to see what's new

!Updated 2014-Jan-14 By Mark Goldberg 
!  - Added Ability to Generate Classes (no parameters )
!  - Refactored some of the code.
!  - A Lot of code is still badly written
!                                  
!EndRegion
!Region Ancient Documentation
! =========================================================================================
! Original Source code copied from shipping example of Libmaker that came with C5EEB3
! Updated: MG == by Mark Goldberg of Monolith Custom Computing, Inc.
!
! NOTE: MGLibMkr was built using C5EE, as a result the file drivers will not be recognized
!         by other versions of CW (ex: C4),
!         Work around:
!           a) Delete the 2 lines in Library, object and resource files that have the % symbols showing
!           b) Under Database driver libraries, add the ASCII and DOS drivers
!
!
! Updates:
!   MG 10/6/98
!       Changed user interface:
!          replaced buttons with toolbar
!          using MGResizeClass to resize the listbox
!          using 32-bit filedialogs
!          added EXE to file formats on the AddFile FileDialog
!          Fixed a minor bug in EnableControls
!          added VCR controls to listbox
!          Made the 1st column resizable (so can see lengthy prototypes)
!          Save & Restore 1st column width
!          Save & Restore Window Position & Size
!       Added:
!          Make CW Map feature
!          add command line support for file name to load
!          if do not give a file at the command line, then will auto-accept the fileAdd
!
!       Doesn`t Work:
!          Added DropID('~FILE') but it doesn`t seem to accept dragging from Explorer or WinFile !!
!            (failed under C5EEB3 & C4a)
!
!   MG 10/18/98 (1:20 - 4:15 )
!        Moved DropID('~FILE') from List control to Window, now it works!
!        Added several features of Arnor Baldvinssons, see lines marked with !AB
!        Changed Search system of Arnor`s to use the ExportQ and Colorize and set Icon to indicate Found Lines
!          wrote:
!           SetThisColor     (Long argNFG, Long argNBG,*FormatColorSet argFCS)
!           SetGloColors
!        Added Myself to Arnor`s help about window
!           (see Help about window documentation for additional changes)
!        Moved much of the "technical" file definitions into FileDefs.INC
!        Added::  ?List1{PROP:Format} = ?List1{PROP:Format} !<-- Added to force re-draw of Q (needed in C5EEB4 and likely elsewhere)
!        Changed the Search Window dramatically, renamed the search names too
!
!
!   MG 10/20/98 (7:15 - 10:00 )
!        Change PropList:Width,1 when Window{prop:Width} changes, so the function column resizes vs. the ordinal column
!          Removed save/restore column width
!        Added an icon for search (copied from a sample program posted by Andy Stapleton) -- is this OK?
!        Refined setting Icon for TreeLevel = +/- 1
!        Refined Contains search method, so that a single letter search will hit when found in the first character of the symbol
!        Changed Sort Buttons to a Radio Option, using a blank icon to accomplish a latched look
!        Added hot-letters to several buttons in the toolbar (see their Tip Help for indication)
!        Added support to launch a notepad with the just created CW MAP
!
!        Doesn`t Work:
!           Locator button on VCR control
!           CW MAP: lcl.symbol = lcl:symbol[ 1 : lcl.instringLoc - 1 ]  !BUG: for some DLL`s, I found one compiled with CBuilder 1.0, that this would be bad for
!
!   MG 11/16/98
!   MG 01/19/99 Added ",DLL(1),RAW" to map export
!   MG 12/07/01 Recompiled using C55eeF
!
!   MG  1/22/03 Quick and dirty addtion of the Ron Schofields .API file format --- UNFINISHED
!
! Future Features:
!   b) Keep a list of recent .DLL`s that have been opened
!   c) Add an option to auto-create a CW map with the same base name as the .LIB being saved
!   d) When delete a function, then mark it with colorization and an icon
!       i) add support for show/hide deleted functions
!      ii) pay attention to deleted functions when writing a .LIB
!     iii) add support for undeleting functions
!      iv) consider allowing deletion based on the results of the current search
!   e) Support for multiple files in a drag & drop situation
!   f) Support for multiple files in a AddFile
!   g) Enhance command line features
!   h) Support Recent file lists
!   i) have reasonable defaults for filenames when GenerateMap & WriteLib

! TODO (notes: June 20 2005)
!   replace color work with STYLES
!   remove lcl/rou GROUPS, use LCL: and ROU:
! =========================================================================================
!EndRegion Ancient Documentation

     MAP
        ReadExecutable
        DumpPEExportTable  (ULONG RawAddr, ULONG VirtAddr)
        DumpNEExports
        WriteLib
        ReadLib
        InfoWindow
        GenerateMap(byte argRonsFormat)
        GenerateClasses()
        SelectQBE
        SortExportQ(BYTE ForcedOrder=0)                 !CB Put all sorting here
        SelectExportQ(LONG ListFEQ, ExportQ ExportQType, LONG QPointer)  !Expand Tree before select symbol        
        ExportQ_to_FilterQ()
        ExportQCompress   (ExportQ ExportQType)         !CB Remove Empty Level 1
        LongestSymbol     (ExportQ ExportQType),LONG    !in ExportQ Level 2
        CleanSymbol       (STRING xSymbol),STRING
        AppendAscii       (STRING xLine)
        SubtractFileLIB   (<*string InFileName>)        !CB added subtract Feature mainly for Win32.LIB       
        CopyQueue         (*queue q1,*queue q2)
        SkipAddExportQ(),BOOL
        SetSearchFlagInExportQ(BYTE NewSearchFlag, BYTE SyncFoundCount, BYTE PutExportQ)  !CB
        WriteTextEXP()
        CopyList2Clip()
        SetCueBanner(LONG TextSingleFEQ, STRING CueBannerText, BOOL OnFocusShows)
        FileNamePopupClarion(*STRING OutFileName),BOOL
        FileNamePopupSystem32(*STRING OutFileName),BOOL
        PopupUnder(LONG UnderCtrlFEQ, STRING PopMenu),LONG
        ListInit(SIGNED ListFEQ)
        RegExHelper(STRING AtX, STRING AtY, STRING AtW)
        ExistsFile(STRING FileNm),BOOL     !Is it a File and not Folder?
        ExistsFolder(STRING FolderNm),BOOL !Is it a Folder and not File?        
        WinSxSPicker(*STRING OutFileName),BOOL
        DB(STRING DebugMsg)
        MODULE('Win32')
            GetWindowsDirectory(*CString lpBuffer, LONG uSize),UNSIGNED,PASCAL,RAW,PROC,DLL(1),NAME('GetWindowsDirectoryA')
            GetSystemDirectory(*CSTRING lpBuffer, LONG uSize),LONG,PASCAL,RAW,DLL(1),NAME('GetSystemDirectoryA')
            SendMessageW(SIGNED hWnd, UNSIGNED Msg, UNSIGNED wParam, SIGNED lParam ),PASCAL,DLL(1),RAW,SIGNED,PROC!,NAME('SendMessageW')
            GetFileAttributes(*CString szFileName),LONG,PASCAL,RAW,DLL(1),NAME('GetFileAttributesA')
            OutputDebugString(*cstring Msg),PASCAL,RAW,DLL(1),NAME('OutputDebugStringA')
        END
     END

qINI_FileName  Equate('MGLibMkr.ini') !CB with Manifest PUT not working to VirtualStore
qINI_File      STRING(260) !CB EXE Path \ MGLibMkr.INI. Move to Registry?
qOrdColWidth   Equate(41) 

     include(   'KeyCodes.clw'),ONCE
     include(    'Equates.clw'),ONCE
     include(   'FileDefs.inc'),ONCE
     include('ResizeClass.inc'),ONCE
SubFileName   LIKE(FileName),STATIC     
MGResizeClass ResizeClassType

ListStyle:None      EQUATE(0) 
ListStyle:Opened    EQUATE(1) 
ListStyle:Closed    EQUATE(2) 
ListStyle:Found     EQUATE(3)

Glo     Group
SortOrder   String('Original')
DisplayCount Long
FoundCount   Long !MG Added 12/7/01
IsSubtracting byte
WriteFiltered byte  !CB - WTF is Filtered for if it is not to Save?
        END !glo
AddExportQGrp   GROUP,PRE(AddEx)    !CB 09/06/20 moved out of SkipAddExportQ so can be cleared
OmtIncAllDllX BYTE                  !Symbols that should be excluded  1=Omit All  2=Include All  e.g. do NOT import DllGetVersion
TagIncAllDllX BYTE                           !1=Tag them (when include all)
OmtIncAllDups BYTE                  !Symbols that Duplicate           1=Omit All  2=Include All
TagIncAllDups BYTE                           !1=Tag them (when include all)
                END                 
WindowsDir CSTRING(256)
System32  CSTRING(256)        
NDX       LONG
OrgOrder  LONG          !CB Index separate from Q to assure cannot get messed up
ExportQ   QUEUE,PRE(EXQ)
symbol      STRING(128)  !1
Icon        LONG         !2   3=Tagged Icon  0=Untagged No Icon  1=Lib Opened 2=Lib Closed 
TreeLevel   SHORT        !3   1=Lib/Dll  2=Symbol
StyleNo     LONG         !4   See ListStyle:*  3=:Found Tagged, 0=:None UnTag, 1=:Opened 2=:Closed
Ordinal     LONG         !5
Module      STRING(260)  !6   !Modified by MJS to allow for longer file names
OrgOrder    LONG         !7   !AB Unique ID of Import
SearchFlag  Byte         !8   !MG See SearchFlag::* below  1=::Found (Tagged)   0=::Default (Not Tagged)
SymbolLwr   STRING(128)  !9   !CB Lower Case for Sort and Find  (but Windows Symbol names are case sensitive)
          END
FilteredQ QUEUE(ExportQ),PRE(FLTQ)
          END
WriteExpQ QUEUE(ExportQ),PRE(WrtQ)  !Either ExportQ or FilteredQ depending on Glo.WriteFiltered
          END
ASCIIfile FILE,DRIVER('ASCII'),PRE(ASCII),CREATE,NAME(FileName)
            Record
Line          String(2 * Size(ExportQ.symbol) + 30)
            END
          END
TxtFileName STRING(260),STATIC
TxtFile  FILE,PRE(Txt),DRIVER('ASCII','/FILEBUFFERS=20'),CREATE,NAME(TxtFileName)
Record     RECORD
Line         STRING(256)
           END
         END
          
!Region - Main Window Code Region
SearchFlag::Default       Equate(0)
SearchFlag::Found         Equate(1)

LOC:FoundRec  LONG        !AB   (mg, shows found count on the window -- todo)
LocatorText   STRING(64)
window WINDOW('MGLibMkr Plus by Carl Barnes'),AT(,,288,257),CENTER,GRAY,IMM,SYSTEM,ICON('LIBRARY.ICO'),FONT('Segoe UI',9),ALRT(CtrlShiftD), |
            DROPID('~FILE'),RESIZE
        TOOLBAR,AT(0,0,288,42),USE(?TOOLBAR1),COLOR(COLOR:ACTIVECAPTION)
            BUTTON('&O'),AT(2,2,14,14),USE(?AddFile),ICON(ICON:New),ALRT(MouseRight),TIP('Open / Add File... (ALT+O)<13>' & |
                    '<10>Right-click for Popup of common choices'),FLAT,LEFT
            BUTTON('&M'),AT(19,2,14,14),USE(?SubtractFile),DISABLE,ICON(ICON:Cut),ALRT(MouseRight),TIP('Subtract .LIB or' & |
                    ' .DLL ... (ALT+M = Minus)<13><10>Normally Subtract Clarion Win32.LIB<13,10>Right-click for Popup of' & |
                    ' common choices'),FLAT,LEFT
            BUTTON('&T'),AT(53,2,14,14),USE(?TagPopup),DISABLE,ICON('found.ico'),TIP('Tag Options: Tag All, Untag All, D' & |
                    'elete Tagged... (ALT+T)'),FLAT,LEFT
            BUTTON('&S'),AT(70,2,14,14),USE(?SaveAs),DISABLE,ICON(ICON:Save),TIP('Save .LIB As... (ALT+S)'),FLAT,LEFT
            BUTTON,AT(87,2,14,14),USE(?CopyListBtn),DISABLE,ICON(ICON:Copy),TIP('Copy Symbol names from All or Filtered List'),FLAT,LEFT,SKIP                    
            BUTTON,AT(105,2,15,14),USE(?Clear),DISABLE,ICON('clear.ico'),TIP('Empty the listbox'),FLAT,LEFT
            BUTTON,AT(122,2,14,14),USE(?Exit),STD(STD:Close),ICON('Exit.ico'),TIP('Exit the program'),FLAT,LEFT
            BUTTON,AT(245,2,14,14),USE(?ReRun),ICON(ICON:VcrPlay),TIP('Run Another Instance of LibMaker'),FLAT,LEFT,SKIP
            BUTTON,AT(267,2,14,14),USE(?Info),ICON(ICON:Help),TIP('Info'),FLAT,LEFT,SKIP
            CHECK('Write Filtered'),AT(145,2),USE(glo.WriteFiltered),TIP('Saving LIBs, EXPs, MAPs, Classes, etc<13,10>us' & |
                    'e the Filtered list if any are tagged.')
            BUTTON('Write E&XP'),AT(145,15,58,10),USE(?MakeEXP),DISABLE,TIP('Write EXP file for Exporting Symbols from L' & |
                    'ibrary.<13,10>If you use EXPORT attrbiute this will make you the EXP file.'),FLAT,LEFT
            BUTTON('Generate Code'),AT(145,28,58,10),USE(?GenerateCode),DISABLE,TIP('Generate Code for MAP or CW Classes...'), |
                    FLAT,LEFT
            OPTION('Sort Order'),AT(4,17,130,23),USE(glo.SortOrder),BOXED
                RADIO('Original'),AT(8,25,37,12),USE(?glo:SortOrder:Radio1)
                RADIO('Name'),AT(52,25,31,12),USE(?glo:SortOrder:Radio2)
                RADIO('Ordinal'),AT(89,25,34,12),USE(?glo:SortOrder:Radio3)
            END
            PROMPT('Found'),AT(261,19,,8),USE(?GLO:FoundCount:Prompt),TRN,FONT(,8,COLOR:Blue,FONT:regular,CHARSET:ANSI)
            STRING(@n6),AT(227,19,29,8),USE(glo.FoundCount),TRN,RIGHT,FONT(,,COLOR:Blue,,CHARSET:ANSI)
            STRING('Total'),AT(261,29,,8),USE(?GLO:DisplayCount:Prompt),TRN,FONT(,8,,FONT:regular,CHARSET:ANSI)
            STRING(@n6),AT(227,29,29,8),USE(Glo:DisplayCount),TRN,RIGHT
            BUTTON('&F'),AT(36,4,14,12),USE(?FindButton),DISABLE,KEY(CtrlF),ICON('Cs_srch.ico'),TIP('Search for string a' & |
                    'nd mark symbols (ALT+F)<13,10>You may also toggle marks with Enter, Insert, or Mouse'),FLAT,LEFT
        END
        GROUP,AT(90,2,196,15),USE(?LocateGrp),HIDE
            TEXT,AT(94,5,139,10),USE(LocatorText),SKIP,FONT('Consolas',9,,FONT:regular),ALRT(EnterKey),SINGLE
            BUTTON('&Next'),AT(237,4,23,11),USE(?LocateNext),SKIP
            BUTTON('&Prev'),AT(262,4,23,11),USE(?LocatePrev),SKIP
        END
        SHEET,AT(4,4,283,210),USE(?SHEET)
            TAB(' All '),USE(?TAB:All)
                LIST,AT(8,23,272,187),USE(?List1),DISABLE,VSCROLL,COLUMN,VCR,FROM(ExportQ),FORMAT('143L(2)|MIYT~Module a' & |
                        'nd function~30R(3)~Ordinal~@N_5B@'),ALRT(MouseLeft2), ALRT(EnterKey), ALRT(DeleteKey), |
                         ALRT(InsertKey), ALRT(CtrlC)
            END
            TAB(' Filtered '),USE(?TAB:Filtered),ICON('found.ico')
                LIST,AT(8,23,272,187),USE(?Filtered),DISABLE,VSCROLL,COLUMN,VCR,FROM(FilteredQ), |
                        FORMAT('160L(2)|MIYT~Symbol~@s255@70L(2)|M~Module~@s255@#6#30R(3)~Ordinal~@N' & |
                        '_5B@#2#'),ALRT(MouseLeft2), ALRT(DeleteKey), ALRT(CtrlC), ALRT(EnterKey)
            END
            TAB('D'),USE(?TAB:DebugWrite),HIDE
                LIST,AT(8,23),USE(?ListWrite),FULL,VSCROLL,TIP('WriteExpQ is Filtered with Tree so can use in place o' & |
                        'f ExportQ'),VCR,FROM(WriteExpQ),FORMAT('143L(2)|MIYT~Module and function~30R(3)~Ordinal~@N_5B@')
            END
        END
    END
   CODE
   DO PreAccept
   DO AcceptLoop
   DO PostAccept
!------------------------------------------------------
PostAccept           ROUTINE
   MGResizeClass.Close_Class
   if ~0{prop:Iconize}
      PutIni('Position','X',0{prop:XPOS } ,qINI_FILE)
      PutIni('Position','Y',0{prop:YPOS } ,qINI_FILE)
      PutIni('Position','W',0{prop:Width} ,qINI_FILE)
      PutIni('Position','H',0{prop:Height},qINI_FILE)
   !  PutIni('Position','Col1W',?List1{proplist:Width,1},qINI_FILE)
   end
   PutIni('Settings','WriteFiltered',glo.WriteFiltered,qINI_FILE)
!------------------------------------------------------
PreAccept            ROUTINE
   qINI_File=COMMAND('0') !CB Path\Name.EXE for where toput INI file. With Manifest cannot PUT to VirtStore
   Ndx=INSTRING('\',qINI_File,-1,SIZE(qINI_File))
   qINI_File=SUB(qINI_File,1,Ndx) & qINI_FileName
   SYSTEM{PROP:PropVScroll}=1
   SYSTEM{7A7Dh}=MSGMODE:CANCOPY  !PROP:MsgModeDefault added sometime in C11
   IF ~GetWindowsDirectory(WindowsDir,SIZE(WindowsDir)) then WindowsDir='c:\wINDOWS'.
   IF ~GetSystemDirectory(System32,size(System32)) then System32='c:\wINDOWS\system32'.
   OPEN(window)
   System{prop:Icon} = window{prop:icon} !Have any non-minimize-able additional windows share the same Icon
   ?SHEET{PROP:NoTheme}=1   !CB Visual Styles
   ?TOOLBAR1{PROP:Color}=8000001Bh !COLOR:GradientActiveCaption added C11 but works since C8
   SetCueBanner(?LocatorText,'Symbol Name to Find ...',0)   
   ExportQ.orgorder = 0     !AB

   ! ?list1{prop:vcr}=TRUE   !Suppress the ? button in the vcr
   window {PROP:minheight}  = 76
   window {PROP:minwidth }  = window{PROP:width}
   MGResizeClass.Init_Class(window)
   MGResizeClass.Add_ResizeQ(?List1          ,'L/T R/B')
   MGResizeClass.Add_ResizeQ(?Filtered       ,'L/T R/B')
   MGResizeClass.Add_ResizeQ(?SHEET          ,'L/T R/B')
   MGResizeClass.Add_ResizeQ(?TAB:All        ,'L/T R/B')
   MGResizeClass.Add_ResizeQ(?TAB:Filtered   ,'L/T R/B')
	
   !---- apparently there is a problem with my resize, when I pass an FEQ from the toolbar
   !---- this has the effect of breaking the resize, although the INITIAL resize works, which leads me to believe that
   !---- the problem is related to my use of 0{prop:  } code
!      MGResizeClass.Add_ResizeQ(?GLO:DisplayCount:Prompt,'R/T R/T')
!      MGResizeClass.Add_ResizeQ(?GLO:DisplayCount       ,'R/T R/T')
   !---- apparently there is a problem with my resize, when I pass an FEQ from the toolbar -- end
  !MGResizeClass.Add_ResizeQ(?SearchList,'L/T R/B') !Problem!

!INI moved to EXE folder so could be on diff systems with diff rez and monitors
!   0{prop:XPOS}   = GetIni('Position','X',0{prop:XPOS } ,qINI_FILE)
!   0{prop:YPOS}   = GetIni('Position','Y',0{prop:YPOS } ,qINI_FILE)
!   0{prop:Width}  = GetIni('Position','W',0{prop:Width} ,qINI_FILE)
!   0{prop:Height} = GetIni('Position','H',0{prop:Height},qINI_FILE)
   glo.WriteFiltered = GetIni('Settings','WriteFiltered',1,qINI_FILE)

   !?List1{proplist:Width,1} = GetIni('Position','Col1W',?List1{proplist:Width,1},qINI_FILE)
   MGResizeClass.Perform_Resize()
   ListInit(?List1)    !Set IconList, Styles, Column 1 Width
   ListInit(?Filtered)
   ListInit(?ListWrite)

!   ReadFileName  = GetIni('RecentFiles','Read' ,ReadFileName ,qINI_File)
!   WriteFileName = GetIni('RecentFiles','Write',WriteFileName,qINI_File)
!   AsciiFileName = GetIni('RecentFiles','Ascii',AsciiFileName,qINI_File)
   
   DO PreAccept:CommandLine

!------------------------------------------------------
PreAccept:CommandLine ROUTINE    

   !Message('Command()|['& Command() &']||Path()|['& PATH() &']')
   IF COMMAND('READ') THEN
      FileName = CLIP(LEFT(COMMAND('READ')))
      IF ~EXISTS(FileName) THEN
          Message('Command Line READ file does not exist||' & COMMAND())
          EXIT 
      END
      DO FileAdded
   ELSE  !File name dropped on EXE ? 
      FileName = CLIP(LEFT(Command())) 
      IF MATCH(FileName,' .:\\',Match:Regular) THEN !Has <32>X:\
         DO CmdLine_Multi_Or_SendTo ; EXIT
      END
      IF EXISTS(FileName) THEN DO FileAdded ; EXIT.
      CLEAR(FileName)
      SETKEYCODE(MouseRight)
      POST(Event:AlertKey,?AddFile)
     ! POST(Event:Accepted,?AddFile)
   END

   IF COMMAND('WRITE') AND RECORDS(ExportQ)
      FileName = COMMAND('WRITE')
                                          ! DBG.Debugout('Write Filename['& CLIP(FileName) &']')
	  SETCURSOR(CURSOR:Wait)                                                     
	  WriteLib()
	  SETCURSOR()                                       
      IF ~COMMAND('/CLOSE') ! assume user has automated system running and doesn't want UI
         MESSAGE('[' & CLIP(FileName) &'] written','MG Library Maker', SYSTEM{PROP:Icon})
      END 
   END 
                                          ! DBG.Debugout('COMMAND(''/CLOSE'')['& COMMAND('/CLOSE') &']')
   IF COMMAND('/CLOSE')        !??? Why READ and CLOSE?
      POST(EVENT:CloseWindow)
   END
   DISPLAY()

CmdLine_Multi_Or_SendTo ROUTINE !Dropped file names on EXE, same as SEND TO
    DATA
Filez   &STRING
P1      LONG    
L1      LONG    
    CODE
!??? \\UNCServer\Share ????            Space X:\    cannot have Colon in file name 
!Format:[ D:\Clarion91\Lib\debug\ClaSQAl.lib D:\Clarion91\Lib\debug\ClaSCAl.lib D:\Clarion91\Lib\debug\ClaSQA.lib]    
  !Message('Cmd Line has multiple!||[' & Command() & ']','debug')  
  L1 = LEN(CLIP(Command())) + 1
  FileZ &= NEW(STRING(L1))
  FileZ = LEFT(Command())        !Rmv leading space
  LOOP 100 TIMES !Prevent hang
      P1=STRPOS(FileZ,' .:\\')   !Find: <32>C:\
      IF ~P1 THEN P1=L1.
      FileName=LEFT(SUB(FileZ,1,P1))
      FileZ=LEFT(SUB(FileZ,P1,L1))
      IF EXISTS(FileName) THEN DO FileAdded.
  WHILE FileZ
  DISPOSE(FileZ)   
   
LoadWin32LibDllsRtn ROUTINE 
    IF MESSAGE('This will load the 13 Windows System32 DLLs referenced in Clarion Win32.LIB.'&|
         '|(Kernel User GDI Comctl ComDlg Winnm MPR AdvApi Ole OleAut OleDlg Shell)'&|
         '||You can use this to subtract the shipping Win32.Lib to see what''s missing.'&|
         '|The Win32.lib shipped is typically limited to Windows XP and prior.'&|
         '||This load will contain a lot of duplicates because Windows has added many '&|
         '|Export forwards. You should omit them, or tag them and '&|
         'review then delete.', |
         'Create Win32.Lib from Windows DLLs', ICON:Question, '&Load DLLs|&Cancel', 2) = 2 THEN EXIT. 
    DO Accepted:Clear 
    CLEAR(AddExportQGrp)   !Ask about Dups 
    AddEx:OmtIncAllDllX=1  !1=Omit All Problems e.g. DllGetVersion

    FileName='Kernel32.DLL' ; DO LoadOneWin32LibRtn
    FileName='User32.DLL'   ; DO LoadOneWin32LibRtn
    FileName='GDI32.DLL'    ; DO LoadOneWin32LibRtn
    FileName='COMCTL32.DLL' ; DO LoadOneWin32LibRtn
    FileName='COMDLG32.DLL' ; DO LoadOneWin32LibRtn
    FileName='WINSPOOL.DRV' ; DO LoadOneWin32LibRtn
    FileName='WINMM.DLL'    ; DO LoadOneWin32LibRtn
    FileName='MPR.DLL'      ; DO LoadOneWin32LibRtn
    FileName='ADVAPI32.DLL' ; DO LoadOneWin32LibRtn
    FileName='OLE32.DLL'    ; DO LoadOneWin32LibRtn
    FileName='OLEAUT32.DLL' ; DO LoadOneWin32LibRtn
    FileName='OLEDLG.DLL'   ; DO LoadOneWin32LibRtn
    FileName='Shell32.DLL'  ; DO LoadOneWin32LibRtn
!Now in GDI32    FileName=System32 &'\LPK.DLL' ; ReadExecutable()       
    DO CollapseAll
    0{PROP:Text}='Loaded DLLs to Create Win32.LIB'
LoadOneWin32LibRtn ROUTINE
    FileName=System32 &'\' & FileName
    0{PROP:Text}='Load ' & FileName 
    ReadExecutable()   
CollapseAll ROUTINE
    LOOP Ndx=1 TO RECORDS(ExportQ)  
        GET(ExportQ,Ndx) ; IF ExportQ.TreeLevel<>1 THEN CYCLE.
        ExportQ.TreeLevel=-1 
        ExportQ.Icon=2
        ExportQ.StyleNo=ListStyle:Closed
        PUT(ExportQ)
    END
!------------------------------------------------------
!Region Accept Loop
AcceptLoop           ROUTINE   
   ACCEPT	 
     CASE ACCEPTED()
       OF ?glo:SortOrder   ; DO Accepted:SortOrder
       OF ?FindButton      ; DO Accepted:FindButton
       OF ?MakeEXP         ; DO Accepted:MakeEXP
       OF ?GenerateCode    ; DO Accepted:GenerateCode
       OF ?ReRun           ; RUN(Command('0'))
       OF ?Info            ; InfoWindow()
       OF ?CopyListBtn     ; CopyList2Clip()
       OF ?Clear           ; DO Accepted:Clear
       OF ?AddFile         ; DO Accepted:AddFile
       OF ?SubtractFile    ; DO Accepted:SubtractFile
       OF ?SaveAs          ; DO Accepted:SaveAs 
       OF ?TagPopup        ; DO Accepted:TagPopup 
       OF ?LocateNext
       OROF ?LocatePrev    ; DO Accepted:LocateNext
     END 

     CASE EVENT()
       OF EVENT:expanded
     OROF EVENT:contracted ; DO OnExpandContract
       OF EVENT:AlertKey   ; DO OnAlertKey
       OF EVENT:NewSelection ; DO OnNewSelection
       OF Event:Drop       ; DO OnDrop
       OF EVENT:Locate     ; DO OnLocate				
       OF EVENT:TabChanging; ExportQ_to_FilterQ() !could be a bit more subtle, to only catch when switching TO the FilterQ			
     END      
     
     IF MGResizeClass.Perform_Resize()
        ?List1   {proplist:Width,1} = ?List1   {prop:Width} - qOrdColWidth
		?Filtered{proplist:Width,1} = ?Filtered{prop:Width} - qOrdColWidth - ?Filtered{proplist:Width,2}  
     END
     
     DO EnableDisable
     DO Set_DisplayCount
   END 

!------------------------------------------------------
Set_DisplayCount     ROUTINE
  Glo:DisplayCount = Records(ExportQ)
  !todo: refine this value as desired,
  !  show a count of found records when a search is in effect
  !  show only level 2 records
!------------------------------------------------------
EnableDisable       ROUTINE
    DATA
D1 STRING(1)  !09/06/20 CB replace =?List1{PROP:Disable} repeated below
    CODE
  D1=CHOOSE(RECORDS(ExportQ)=0,'1','')
  ?List1        {PROP:Disable}=D1
  ?Filtered     {PROP:Disable}=D1
  ?SubtractFile {PROP:Disable}=D1
  ?SaveAs       {PROP:Disable}=D1
  ?Clear        {PROP:Disable}=D1
  ?CopyListBtn  {PROP:Disable}=D1
  ?MakeEXP      {PROP:Disable}=D1
  ?GenerateCode {PROP:Disable}=D1
  ?Glo:SortOrder{PROP:Disable}=D1
  ?FindButton   {PROP:Disable}=D1 
  ?TagPopup     {PROP:Disable}=D1 
  ?LocateGrp    {PROP:Hide}   =D1
  EXIT
!Region Accepted ROUTINEs
Accepted:SortOrder ROUTINE
   GET(ExportQ,CHOICE(?List1))    !CB preserve selected   
   SetCursor(CURSOR:Wait)
   SortExportQ()
!   Execute Choice(?glo:SortOrder)
!      BEGIN; SORT(ExportQ,ExportQ.OrgOrder)                                    ; SORT(FilteredQ,FLTQ:orgorder)                           ; END
!      BEGIN; SORT(ExportQ,ExportQ.Module,ExportQ.TreeLevel,ExportQ.symbolLwr)  ; SORT(FilteredQ,FLTQ:Module,FLTQ:TreeLevel,FLTQ:symbolLwr)  ; END
!      BEGIN; SORT(ExportQ,ExportQ.Module,ExportQ.TreeLevel,ExportQ.Ordinal)    ; SORT(FilteredQ,FLTQ:Module,FLTQ:TreeLevel,FLTQ:Ordinal) ; END
!   END
   ?List1   {PROP:Format} = ?List1   {PROP:Format} !<-- Added to force re-draw of Q (needed in C5EEB4 and likely elsewhere)
   ?Filtered{PROP:Format} = ?Filtered{PROP:Format} !<-- Added to force re-draw of Q (needed in C5EEB4 and likely elsewhere)
   GET(ExportQ,ExportQ.OrgOrder)           !CB preserve selected   
   ?List1{PROP:Selected}=POINTER(ExportQ)  !CB preserve selected   
   SetCursor()
   
Accepted:FindButton ROUTINE
   SelectQBE()
   ?List1   {PROP:Format} = ?List1   {PROP:Format} !<-- Added to force re-draw of Q (needed in C5EEB4 and likely elsewhere)
   ?Filtered{PROP:Format} = ?Filtered{PROP:Format} !<-- Added to force re-draw of Q (needed in C5EEB4 and likely elsewhere)
   IF ?Sheet{PROP:ChoiceFEQ}=?TAB:Filtered THEN ExportQ_to_FilterQ().
   DO Set_DisplayCount

Accepted:GenerateCode  ROUTINE
    IF ~RECORDS(ExportQ) THEN EXIT.
    SortExportQ(1)
    IF Glo.WriteFiltered THEN 
       ExportQ_to_FilterQ()   !CB call also makes WriteExpQ
       IF Glo.FoundCount THEN SELECT(?TAB:Filtered).   !CB show Filtered as Reminder
    END
    IF ~Glo.WriteFiltered OR ~Glo.FoundCount THEN
        FREE(WriteExpQ)  
        CopyQueue(ExportQ,WriteExpQ)
    END
    DISPLAY
    EXECUTE MESSAGE('What Code would you like to generate?','Generate Code', SYSTEM{PROP:Icon},|
                 '&Map|Map (&Rons)|&Classes|Cl&ose')
       GenerateMap( FALSE )
       GenerateMap( TRUE  )
       GenerateClasses()
    END
    SortExportQ()
    EXIT
  
Accepted:Clear ROUTINE       
   FREE(ExportQ); FREE(FilteredQ)
   OrgOrder = 0 !CB
   CLEAR(AddExportQGrp)   !CB Ask about Dups 
   Window{PROP:Text} = 'LibMaker'
   DISPLAY

Accepted:AddFile  ROUTINE
    DATA
Filez   STRING(5000) !CB Muti
PthEnd  LONG    
P1      LONG    
P2      LONG    
L1      LONG    
   CODE
   Filez = FileName
   IF ~FileDialog('Import Symbols from Files(s) ...', Filez, |
                 'DLLs and LIBs|*.dll;*.lib|DLL files (*.dll)|*.dll|LIB files (*.lib)|*.lib' & |
                 '|Drivers (*.drv)|*.drv|OCX files (ocx.*)|*.ocx' & |
                 '|Executables (*.exe)|All files (*.*)|*.*', |
                 FILE:LongName + FILE:Multi) THEN EXIT.

    P1=INSTRING('|',Filez)
    IF ~P1 THEN
        FileName = Filez
        DO FileAdded ; DISPLAY
        EXIT
    END
    !Multi Format: C:\Windows\System32|kernel32.dll|mfasfsrcsnk.dll|mfc40u.dll
    Filez[P1]='\' ; PthEnd = P1 ; P1 += 1
    L1=LEN(CLIP(Filez))
    LOOP P2=P1 TO L1
         IF P2 < L1 AND Filez[P2+1]<>'|' THEN CYCLE.
         FileName = Filez[1 : PthEnd] & FileZ[P1 : P2]  
         DO FileAdded ; DISPLAY
         P1 = P2 + 2 ; P2 = P1
    END 
    EXIT
   
Accepted:SubtractFile  ROUTINE
    SubtractFileLIB() ; DO EnableDisable ; SELECT(?LIST1)

Accepted:SaveAs   ROUTINE       
   IF RECORDS(ExportQ)>0
     Clear(FileName) !FileName = WriteFileName
     IF FileDialog('Save OMF library definition as ...', FileName, 'Library files (*.lib)|*.lib',FILE:Save+FILE:LongName)
        Window{PROP:Text} = 'LibMaker - ' & CLIP(FileName) !AB
        SetCursor(CURSOR:Wait)                             !AB
        !CB SORT(ExportQ,ExportQ.orgorder)                     !AB
        !CB SORT(FilteredQ,FLTQ:orgorder)
        DO SortByOrgOrderRtn  !CB for WriteFiltered and more
        WriteLib()
        SetCursor()                                        !AB
        !WriteFileName = FileName
     END
   END

SortByOrgOrderRtn ROUTINE
   IF Glo.WriteFiltered THEN 
      ExportQ_to_FilterQ()                            !CB assure filled 
      IF Glo.FoundCount THEN SELECT(?TAB:Filtered).   !CB show Filtered as Reminder
   END
   GET(ExportQ,CHOICE(?List1))         !CB Preserve Select 
   glo.SortOrder=1                     !CB Sort Radio to Orig   
   SortExportQ()
!   SORT(ExportQ,ExportQ.orgorder)                     !AB
   GET(ExportQ ,ExportQ.orgorder)      !CB Preserve Select
!   SORT(FilteredQ,FLTQ:orgorder)
   DISPLAY

Accepted:MakeEXP   ROUTINE       
   IF RECORDS(ExportQ)=0 THEN EXIT.
   DO SortByOrgOrderRtn
   WriteTextEXP()

Accepted:TagPopup ROUTINE
    DATA
TagOp   BYTE,AUTO
QNdx    LONG,AUTO
NoTags  PSTRING(3)   
    CODE
    NoTags=CHOOSE(~Glo:FoundCount,'|~','|')
    TagOp=POPUPunder(?TagPopup,|
                         'Tag All' & |         !#1            
                NoTags & 'Un Tag All' & |      !#2
                NoTags & 'Toggle All' & |      !#3
                '|-' & |
                NoTags & 'Delete Tagged' & |   !#4
                NoTags & 'Delete UnTagged' & | !#5
                '|-' & |
                '|Tag Duplicate Names')        !#6 
    CASE TagOp 
    OF 0 ; EXIT
    OF 6
       DO Accepted:TagDuplcates
       DISPLAY
       EXIT
    END 
    SELECT(?TAB:All)
    LOOP QNdx=RECORDS(ExportQ) TO 1 BY -1
         GET(ExportQ,QNdx) ; IF ExportQ.TreeLevel <> 2 THEN CYCLE.
         CASE TagOp
          OF 1  ; SetSearchFlagInExportQ(SearchFlag::Found,1,1)
          OF 2  ; SetSearchFlagInExportQ(SearchFlag::Default,1,1) 
          OF 3  ; SetSearchFlagInExportQ(CHOOSE(ExportQ.SearchFlag=SearchFlag::Found,SearchFlag::Default,SearchFlag::Found),1,1) 
          OF 4  ; IF ExportQ.SearchFlag = SearchFlag::Found THEN 
                     Glo:FoundCount -= 1
                     DELETE(ExportQ)
                  END   
          OF 5  ; IF ExportQ.SearchFlag = SearchFlag::Default THEN 
                     DELETE(ExportQ)
                  END
         END    
    END   
    DISPLAY
    EXIT
Accepted:TagDuplcates ROUTINE !09/2/20 Noticed in Create Win32.lib had 59 dups
    DATA
DupQ   QUEUE(ExportQ),PRE(DupQ)
       END    
CntDp  LONG
QnX    LONG
TagAll BOOL 
    CODE
    CASE Message('Tag All Duplicates or Tag 1 of 2?','Tag Dups',,'Tag All|1 of 2|Cancel')
    OF 1 ; TagAll=1
    OF 3 ; EXIT
    END 
    CopyQueue(ExportQ,DupQ)
    loop QNx=RECORDS(DupQ) TO 1 BY -1
         GET(DupQ,QNx)
         DELETE(DupQ)   
         IF DupQ:TreeLevel<>2 THEN CYCLE. 
         GET(DupQ,DupQ:Symbol,DupQ:TreeLevel)  !Names are Case Sensitive so not DupQ:SymbolLwr
         IF ERRORCODE() THEN CYCLE. 
         CntDp += 1
         IF TagAll THEN 
            GET(ExportQ,POINTER(DupQ)) ; SetSearchFlagInExportQ(SearchFlag::Found,1,1)  
         END
         GET(ExportQ,QNx)              ; SetSearchFlagInExportQ(SearchFlag::Found,1,1)  
    end
    Message('Found ' & CntDp & ' duplcate symbol names.')
    EXIT
Accepted:LocateNext ROUTINE
    DATA
NextPrev SHORT    
QNdx LONG
Locate   PSTRING(64)
    CODE
  UPDATE(?LocatorText)  !RTL Bug: An ENTRY with SKIP will not get Accepted on Alt+&Key press so you have stale value
  IF ~LocatorText OR ~RECORDS(ExportQ) THEN EXIT.
  Locate=CLIP(LEFT(LOWER(LocatorText)))
  NextPrev = CHOOSE(?=?LocatePrev,-1,1)
  IF ?Sheet{PROP:ChoiceFEQ} = ?TAB:All THEN
     IF ~RECORDS(ExportQ) THEN EXIT.
     QNdx = CHOICE(?List1)
     IF ~QNdx AND NextPrev = -1 THEN QNdx=RECORDS(ExportQ)+1.
     LOOP
        QNdx += NextPrev
        GET(ExportQ, QNdx)
        IF ERRORCODE() THEN BREAK.
        IF INSTRING(Locate,ExportQ.SymbolLwr,1,1) > 0 THEN 
           SelectExportQ(?List1, ExportQ , QNdx)  !Expand Tree before select symbol  !?List1{PROP:Selected}=QNdx
           BREAK
        END
     END
  ELSE
     IF ~RECORDS(FilteredQ) THEN EXIT.
     QNdx = CHOICE(?Filtered)
     IF ~QNdx AND NextPrev = -1 THEN QNdx=RECORDS(FilteredQ)+1.
     LOOP
        QNdx += NextPrev
        GET(FilteredQ, QNdx)
        IF ERRORCODE() THEN BREAK.
        IF INSTRING(Locate,FilteredQ.SymbolLwr,1,1) > 0 THEN
           ?Filtered{PROP:Selected}=QNdx  ! SELECT(?Filtered, QNdx)
           BREAK
        END
     END  
  END
!EndRegion
!Region OnEvent ROUTINEs

OnExpandContract     ROUTINE     
    IF FIELD()<>?List1 THEN EXIT.   !CB fix? -->!TODO: Bug FilterdQ
    GET(ExportQ, 0+?List1{PROPLIST:MouseDownRow})
    ExportQ.treelevel = -ExportQ.treelevel
    ExportQ.icon = 3 - ExportQ.icon
    ExportQ.StyleNo = CHOOSE(ExportQ.Icon,ListStyle:Opened,ListStyle:Closed)
    PUT(ExportQ)
    DISPLAY(?list1)

OnAlertKey           ROUTINE       
   DATA
F  LONG
FileNmRef &STRING
IsADD BOOL
GotFile BOOL
   CODE
   F=FIELD()
   CASE F         !?SHEET{PROP:ChoiceFEQ}      !CB fix delete b ug
   OF ?AddFile OROF ?SubtractFile
      IsADD=CHOOSE(F=?AddFile) !Some messy code here to have Popup reuse existing routines and procedures
      IF IsADD THEN FileNmRef &= FileName ELSE FileNmRef &= SubFileName.
      CASE POPUPunder(F,'Windows DLLs...|Windows System 32 folder|WinSxS folder (Side by Side)|-|Clarion Folder...' & |
                      CHOOSE(~IsADD,'','|-|Load 13 DLLs to Create Clarion Win32.LIB'))
      OF 1 ; IF ~FileNamePopupSystem32(FileNmRef) THEN EXIT. ; GotFile=1
      OF 2 ; Filename=System32 ; POST(EVENT:Accepted,F) ; EXIT
      OF 3 ; IF ~WinSxSPicker(FileNmRef) THEN EXIT. ;  GotFile=1 
      OF 4 ; IF ~FileNamePopupClarion(FileNmRef) THEN EXIT. ; GotFile=1
      OF 5 ; DO LoadWin32LibDllsRtn ; EXIT
      END
      IF ~GotFile THEN EXIT.
      IF ExistsFile(FileNmRef) THEN
         IF IsADD THEN 
            DO FileAdded 
         ELSE 
            SubtractFileLIB(FileNmRef) ; DO EnableDisable ; SELECT(?LIST1)
         END
      ELSE
         POST(EVENT:Accepted,F)
      END

   OF ?List1            ! ?TAB:All
      GET(ExportQ, CHOICE(?List1)) 
      IF ERRORCODE() THEN EXIT.
      CASE KeyCode()
      OF CtrlC ; SETCLIPBOARD(ExportQ.Symbol)             !CB Copy
      OF MouseLeft2  OROF EnterKey OROF InsertKey         !CB Tag with Dbl Click or Enter Key
         IF ExportQ.TreeLevel<>2 THEN EXIT.
         IF ExportQ.Icon = 3 THEN !tagged now so untag
            SetSearchFlagInExportQ(SearchFlag::Default,1,0)
         ELSE
            SetSearchFlagInExportQ(SearchFlag::Found,1,0)  
         END
         PUT(ExportQ)

      OF DeleteKey
         DELETE(ExportQ)            ; SetSearchFlagInExportQ(SearchFlag::Default,1,0) 
         IF (ExportQ.treelevel<2)
           GET(ExportQ, CHOICE(?List1))
           LOOP WHILE (ExportQ.treelevel=2)
             DELETE(ExportQ)        ; SetSearchFlagInExportQ(SearchFlag::Default,1,0) 
             GET(ExportQ, CHOICE(?List1))
             IF ERRORCODE() THEN BREAK.
           END
         END        
         DISPLAY      
      END !CASE KeyCode()

   OF ?Filtered                 ! ?TAB:Filtered  !CB Filter Tab Deletes
      GET(FilteredQ, CHOICE(?Filtered))
      IF ERRORCODE() THEN EXIT.
      CASE KeyCode()
      OF CtrlC ; SETCLIPBOARD(FilteredQ.Symbol)
      OF MouseLeft2 OROF EnterKey                !CB Dbl click finds in Export Q or Popup "? Locate on All Exports Tab"
         ExportQ.symbol = FilteredQ.symbol
         GET(ExportQ,ExportQ.symbol)
         IF ~ERRORCODE() THEN 
             SelectExportQ(?List1, ExportQ ,POINTER(ExportQ))   !Expands Tree
             SELECT(?List1)
         END 
      OF DeleteKey
         DELETE(FilteredQ) 
         ExportQ.symbol = FilteredQ.symbol
         GET(ExportQ,ExportQ.symbol)
         IF ~ERRORCODE() THEN  SetSearchFlagInExportQ(SearchFlag::Default,1,1).   !UnTag
         DISPLAY        
      END !CASE KeyCode()
      
   OF ?LocatorText
      IF KEYCODE()=EnterKey THEN   !CB
         UPDATE
         POST(EVENT:Accepted, ?LocateNext)
      END
   OF 0
        IF KEYCODE()=CtrlShiftD THEN UNHIDE(?TAB:DebugWrite).
   END

OnNewSelection ROUTINE !CB
   CASE FIELD()         !?SHEET{PROP:ChoiceFEQ}      !CB fix delete b ug
   OF ?List1            ! ?TAB:All
      GET(ExportQ, CHOICE(?List1)) 
      IF ERRORCODE() THEN EXIT.
      CASE KeyCode()
      OF MouseRight  ; DO OnNewSelectionMouseRightList1 !CB
      END

   OF ?Filtered                 ! ?TAB:Filtered  !CB Filter Tab Deletes
      GET(FilteredQ, CHOICE(?Filtered))
      IF ERRORCODE() THEN EXIT.
      CASE KeyCode()
      OF MouseRight ; SETKEYCODE(0)
         CASE POPUP('[31763(700)]Symbol "' & CLIP(FilteredQ.symbol) &'" in "'& CLIP(FilteredQ.Module) &'"|-|' & |   !Header 1-1=0
                    '[' & PROP:Icon & '(~Found.ico)]UnTag / Remove from Filtered<9>Delete' &|
                    '|-|' & |
                    '[' & PROP:Icon & '(~Copy.ico)]Copy to Clipboard<9>Ctrl+C' & |
                    '|-|' & |
                    '[' & PROP:Icon & '('&ICON:VCRlocate &')]Locate on All Exports Tab<9>Enter') - 1
           OF 1 ; SETKEYCODE(DeleteKey) ; DO OnAlertKey 
           OF 2 ; SETCLIPBOARD(FilteredQ.Symbol) 
           OF 3 ; SETKEYCODE(EnterKey) ; DO OnAlertKey 
         END         
      END
   END

OnNewSelectionMouseRightList1 ROUTINE 
    DATA
PN  SHORT,AUTO
Lvl SHORT    
    CODE
    SETKEYCODE(0)
    PN=POPUP('[31763(700)]Symbol "' & CLIP(ExportQ.symbol) &'" in "'& CLIP(ExportQ.Module) &'"|-|' & |  !Header 1-1=0
             '[' & PROP:Icon & '(~Found.ico)]Tag Toggle<9>Enter' &|
             '|-|' & |
             '[' & PROP:Icon & '(~Copy.ico)]Copy to Clipboard<9>Ctrl+C' & |
             '|-|' & |
             '[' & PROP:Icon & '(~DelLine.ico)]Delete this item<9>Delete' & |
             '|-|Expand All|Collapse All' ) - 1

    CASE PN                   
      OF 1 ; SETKEYCODE(EnterKey) ; DO OnAlertKey  ; EXIT
      OF 2 ; SETCLIPBOARD(ExportQ.Symbol)          ; EXIT
      OF 3 ; SETKEYCODE(DeleteKey) ; DO OnAlertKey ; EXIT
      OF 4 ; Lvl=-1
      OF 5 ; Lvl=1
      ELSE ; EXIT
    END
    IF Lvl THEN 
       LOOP Ndx=1 TO RECORDS(ExportQ)
           GET(ExportQ,Ndx) 
           IF ExportQ.TreeLevel <> Lvl THEN CYCLE.
           ExportQ.TreeLevel *= -1
           ExportQ.Icon = 3 - ExportQ.Icon
           ExportQ.StyleNo = CHOOSE(ExportQ.Icon,ListStyle:Opened,ListStyle:Closed)
           PUT(ExportQ)
       END
       DISPLAY
    END

OnDrop               ROUTINE
    DATA    
Filez   &STRING
P1      LONG    
L1      LONG    
    CODE       
  FileName = LEFT(DropID())
!Drop Multi: [D:\Clarion91\Lib\ClaSCA.lib,D:\Clarion91\Lib\ClaSCAl.lib,D:\Clarion91\Lib\ClaSQA.lib,D:\Clarion91\Lib\ClaSQAl.lib   
  IF ~MATCH(FileName,',.:\\',Match:Regular) THEN  !CB Support Multi
      DO FileAdded
      EXIT
  END
  !Message('DropID has multiple!||[' & DropID() & ']','debug')  
  L1 = LEN(CLIP(DropID())) + 1
  FileZ &= NEW(STRING(L1))
  FileZ = LEFT(DropID())        !Rmv leading space
  LOOP 100 TIMES !Prevent hang
      P1=STRPOS(FileZ,',.:\\')   !Find: ,C:\
      IF ~P1 THEN P1=L1.
      FileName=LEFT(SUB(FileZ,1,P1-1))
      FileZ=LEFT(SUB(FileZ,P1+1,L1))
      IF EXISTS(FileName) THEN DO FileAdded.
  WHILE FileZ
  DISPOSE(FileZ) 
  SETDROPID('')

OnLocate             ROUTINE       
  BEEP()  !debugging beep (never happens--bug!)
  !note, ,VCR(?FindButton) doesn't work for me either
  Post(Event:Accepted,?FindButton)
!EndRegion OnEvent ROUTINEs
!EndRegion Accept Loop

!------------------------------------------------------
FileAdded            ROUTINE
  SetCursor(CURSOR:Wait)
  IF UPPER(RIGHT(FileName,4))='.LIB' THEN  !01/02/22 was: IF INSTRING('.LIB', UPPER(FileName), 1, 1)
    ReadLib()
  ELSE
    ReadExecutable()
  END
  Window{PROP:Text} = 'LibMaker - ' & CLIP(FileName)  !AB  !MG Note: This will only show that LAST file added, when multiple are added
  !PutIni('RecentFiles','Read',FileName,qINI_File)
  SetCursor()
!========================================================================================
!========================================================================================
ReadExecutable       PROCEDURE !gets export table from 16 or 32-bit file or LIB file
sectheaders ULONG   ! File offset to section headers
sections    USHORT  ! File offset to section headers
VAexport    ULONG   ! Virtual address of export table, according to data directory
                    ! This is used as an alternative way to find table if .edata not found
!IMAGE_NT_OPTIONAL_HDR32_MAGIC   equate(10Bh)   ! 267
IMAGE_NT_OPTIONAL_HDR64_MAGIC   equate(20Bh)   ! 523
   CODE
   OPEN(EXEfile, 0) 
   IF ERRORCODE() THEN MESSAGE('Error ' & ErrorCode() &' '& ERROR() &'|Opening File ' & NAME(ExeFile)).
   GET(EXEfile, 1, SIZE(EXE:DOSheader))
   IF EXE:dos_magic = 'MZ' THEN
     newoffset = EXE:dos_lfanew
     GET(EXEfile, newoffset+1, SIZE(EXE:PEheader))
     IF EXE:pe_signature = 04550H THEN
       sectheaders = EXE:pe_optsize+newoffset+SIZE(EXE:PEheader)
       sections = EXE:pe_nsect
       ! Read the "Optional header"
       GET(EXEfile, newoffset+SIZE(EXE:PEheader)+1, SIZE(EXE:Optheader))

       IF EXE:OptHeader.opt_Magic = IMAGE_NT_OPTIONAL_HDR64_MAGIC THEN      !CB 
          IF MESSAGE('This is a 64-bit file which LibMaker cannot load.' & |
                     '||File: '&CLIP(name(exefile)),    |
                     '64-Bit Binary', ICON:Exclamation, |
                      BUTTON:Abort+BUTTON:Ignore)=BUTTON:Abort THEN
             CLOSE(EXEfile)
             RETURN
          END
       END      
       
       IF EXE:opt_DataDirNum THEN
          ! First data directory describes where to find export table
          GET(EXEfile, newoffset+SIZE(EXE:PEheader)+SIZE(EXE:OptHeader)+1,SIZE(EXE:DataDir))
          VAexport = EXE:data_VirtualAddr
       END

       LOOP i# = 1 TO sections
         GET(EXEfile,sectheaders+1,SIZE(EXE:sectheader))
         sectheaders += SIZE(EXE:sectheader)
         IF EXE:sh_SectName = '.edata' THEN
            DumpPEExportTable(EXE:sh_VirtAddr, EXE:sh_VirtAddr - EXE:sh_rawptr)
         ELSIF EXE:sh_VirtAddr <= VAexport AND |
               EXE:sh_VirtAddr+EXE:sh_RawSize > VAexport
            DumpPEExportTable(VAexport, EXE:sh_VirtAddr - EXE:sh_rawptr)
         END
       END
   ELSIF 1=Message('This is a 16-bit DLL.','16-Bit',,'Load|Cancel') THEN  ! dos_magic <> 'MZ
       GET(EXEfile, newoffset+1, SIZE(EXE:NEheader))
       DumpNEExports
     END
   END
   CLOSE(EXEfile)

!========================================================================================
DumpPEExportTable    PROCEDURE(VirtualAddress, ImageBase) !gets export table from a PE format file (32-bit)

NumNames  ULONG
Names     ULONG
Ordinals  ULONG
Base      ULONG
j         LONG
   CODE
   GET(EXEfile, VirtualAddress-ImageBase+1, SIZE(EXE:ExpDirectory))
   NumNames      = EXE:exp_NumNames
   Names         = EXE:exp_AddrNames
   Ordinals      = EXE:exp_AddrOrds
   Base          = EXE:exp_Base
   GET(EXEfile, EXE:exp_Name-ImageBase+1, SIZE(EXE:cstringval))
   !Added code to parse the first character of the module name as when a .Net unmanaged dll
   !A \ seems to be generated as first character
   if EXE:cstringval[1] = '\' THEN
	 EXE:cstringval = EXE:cstringval[2:SIZE(EXE:cstringval)]
   END
   ExportQ.Module    = EXE:cstringval
   ExportQ.Symbol    = EXE:cstringval
   ExportQ.SymbolLwr = lower(EXE:cstringval)    !CB 
   ExportQ.treelevel = 1
   ExportQ.icon      = 1
   ExportQ.ordinal   = 0
   OrgOrder += 1 ; ExportQ.orgorder = OrgOrder  !AB
   ExportQ.SearchFlag= SearchFlag::Default
   ExportQ.StyleNo = ListStyle:Opened
   ADD(ExportQ)

   ExportQ.treelevel = 2
   ExportQ.icon      = 0
   ExportQ.StyleNo = ListStyle:None
   LOOP j = 0 TO NumNames-1
      GET(EXEfile, Names    + j*4 - ImageBase+1, SIZE(EXE:ulongval))
      GET(EXEfile, EXE:ulongval   - ImageBase+1, SIZE(EXE:cstringval)) 
      ExportQ.symbol    = EXE:cstringval
      ExportQ.symbolLwr = lower(EXE:cstringval)
      GET(EXEfile, Ordinals + j*2 - ImageBase+1, SIZE(EXE:ushortval))
      ExportQ.ordinal = EXE:ushortval+Base
      IF SkipAddExportQ() THEN CYCLE.  !CB
      OrgOrder += 1 ; ExportQ.orgorder = OrgOrder  !AB
      ADD(ExportQ)
      IF ExportQ.SearchFlag THEN SetSearchFlagInExportQ(SearchFlag::Default,0,0).  !undo tag by SkipAddExport
   END

!========================================================================================
DumpNEExports   PROCEDURE ! DumpNEexports gets export table from a NE format file (16-bit)
j  LONG
r  LONG
   CODE
! First get the module name - stored as first entry in resident name table
   j =             EXE:ne_nrestab+1
   r = newoffset + EXE:ne_restab+1
   GET(EXEfile, r, SIZE(EXE:pstringval))
   ExportQ.Module     = EXE:pstringval
   ExportQ.symbol     = EXE:pstringval
   ExportQ.symbolLwr  = lower(EXE:pstringval)
   ExportQ.ordinal    = 0
   ExportQ.treelevel  = 1
   ExportQ.icon       = 1
   OrgOrder += 1 ; ExportQ.orgorder = OrgOrder  !AB
   ExportQ.SearchFlag = SearchFlag::Default
   ExportQ.StyleNo = ListStyle:Opened
   ADD(ExportQ)
   r += LEN(EXE:pstringval)+1    !move past module name
   r += 2                        !move past ord#

! Now pull apart the resident name table. First entry is the module name, read above
   ExportQ.treelevel = 2
   ExportQ.icon      = 0
   ExportQ.StyleNo = ListStyle:None
   LOOP
     GET(EXEfile, r, SIZE(EXE:pstringval))
     IF LEN(EXE:pstringval)=0 THEN
       BREAK
     END
     ExportQ.symbol    = EXE:pstringval
     ExportQ.symbolLwr = lower(EXE:pstringval)
     r += LEN(EXE:pstringval)+1
     GET(EXEfile, r, SIZE(EXE:ushortval))
     r += 2
     ExportQ.ordinal = EXE:ushortval
     IF SkipAddExportQ() THEN CYCLE. !CB
     OrgOrder += 1 ; ExportQ.orgorder = OrgOrder  !AB     
     ADD(ExportQ)
     IF ExportQ.SearchFlag THEN SetSearchFlagInExportQ(SearchFlag::Default,0,0).  !undo tag by SkipAddExport
   END

! Now pull apart the non-resident name table. First entry is the description, and is skipped
   GET(EXEfile, j, SIZE(EXE:pstringval)) ;    j += LEN(EXE:pstringval)+1
   GET(EXEfile, j, SIZE(EXE:ushortval))  ;    j += 2
   LOOP
     GET(EXEfile, j, SIZE(EXE:pstringval))
     IF LEN(EXE:pstringval)=0 THEN
       BREAK
     END
     ExportQ.symbol    = EXE:pstringval
     ExportQ.symbolLwr = lower(EXE:pstringval)
     j += LEN(EXE:pstringval)+1
     GET(EXEfile, j, SIZE(EXE:ushortval))
     j += 2
     ExportQ.ordinal = EXE:ushortval 
     IF SkipAddExportQ() THEN CYCLE. !CB
     OrgOrder += 1 ; ExportQ.orgorder = OrgOrder  !AB
     ADD(ExportQ)
     IF ExportQ.SearchFlag THEN SetSearchFlagInExportQ(SearchFlag::Default,0,0).  !undo tag by SkipAddExport
   END

!========================================================================================
WriteLib        PROCEDURE !writes out all info in the export Q to a LIB file
i  ULONG !USHORT !changed 7/24/02, based on note in comp.lang.clarion
   CODE
   CREATE(LIBfile)
   OPEN(LIBfile)
   LOOP i = 1 TO RECORDS(ExportQ)
      GET(ExportQ, i)
      IF ExportQ.treelevel=2 THEN 
         IF Glo.WriteFiltered AND glo.FoundCount THEN  !CB Save Filtered
            FLTQ:OrgOrder=ExportQ.OrgOrder
            GET(FilteredQ,FLTQ:OrgOrder)
            IF ERRORCODE() THEN CYCLE.
         END      
        ! Record size is length of the strings, plus two length bytes, a two byte
        ! ordinal, plus the header length (excluding the first three bytes)
        LIB:typ     = 88H
        LIB:kind    = 0A000H
        LIB:bla     = 1
        LIB:ordflag = 1
        LIB:len = LEN(CLIP(ExportQ.module))+LEN(CLIP(ExportQ.symbol))+2+2+SIZE(LIB:header)-3 
                                                  ADD(LIBfile, SIZE(LIB:header))
        LIB:pstringval = CLIP(ExportQ.symbol) ;   ADD(LIBfile, LEN(LIB:pstringval)+1)
        LIB:pstringval = CLIP(ExportQ.module) ;   ADD(LIBfile, LEN(LIB:pstringval)+1)
        LIB:ushortval  =      ExportQ.ordinal ;   ADD(LIBfile, SIZE(LIB:ushortval))
      END
   END
   CLOSE(LIBfile)

!========================================================================================
ReadLib             PROCEDURE !reads back in a LIB file output by WriteLib above or by IMPLIB etc

!Region  '****** softvelocity.products.c55ee.bugs posting ******')
!@   > I used LibMaker to list the contents of Win32.lib .  I know that the
!@   > Registry APIs are in the lib but LibMaker does not list them.  Is there
!@   > a bug in LibMaker?
!@
!@   Yes, there is a stupid bug: the ReadLib function uses variables of USHORT
!@   type as pointers in the LIB file. As result, files with size over 64KB
!@   can't be processed correctly. I think, LibMaker's sources are provided
!@   among examples. Therefore you can fix that bug yourself: just change types
!@   of ii and jj variables local for ReadLib from USHORT to LONG and remake
!@   the program.
!@
!@   Alexey Solovjev
!
!  MG Note: Need to review c55\examples\src\libmaker\libmaker.clw
!           pay attention to line just below "if modulename <> lastmodule"
!           Also notice the NEW ExportQ.Libno field
!EndRegion 

i          LONG   !changed 12/20/00 due to Alexey Solovjev Posting USHORT
j          LONG   !changed 12/20/00 due to Alexey Solovjev Posting USHORT
lastmodule STRING(20)
modulename STRING(20)
symbolname STRING(128)
ordinal    USHORT
   CODE
   OPEN(LIBfile, 0)
   i = 1
   ExportQ.SearchFlag= SearchFlag::Default
   LOOP
      GET(LIBfile, i, SIZE(LIB:header))     ! Read next OMF record
      IF ERRORCODE() OR LIB:typ = 0 OR LIB:len = 0 THEN
         BREAK                              ! All done
      END
      j = i + SIZE(LIB:header)              ! Read export info from here
      i = i + LIB:len + 3                   ! Read next OMF record from here
      IF LIB:typ = 88H AND LIB:kind = 0A000H AND LIB:bla = 1 AND LIB:ordflag = 1 THEN
          GET(LIBfile, j, SIZE(LIB:pstringval))
          symbolname = LIB:pstringval
          j += LEN(LIB:Pstringval)+1
          GET(LIBfile, j, SIZE(LIB:pstringval))
          modulename = LIB:pstringval
          j += LEN(LIB:Pstringval)+1
          GET(LIBfile, j, SIZE(LIB:ushortval))
          ordinal = LIB:ushortval
          IF modulename <> lastmodule      ! A LIB can describe multiple DLLs
             lastmodule = modulename
             ExportQ.treelevel = 1
             ExportQ.icon = 1
             ExportQ.symbol = modulename
             ExportQ.symbolLwr = lower(modulename)
             ExportQ.module = modulename
             ExportQ.ordinal = 0
             OrgOrder += 1 ; ExportQ.orgorder = OrgOrder  !AB
             ExportQ.StyleNo = ListStyle:Opened
             ExportQ.SearchFlag= SearchFlag::Default
             ADD(ExportQ)
             ExportQ.StyleNo = ListStyle:None
          END
          ExportQ.treelevel = 2
          ExportQ.icon = 0
          ExportQ.symbol    = symbolname
          ExportQ.symbolLwr = lower(symbolname)
          ExportQ.module = modulename
          ExportQ.ordinal = ordinal
          IF SkipAddExportQ() THEN CYCLE.  !CB
          OrgOrder += 1 ; ExportQ.orgorder = OrgOrder  !AB
          ADD(ExportQ)
          IF ExportQ.SearchFlag THEN SetSearchFlagInExportQ(SearchFlag::Default,0,0).  !undo tag by SkipAddExport
      END
   END
   CLOSE(LIBfile)

!========================================================================================
InfoWindow PROCEDURE()
ExeName0 STRING(260),AUTO 
infowin WINDOW('About LibMaker'),AT(,,234,200),GRAY,SYSTEM,FONT('Segoe UI',9,,FONT:regular),PALETTE(256)
        PANEL,AT(6,6,74,86),USE(?Panel1),BEVEL(5)
        IMAGE('AB.JPG'),AT(14,14),USE(?Image1)
        GROUP,AT(85,6,139,86),USE(?Group),FONT(,,COLOR:CAPTIONTEXT),COLOR(COLOR:Black)
            BOX,AT(85,6,139,86),USE(?Box1),COLOR(COLOR:Black),FILL(COLOR:INACTIVECAPTION),LINEWIDTH(2)
            STRING('LibMaker'),AT(87,9,135,16),USE(?String1),TRN,CENTER,FONT('Arial',14,,FONT:bold)
            STRING('Release:'),AT(89,31,35),USE(?ReleasePmt),TRN,RIGHT
            STRING('yyyy.mm.dd by Release Info'),AT(132,31),USE(?ReleaseString),TRN
            STRING('Creator:'),AT(89,50,35),USE(?String3),TRN,RIGHT
            STRING('Arnor Baldvinsson'),AT(132,50),USE(?String4),TRN
            STRING('E-mail:'),AT(89,63,35),USE(?String7),TRN,RIGHT
            STRING('<<Arnor@IceTips.com>'),AT(132,63),USE(?String8),TRN
            STRING('http://www.IceTips.com'),AT(132,76,76,10),USE(?String10),TRN
        END
        STRING('Modifications by:'),AT(4,102),USE(?String9)
        STRING('Mark Goldberg, Mark Sarson & Carl Barnes'),AT(60,102),FULL,USE(?String11), |
                FONT(,,,FONT:bold)
        BUTTON('&Close'),AT(185,161,45,16),USE(?Button1),STD(STD:Close),ICON('Exit.ico'),DEFAULT,LEFT
        STRING('TIP:'),AT(12,115),USE(?TipStr1),FONT(,,,FONT:bold)
        PROMPT('Right-Click on Add / Subtract buttons and Lists for Popups'),AT(33,115),USE(?TipPmt1)
        STRING('TIP:'),AT(12,127),USE(?TipStr2),FONT(,,,FONT:bold)
        PROMPT('You can drag a DLL from Explorer and drop on LibMaker'),AT(33,127),USE(?TipPmt2)
        STRING('TIP:'),AT(12,140,13,10),USE(?TipStr3),FONT(,,,FONT:bold)
        PROMPT('Command line arguments:<13,10>    READ="FileName"<13,10>    WRITE="FileName"<13,10> ' & |
                '   /CLOSE'),AT(33,140,130,36),USE(?TipPmt3)
        TEXT,AT(4,182,226,18),USE(ExeName0),TRN
    END
   CODE
   ExeName0=COMMAND('0')
   OPEN(infowin)
   ?ReleaseString{PROP:Text}=ReleaseInfoString
   ACCEPT 
   END
!========================================================================================
SelectQBE    PROCEDURE  !most of procedure written by Carl Barnes
   !Updates:
   !  08/11/19 CB
   !      Method ADDED Ends With, Wild, RegEx. Case Senstive
   !      See List of Found then allow Tag, Untag, Delete. Can delete from Find List with Delete key
   !FYI a problem with Wild Cards is they are not contrains unless * on Begin and End

Lcl:SearchString    STRING(100),STATIC
Lcl:Method          BYTE,STATIC
Lcl:CaseSenstive    BYTE,STATIC
Lcl     group,PRE(Lcl)
SearchOption    BYTE
ExportQ_Rec     Long
SearchFlag      like(ExportQ.SearchFlag)
UntagCnt        Long
        end
QNdx    LONG,AUTO
TagBtn  LONG,AUTO
FndCnt  LONG,AUTO
FindQ   QUEUE(ExportQ),PRE(FindQ)
PtrExpQ     LONG
        END
Swindow WINDOW('Search Symbols'),AT(,,240,180),GRAY,IMM,SYSTEM,FONT('Segoe UI',9,,FONT:regular),RESIZE
        PROMPT('&Find Text:'),AT(3,4,38),USE(?Find:Prompt)
        ENTRY(@s100),AT(36,4,200,11),USE(Lcl:SearchString)
        PROMPT('&Method:'),AT(3,20,38),USE(?Meth:Prompt)
        LIST,AT(36,19,78,11),USE(Lcl:Method),TIP('For Wild cards don''t forget * Begin/End * to find' & |
                ' any where<13,10,13,10>RegEx is Clarion MATCH(,,Match:Regular) command.'),DROP(9), |
                FROM('Contains|#1|Starts With|#2|Ends With|#3|Wild *? Cards ?*|#4|Regular Expression|#5')
        CHECK('Case'),AT(118,20),USE(Lcl:CaseSenstive),SKIP,TIP('Case Sensitive')
        BUTTON('&Search'),AT(151,18,46,15),USE(?SearchButton),ICON('BINOCULR.ICO'),DEFAULT,LEFT
        BUTTON('&Cancel'),AT(201,18,36,15),USE(?CloseButton),STD(STD:Close)
        PANEL,AT(4,36,232,2),USE(?PANEL1),BEVEL(0,0,0600H)
        GROUP,AT(0,38,236,143),USE(?ResulsGrp),DISABLE
            BUTTON('&Tag'),AT(4,41,48,15),USE(?TagBtn),ICON('found.ico'),TIP('Tag the found symbols ' & |
                    'in the All List.'),LEFT
            BUTTON('&UnTag'),AT(58,41,48,15),USE(?UnTagBtn),ICON('UnTag.ico'),TIP('UnTag the found s' & |
                    'ymbols in the All List.'),LEFT
            BUTTON('&Delete'),AT(113,41,48,15),USE(?DeleteBtn),ICON('DelLine.ico'),TIP('Delete the f' & |
                    'ound symbols from the All List.'),LEFT
            LIST,AT(4,60,232,115),USE(?List:FindQ),VSCROLL,FROM(FindQ),FORMAT('143L(2)FIY~Symbols Fo' & |
                    'und by Search~@s128@'),ALRT(DeleteKey),ALRT(CtrlC)
        END
        BUTTON('RgX'),AT(207,44,19,12),USE(?RegExBtn),TIP('Syntax of Clarion Match RegEx Operators')
    END
                      !FORMAT('143L(2)F*I~Symbols Found by Search~')
MGResizeCls ResizeClassType
Pz LONG,DIM(4),STATIC    
  CODE
  IF ~Lcl:Method THEN   !CB made Static, so other MGs running do not mess up
     Lcl:SearchString = GetIni('Search','For2'   , ,qINI_File)
     Lcl:Method       = GetIni('Search','Method2',1,qINI_File)
     Lcl:CaseSenstive = GetIni('Search','Case2',  0,qINI_File)
  END  
  OPEN(Swindow)
  0{PROP:minheight}  = .7 * 0{PROP:height}
  0{PROP:minwidth }  = 0{PROP:width}
   MGResizeCls.Init_Class(Swindow)
   MGResizeCls.Add_ResizeQ(?List:FindQ,'L/T R/B')  
  ?List:FindQ{PROP:iconlist, 3} = '~Found.ico'
  IF Pz[3]>9 THEN 
     SETPOSITION(0,Pz[1],Pz[2],Pz[3],Pz[4])
     POST(EVENT:Sized)
  END
  ACCEPT
    case EVENT()
    of EVENT:Sized ; MGResizeCls.Perform_Resize()
    end
    case ACCEPTED()
      of ?SearchButton 
         IF ~Lcl:SearchString THEN SELECT(?Lcl:SearchString) ; CYCLE .
         DO SearchQueue ; DO EnableRtn
      of ?TagBtn orof ?UnTagBtn orof ?DeleteBtn
         TagBtn = ACCEPTED()
         DO Update_ExportQ
         BREAK
      of ?RegExBtn ; START(RegExHelper,,0{PROP:XPos},0{PROP:YPos},0{PROP:Width})
    end  !case Accepted 
    CASE FIELD()
    OF ?List:FindQ 
       GET(FindQ,CHOICE(?List:FindQ))
       CASE EVENT()
       OF EVENT:AlertKey 
          CASE KEYCODE()
          OF CtrlC     ; SetClipboard(FindQ:Symbol)
          OF DeleteKey ; DELETE(FindQ) ; DO EnableRtn
          END
       OF EVENT:NewSelection
          IF KEYCODE()=MouseRight THEN 
             SetKeyCode(0)
             CASE POPUP('Copy Symbol Name<9>Ctrl+C|-|Delete Symbol<9>Delete')
             OF 1 ; SetKeyCode(CtrlC)     ; POST(EVENT:AlertKey,?)
             OF 2 ; SetKeyCode(DeleteKey) ; POST(EVENT:AlertKey,?)
             END
          END
       END
    END
  END !Accept Loop
  MGResizeCls.Close_Class()
  PutIni('Search','For2'   ,Lcl:SearchString,qINI_File)
  PutIni('Search','Method2',Lcl:Method,qINI_File)
  PutIni('Search','Case2',Lcl:CaseSenstive,qINI_File)
  GETPOSITION(0,Pz[1],Pz[2],Pz[3],Pz[4])
  RETURN
EnableRtn ROUTINE
     FndCnt = RECORDS(FindQ)
     ?ResulsGrp{PROP:Disable}=CHOOSE(~FndCnt)
     ?UnTagBtn{PROP:Disable}=CHOOSE(Lcl:UntagCnt=0) 
     ?List:FindQ{PROPLIST:Header,1} = CHOOSE(~FndCnt,'',FndCnt&' ') & 'Symbols Found by Search'     
     DISPLAY 
!-------------------------------------------------------------
SearchQueue        ROUTINE  !of SelectQBE
    DATA
MatchString     Pstring(Size(Lcl:SearchString)+1)  !Use Pstring, so don't have to clip inside of "contains" loop
SymbolString    &STRING
LenMS           uShort,auto !Length of SearchStrign
LenSym          uShort,auto !Length of Symbol
FoundIt         BOOL
      CODE
  FREE(FindQ) ; Lcl:UntagCnt=0
  IF ~Lcl:CaseSenstive THEN
     MatchString = clip(lower(lcl:SearchString)) 
     SymbolString &= ExportQ.SymbolLwr
  ELSE
     MatchString = clip(lcl:SearchString)
     SymbolString &= ExportQ.Symbol
  END
  LenMS=LEN(MatchString)  
  !message(Lcl:Method &'|SymbolString='& SymbolString &'|MatchString=' & MatchString)
  LOOP QNdx=1 TO RECORDS(ExportQ)
       GET(ExportQ,QNdx)
       IF ExportQ.treelevel <> 2 THEN CYCLE.
       CASE Lcl:Method
       OF 1 ; FoundIt=Instring(MatchString,SymbolString,1,1)             !Contains 
       OF 2 ; FoundIt = CHOOSE(SymbolString[1 : LenMS] = MatchString)    !Starts With 
       OF 3 ; LenSym=LEN(CLIP(SymbolString))                             !Ends With
              FoundIt = CHOOSE(LenSym>=LenMS AND SymbolString[LenSym-LenMS+1 : LenSym] = MatchString)
       OF 4 ; FoundIt=MATCH(SymbolString,MatchString,Match:Wild)      !Wild Cards
       OF 5 ; FoundIt=MATCH(SymbolString,MatchString,Match:Regular)   !Regular
       END
       IF FoundIt THEN
          FindQ = ExportQ 
          FindQ:PtrExpQ = QNdx
          ADD(FindQ,FindQ:SymbolLwr) 
          IF ExportQ.SearchFlag = SearchFlag::Found THEN Lcl:UntagCnt += 1.
       END
  END
  EXIT
!-------------------------------------------------------------
Update_ExportQ  ROUTINE  !of SelectQBE, called by SearchQueue
    SORT(FindQ,-FindQ:PtrExpQ)      !Must do in reverse for delete so numbers don't chnage
    LOOP QNdx=1 TO RECORDS(FindQ)
         GET(FindQ,QNdx)
         GET(ExportQ,FindQ:PtrExpQ) ; IF ERRORCODE() THEN CYCLE.
         CASE TagBtn
          OF ?TagBtn     ; SetSearchFlagInExportQ(SearchFlag::Found,1,1)
          OF ?UnTagBtn   ; SetSearchFlagInExportQ(SearchFlag::Default,1,1) 
          OF ?DeleteBtn
             IF ExportQ.SearchFlag = SearchFlag::Found THEN Glo:FoundCount -= 1.
             DELETE(ExportQ)
         END    
    END
    EXIT
!=======================================================================    
SetSearchFlagInExportQ  PROCEDURE(BYTE NewSearchFlag, BYTE SyncFoundCount, BYTE PutExportQ)
    CODE
    IF ExportQ.SearchFlag = NewSearchFlag THEN RETURN. !Already what we want  
    IF ExportQ.TreeLevel <> 2 THEN RETURN. !Not a file?
    CASE NewSearchFlag
    OF SearchFlag::Found    !I.e.=1
        IF SyncFoundCount THEN Glo:FoundCount += 1.
        ExportQ.SearchFlag  = SearchFlag::Found
        ExportQ.StyleNo = ListStyle:Found
        ExportQ.Icon     = 3 !found
        IF PutExportQ THEN Put(ExportQ).

    OF SearchFlag::Default  !I.e.=0
        IF SyncFoundCount THEN Glo:FoundCount -= 1.
        ExportQ.SearchFlag  = SearchFlag::Default
        ExportQ.StyleNo = ListStyle:None
        ExportQ.Icon     = 0            
        IF PutExportQ THEN Put(ExportQ).
   END
   RETURN
!=======================================================================
SortExportQ  PROCEDURE(BYTE ForcedOrder=0) !CB Put all sorting here
    CODE
    IF 0=ForcedOrder THEN ForcedOrder = Glo.SortOrder .
    Execute ForcedOrder
      BEGIN; SORT(ExportQ,ExportQ.OrgOrder)                                    ; SORT(FilteredQ,FLTQ:orgorder)                           ; END
      BEGIN; SORT(ExportQ,ExportQ.Module,ExportQ.TreeLevel,ExportQ.symbolLwr)  ; SORT(FilteredQ,FLTQ:Module,FLTQ:TreeLevel,FLTQ:symbolLwr)  ; END
      BEGIN; SORT(ExportQ,ExportQ.Module,ExportQ.TreeLevel,ExportQ.Ordinal)    ; SORT(FilteredQ,FLTQ:Module,FLTQ:TreeLevel,FLTQ:Ordinal) ; END
    END
    RETURN
!=======================================================================
ExportQ_to_FilterQ   PROCEDURE()
X LONG,AUTO
PutIt BYTE,AUTO
  CODE
  !? Should I just Keep the Tree in Filtered ... Too much code to review, e.g. Found=Records
  SortExportQ(1) !Must be in Original Order to gen WriteExpQ
  ExportQCompress(ExportQ)
  FREE(WriteExpQ) ; CLEAR(WriteExpQ)
  FREE(FilteredQ)
  LOOP X = 1 TO RECORDS(ExportQ)
     GET(ExportQ, X)
     CASE ExportQ.Icon
     OF 1 OROF 2
        PutIT=CHOOSE(WriteExpQ.Icon <> 3 AND X > 1)  !Did I add a Library but no Functions
        WriteExpQ = ExportQ 
        WriteExpQ.TreeLevel = ABS(WriteExpQ.TreeLevel) 
        IF PutIT THEN PUT(WriteExpQ) ELSE ADD(WriteExpQ).

     OF 3 !Found  !0=Not FOund
        FilteredQ = ExportQ ; ADD(FilteredQ)
        WriteExpQ = ExportQ ; ADD(WriteExpQ)
     END  
  END
  IF WriteExpQ.Icon <> 3 THEN DELETE(WriteExpQ).  !Widowed Library
  Glo:FoundCount = RECORDS(FilteredQ)  !CB fix any buggy counts
  SortExportQ()  !Sort as user has it
  RETURN
!========================================================================================
ExportQCompress PROCEDURE(ExportQ ExpQ2Do) !CB Many ways to Delete so want to remove empty parents
QRec          LONG,AUTO
LastTreeLevel SHORT(0)
   CODE   
   QRec=1
   LOOP
      GET(ExpQ2Do, QRec) 
      IF ERRORCODE() THEN
         IF ABS(LastTreeLevel)=1 THEN
            GET(ExpQ2Do, QRec-1) ; DELETE(ExpQ2Do)  !GET Last and Delete
         END
         BREAK
      END
      IF  ABS(ExpQ2Do.treelevel)=1 |       !This is a Level 1 Module
      AND ABS(    LastTreeLevel)=1 THEN    !And Last Record was Level 1., so there's nothing inbetween
          QRec -= 1             !Move Q Index back to LAST
          GET(ExpQ2Do, QRec)    !GET Last
          DELETE(ExpQ2Do)       !Delete Last it was an Empty Module
          GET(ExpQ2Do, QRec)    !So now GET QRec is same record we were on
      END
      LastTreeLevel = ExpQ2Do.TreeLevel
      QRec += 1
   END
   RETURN
!EndRegion - Tmp Region

!========================================================================================
SelectExportQ  PROCEDURE(LONG ListFEQ, ExportQ ExpQ2Do, LONG QPointer)  !Expand Tree before select symbol
!Selecting a line in a Collapsed Tree and a Scroll seems to be trouble 
QX LONG,AUTO
    CODE
    LOOP QX=QPointer TO 1 BY -1
        GET(ExpQ2Do,QX)     
        CASE ExpQ2Do.TreeLevel
        OF 1  ; BREAK 
        OF -1   
               ExportQ.TreeLevel = 1
               ExportQ.Icon = 3 - ExportQ.Icon
               ExportQ.StyleNo = CHOOSE(ExportQ.Icon,ListStyle:Opened,ListStyle:Closed)
               PUT(ExportQ)
               ListFEQ{PROP:Selected}=QX
               DISPLAY      !Hope to Get Scrollbar?
               BREAK
        END
    END
    ListFEQ{PROP:Selected}=QPointer
    RETURN 
    
!========================================================================================
GenerateMap      PROCEDURE(BYTE argRonsFormat) !writes out all info in the export Q to an ASCII file, to aid in building CW Map statements
  !Created 10/6/98 by Monolith Custom Computing, Inc.  
  !todo, review whole ExportQ, and indent to maximum for a prettier file

  !Consider support for Ron Schofield's .API file format
  ! <header>
  ! <prototypes>
  !
  ! a single header is:
  ! API Name|Library|DLL|Equate File|Structure File
  !
  ! multiple prototypes, one per line:   one is:
  ! Procedure|Prototype|Proc Attributes|Proc Name

lcl             group
exq_rec            USHORT
symbol             like(ExportQ.symbol)!,auto
LenModName         ushort
CurrModule         like(ExportQ.Module) !no auto
Assumed_Attributes string(',pascal,raw,dll(1)')
MaxSymLen   LONG
                end

   CODE
   CLEAR(FileName)
   IF ~FileDialog('Save Clarion Map definition as ...', FileName, |
       'CW Source and Includes (*.clw,*.inc)|*.clw;*.inc|Source Only|*.clw|Includes Only|*.inc|All|*.*',FILE:Save+FILE:LongName)
       RETURN
    END
   
   lcl.MaxSymLen = LongestSymbol(WriteExpQ) !in ExportQ Level 2   

   CREATE(Asciifile)
   OPEN  (Asciifile)
   LOOP lcl.exq_rec = 1 TO RECORDS(WriteExpQ)
      GET(WriteExpQ, lcl.exq_rec)
      case ABS(WriteExpQ.Treelevel)    !CB added ABS
        of 1 ; DO TreeLevel1
        of 2
              !CB using WriteExpQ so no need to check Filtered
              !IF Glo.WriteFiltered AND glo.FoundCount THEN  !CB Save Filtered
              !   FLTQ:OrgOrder=WriteExpQ.OrgOrder
              !   GET(FilteredQ,FLTQ:OrgOrder)
              !   IF ERRORCODE() THEN CYCLE.
              !END      
              DO TreeLevel2
      end 
   END

   if argRonsFormat
        Ascii:Line = ''
   else Ascii:Line = '<32>{5}end !module(''' & clip(lcl.CurrModule) & ''')'     !Should be what we saved the .LIB as
   end
   
   Add  (ASCIIFile)
   CLOSE(Asciifile)

   IF MESSAGE('Would you like to view the MAP file with Notepad?||'&FileName,'View Map',ICON:Question,Button:Yes+Button:No)=Button:Yes
      RUN('notepad ' & CLIP(FileName))
   END

TreeLevel1 ROUTINE        
    if lcl.exq_rec >1
       if argRonsFormat
            Ascii:Line = ''
       else Ascii:Line = '<32>{5}end !module(''' & clip(lcl.CurrModule) & ''')'     !Should be what we saved the .LIB as
       end
       Add(ASCIIFile)
       Ascii:Line = ''
       Add(ASCIIFile)
    end
    !---------- replace the .DLL or .EXE module extension with .LIB
    lcl.CurrModule = WriteExpQ.module
    lcl.LenModName = len(clip(lcl.CurrModule))   !01/02/22 revised as Instring is wrong incase name is xxxx.dll.lib
    case upper(sub(lcl.CurrModule,lcl.LenModName-3,4))
    of   '.DLL'
    orof '.EXE'
                lcl.CurrModule[lcl.LenModName - 3 : lcl.LenModName] = '.lib'
    end
    !---------- replace the .DLL or .EXE module extension with .LIB -end
    if argRonsFormat
         ! API Name|Library|DLL|Equate File|Structure File
         Ascii:Line = 'Todo Header Line: API Name|Library|DLL|Equate File|Structure File'  !<--- todo
    else Ascii:Line = '<32>{5}module(''' & clip(lcl.CurrModule) & ''')'             !Should be what we saved the .LIB as
    end
    Add(ASCIIFile)

TreeLevel2 ROUTINE
    lcl.symbol = CleanSymbol(WriteExpQ.symbol)
    if argRonsFormat
         ! Procedure|Prototype|Proc Attributes|Proc Name
         Ascii:Line =             clip(lcl:symbol) &'|<32>{40}|'& clip(lcl.Assumed_Attributes) & '|'   & clip(WriteExpQ.symbol)
    else Ascii:Line = '<32>{8}' & clip(lcl.symbol) & all('<32>',lcl.MaxSymLen - len(clip(lcl.Symbol))) & '(<9>{4}),pascal,raw,dll(1),name(''' & clip(WriteExpQ.symbol) & ''')'
    end
    Add(ASCIIFile)
    EXIT
!========================================================================================
CleanSymbol   PROCEDURE(STRING xSymbol)!,STRING
AtIsAt   LONG,AUTO
   CODE   
   LOOP
      CASE xSymbol[1]
        OF '_'  
      OROF '?'  
      OROF '@'
              xSymbol = xSymbol[ 2 : SIZE(xSymbol) ]
      ELSE    
              BREAK
      END         
   END
   
   AtIsAt = INSTRING('@', xSymbol,1)
   IF  AtIsAt <> 0
       xSymbol = xSymbol[ 1 : AtIsAt - 1 ]  !BUG: for some DLL`s, I found one compiled with CBuilder 1.0, that this would be bad for
   end   
   RETURN xSymbol
   
!========================================================================================
LongestSymbol PROCEDURE(ExportQ ExpQ2Do) !in ExportQ Level 2
QRec          LONG,AUTO
RetMaxSymLen  LONG
CurrSymLen    LONG,AUTO
symbol        LIKE(ExpQ2Do.symbol)!,auto
instringLoc   LONG,AUTO
   CODE   
   RetMaxSymLen = 0
   LOOP QRec = 1 TO RECORDS(ExpQ2Do)
      GET(ExpQ2Do, QRec)
      case ExpQ2Do.treelevel
        of 2; CurrSymLen = len(clip( CleanSymbol( ExpQ2Do.symbol)))
              if RetMaxSymLen < CurrSymLen
                 RetMaxSymLen = CurrSymLen
              end
              !MG: should cache this strip'd symbol in a queue, and process that intead vs. duplicating the effort below
      end
   end
   RETURN RetMaxSymLen
   
!========================================================================================
GenerateClasses   PROCEDURE
QRec           LONG,AUTO
GenDir         CSTRING (FILE:MaxFilePath+1)
CLW_Name       CSTRING (FILE:MaxFilePath+1)
INC_Name       CSTRING (FILE:MaxFilePath+1)
ModuleBaseName CSTRING(SIZE(ExportQ.module) + 1)
ClassName      LIKE(ModuleBaseName)
MaxSymLen      LONG,AUTO !Used For Alignment
   CODE  
   
   DO PromptForFolder !may return
   SETCURSOR(Cursor:Wait)

   MaxSymLen = LongestSymbol(WriteExpQ) + 1 !in WriteExpQ Level 2   
   IF MaxSymLen < 10
      MaxSymLen = 10
   END

   DO Generate:AllINC
   DO Generate:AllCLW
   
   SETCURSOR()
   IF MESSAGE('Would you like to open Explorer?||'&GenDir,'Classes Generated',ICON:Question,Button:Yes+Button:No)=Button:Yes
      RUN('Explorer ' & CLIP(GenDir))
   END
   
PromptForFolder ROUTINE !may return
   GenDir = 'C:\Tmp\GenDir' !<-- TODO Fix
   IF NOT FileDialog('Where should I generate the classes?', GenDir, 'All Directories|*.*',FILE:KeepDir + FILE:NoError + FILE:LongName + FILE:Directory)
      RETURN
   END
   IF GenDir[LEN(GenDir)]<>'\'
      GenDir =   GenDir  & '\'
   END


SetNames            ROUTINE
   ModuleBaseName = SUB(WriteExpQ.module, 1, INSTRING('.',WriteExpQ.module,1,1) - 1)
   CLW_Name       = 'ct' & ModuleBaseName & '.clw'
   INC_Name       = 'ct' & ModuleBaseName & '.inc'
   ClassName      = 'ct' & ModuleBaseName 

Ascii_Start             ROUTINE   
   CREATE(ASCIIfile)
   IF ERRORCODE()
      MESSAGE('Failed to Create File['& FileName &']')
      !Now what?
   END
   OPEN(ASCIIFile)

Ascii_Done              ROUTINE
      FLUSH(ASCIIFile)
      CLOSE(ASCIIfile)

!Region Generate:INC
Generate:AllINC    ROUTINE
   GET(WriteExpQ, 1)
   LOOP WHILE NOT ERRORCODE()
      DO SetNames
      FileName = GenDir & INC_Name
      DO Ascii_Start      
      DO Generate:OneINC_Fill   !<---
      DO Ascii_Done      
      GET(WriteExpQ,  POINTER(WriteExpQ) +1)
   END
   
Generate:OneINC_Methods   ROUTINE
   GET(WriteExpQ,  POINTER(WriteExpQ) +1 )
   LOOP WHILE NOT ERRORCODE()
      CASE ABS(WriteExpQ.treelevel)  !CB added ABS
        OF 1; BREAK
        OF 2; AppendAscii(CLIP(WriteExpQ.symbol) & ALL(' ',MaxSymLen - LEN(CLIP(WriteExpQ.symbol))) &'PROCEDURE(   )')
      END
      GET(WriteExpQ,  POINTER(WriteExpQ) +1)
   END   
   
Generate:OneINC_Fill ROUTINE   
   AppendAscii('!ABCIncludeFile(Yada)')
   AppendAscii('')
   AppendAscii('! This File Was Generated by LibMaker on '& FORMAT(TODAY(),@D1) &' at '& FORMAT(CLOCK(),@T1))
   AppendAscii('! Arguments and return values are unknown')
   AppendAscii('')
   AppendAscii('INCLUDE(<39>API.EQU<39>),ONCE')
   AppendAscii('')
   AppendAscii(ClassName &'  CLASS,TYPE,MODULE(<39>' & CLW_Name & '<39>),LINK(<39>' & CLW_Name & '<39>) !,_YadaLinkMode),DLL(_YadaDllMode)')
   AppendAscii('CONSTRUCT'& ALL(' ',MaxSymLen - 9) &'PROCEDURE()')
   AppendAscii('DESRTRUCT'& ALL(' ',MaxSymLen - 9) &'PROCEDURE()')
   DO Generate:OneINC_Methods
   AppendAscii(ALL(' ', LEN(ClassName) + 2) & 'END')
!EndRegion Generate:INC
   
Generate:AllCLW    ROUTINE
   GET(WriteExpQ, 1)
   LOOP WHILE NOT ERRORCODE()
      DO SetNames
      FileName = GenDir & CLW_Name
      DO Ascii_Start      
      DO Generate:OneCLW_Fill   !<---
      DO Ascii_Done      
      GET(WriteExpQ,  POINTER(WriteExpQ) +1)
   END

Generate:OneCLW_Fill ROUTINE   
   DATA
OrigPtr LONG,AUTO   
   CODE
   AppendAscii('    MEMBER()')
   AppendAscii('    MAP')
                           OrigPtr = POINTER(WriteExpQ)
     DO Generate:OneCLW_MAP
                           GET(WriteExpQ, OrigPtr)
   AppendAscii('    END')
   AppendAscii('')
   AppendAscii('    INCLUDE(<39>'& INC_Name  &'<39>),ONCE')
   AppendAscii('')
   
   WriteExpQ.symbol = 'CONSTRUCT'; DO Generate:OneCLW_Methods:One
   WriteExpQ.symbol = 'DESTRUCT' ; DO Generate:OneCLW_Methods:One
   
   DO Generate:OneCLW_Methods

Generate:OneCLW_MAP     ROUTINE
   DATA
szCleanedSymbol   CSTRING(SIZE(WriteExpQ.symbol) + 1)
   CODE
   GET(WriteExpQ,  POINTER(WriteExpQ) +1 )
   LOOP WHILE NOT ERRORCODE()
      CASE ABS(WriteExpQ.treelevel)  !CB added ABS
        OF 1; BREAK
        OF 2; szCleanedSymbol = CLIP(CleanSymbol(WriteExpQ.symbol))
              AppendAscii('<32>{8}'  |
                          & szCleanedSymbol |
                          & ALL('<32>',MaxSymLen - LEN(szCleanedSymbol)) |
                          & '(<9>{4}),pascal,raw,dll(1),name(<39>' & CLIP(WriteExpQ.symbol) & '<39>)')
      END
      GET(WriteExpQ,  POINTER(WriteExpQ) +1)
   END   

Generate:OneCLW_Methods ROUTINE
   GET(WriteExpQ,  POINTER(WriteExpQ) +1 )
   LOOP WHILE NOT ERRORCODE()
      CASE ABS(WriteExpQ.treelevel) !CB added  ABS
        OF 1; BREAK
        OF 2; DO Generate:OneCLW_Methods:One
      END
      GET(WriteExpQ, POINTER(WriteExpQ) +1)
   END   

Generate:OneCLW_Methods:One   ROUTINE
   DATA
szCleanedSymbol   CSTRING(SIZE(WriteExpQ.symbol) + 1)
   CODE
   !AppendAscii(CLIP(WriteExpQ.symbol) & ALL(' ',MaxSymLen - LEN(CLIP(WriteExpQ.symbol))) &'PROCEDURE(   )')
   szCleanedSymbol = CLIP(CleanSymbol(WriteExpQ.symbol))
   AppendAscii('')
   AppendAscii(ClassName &'.'& szCleanedSymbol    & ALL('<32>',MaxSymLen - LEN(szCleanedSymbol))   & 'PROCEDURE(  )')
   AppendAscii('   CODE')

AppendAscii    PROCEDURE(STRING xLine)  
   CODE
   ASCII:Line = xLine
   ADD(ASCIIFile)

!=======================================================================
SubtractFileLIB  PROCEDURE(<*string InFileName>)  
QNdx        LONG,AUTO
TagBtn      LONG,AUTO
FindQ       QUEUE(ExportQ),PRE(FindQ)
PtrExpQ         LONG
            END
SavExportQ  QUEUE(ExportQ),PRE(SavExpQ)         !A copy of the queue before subtracting
            END
SubtractQ   QUEUE(ExportQ),PRE(SubQ)            !The files to be subtracted
            END
CntExport      long
CntSubtracted  long
CntNotFound    long
!SubFileName    LIKE(FileName),STATIC
 
Swindow WINDOW('Subtract Symbols'),AT(,,240,179),GRAY,IMM,SYSTEM,FONT('Segoe UI',9,,FONT:regular),RESIZE
        ENTRY(@s255),AT(2,3,,11),FULL,USE(SubFileName),SKIP,TRN,FLAT,READONLY
        PROMPT('Subtracting...'),AT(4,17,231,19),USE(?SubSummary)
        BUTTON('&Delete'),AT(4,40,48,15),USE(?DeleteBtn),ICON('DelLine.ico'),TIP('Delete (subtract) ' & |
                'the symbols from the All List.'),LEFT
        BUTTON('&Tag'),AT(58,40,48,15),USE(?TagBtn),ICON('found.ico'),TIP('Tag the symbols in the Al' & |
                'l List.'),LEFT
        BUTTON('&UnTag'),AT(113,40,48,15),USE(?UnTagBtn),ICON('UnTag.ico'),TIP('UnTag the symbols in' & |
                ' the All List.'),LEFT
        BUTTON('&Cancel'),AT(189,40,48,15),USE(?CloseButton),STD(STD:Close)
        LIST,AT(4,60,232,115),USE(?List:FindQ),VSCROLL,FROM(FindQ),FORMAT('166L(2)|FMY~Duplicate Sym' & |
                'bols in Subtract LIB~@s128@22L(6)~Module~L(2)@s128@#6#'),ALRT(DeleteKey)
    END
MGResizeCls ResizeClassType
Pz LONG,DIM(4),STATIC    
    CODE
    IF ~SubFileName THEN 
        IF    EXISTS('C:\Clarion11') THEN SubFileName='C:\Clarion11\Lib\WIN32.LIB'
        ELSIF EXISTS('D:\Clarion11') THEN SubFileName='D:\Clarion11\Lib\WIN32.LIB'.
    END
    IF omitted(InFileName) !or AskForFile
        IF ~FileDialog('Subtract symbols from File ...', SubFileName, |
            'LIB files (*.lib)|*.lib|DLL files (*.dll)|*.dll|OCX files (*.ocx)|*.ocx|DLL, LIB and OCX files|*.dll;*.lib;*.ocx|All files (*.*)|*.*',FILE:LongName + FILE:KeepDir)
            RETURN
        END
        FileName = SubFileName
    ELSE
        FileName=InFileName
    END
    OPEN(Swindow)
    DISPLAY 
    SetCursor(CURSOR:Wait)    
    CopyQueue(ExportQ,SavExportQ)       !Save the original list so we cn use normal routines to load it
    FREE(ExportQ)
    GLO:IsSubtracting = 1
    IF UPPER(RIGHT(FileName,4))='.LIB' THEN  !01/02/22 was: IF INSTRING('.LIB', UPPER(FileName), 1, 1)
       ReadLib()
    ELSE
       ReadExecutable()
    END
    GLO:IsSubtracting = 0
    CopyQueue(ExportQ,SubtractQ)
    FREE(ExportQ)
    LOOP QNdx = 1 to RECORDS(SavExportQ)
         GET(SavExportQ,QNdx)
         ExportQ = SavExportQ
         ADD(ExportQ)         
         IF ExportQ.TreeLevel >= 2       !If Symbol and not module, see if we have Dup or Tagger
             CntExport += 1
             SubtractQ.SymbolLwr = ExportQ.SymbolLwr
             SubtractQ.TreeLevel = ExportQ.TreeLevel
             GET(SubtractQ,SubtractQ.SymbolLwr,SubtractQ.TreeLevel)     !Is the Symbol in the subtract Queue?
             if errorcode()
                CntNotFound += 1
             else
                CntSubtracted+=1                    !Yes we have this symbol
                FindQ = SavExportQ
                FindQ:PtrExpQ = QNdx
                ADD(FindQ)     
             end
         END !IF Symbol 
    END
    SetCursor()

  0{PROP:minheight}  = .7 * 0{PROP:height}
  0{PROP:minwidth }  = 0{PROP:width}
   MGResizeCls.Init_Class(Swindow)
   MGResizeCls.Add_ResizeQ(?List:FindQ,'L/T R/B')  
  ?List:FindQ{PROP:iconlist, 3} = '~Found.ico'
  ?SubSummary{PROP:Text}='Subtract LIB had ' & records(SubtractQ) & ' symbols of which ' & CntSubtracted & ' matched. ' &|
                         'Deleted from the ' & CntExport &' loaded symbols will leave ' & CntExport - CntSubtracted &'. ' & |
                         'Did not find ' &  CntNotFound &'.'  
  IF Pz[3]>9 THEN 
     SETPOSITION(0,Pz[1],Pz[2],Pz[3],Pz[4])
     POST(EVENT:Sized)
  END
  ACCEPT
    case EVENT()
    of EVENT:Sized ; MGResizeCls.Perform_Resize()
    end
    case ACCEPTED()
      of ?TagBtn orof ?UnTagBtn orof ?DeleteBtn
         TagBtn = ACCEPTED()
         DO Update_ExportQ
         BREAK
    end 
    case field()
    of ?List:FindQ 
       GET(FindQ,CHOICE(?List:FindQ))
       IF EVENT()=EVENT:AlertKey AND KEYCODE()=DeleteKey THEN 
          DELETE(FindQ)
       END
    end
  END !Accept Loop
  GETPOSITION(0,Pz[1],Pz[2],Pz[3],Pz[4])
  MGResizeCls.Close_Class
  CLOSE(SWindow)
  ExportQ_to_FilterQ() 
  RETURN
!----------------------
Update_ExportQ  ROUTINE  
    SORT(FindQ,-FindQ:PtrExpQ)      !Must do in reverse for delete so numbers don't chnage
    LOOP QNdx=1 TO RECORDS(FindQ)
         GET(FindQ,QNdx)
         GET(ExportQ,FindQ:PtrExpQ) ; IF ERRORCODE() THEN CYCLE.
         CASE TagBtn
          OF ?TagBtn     ; SetSearchFlagInExportQ(SearchFlag::Found,1,1)
          OF ?UnTagBtn   ; SetSearchFlagInExportQ(SearchFlag::Default,1,1) 
          OF ?DeleteBtn
             IF ExportQ.SearchFlag = SearchFlag::Found THEN Glo:FoundCount -= 1.
             DELETE(ExportQ)
         END    
    END
    EXIT


!-----------------------------------------------------
CopyQueue   procedure(*queue q1,*queue q2)
qnx     long,auto
    code
    loop qnx = 1 to records(q1)
         get(q1,qnx)
         q2=q1
         add(q2)
    end
    return
!-----------------------------------------------------
SkipAddExportQ  PROCEDURE()!,BOOL
SymName     PSTRING(130)
SkipIt      BOOL
UniqueSym   BOOL
SaveExportQ STRING(SIZE(ExportQ)),AUTO
SaveModule  LIKE(ExportQ.module),AUTO
!OmtIncAllDllX BYTE,STATIC       !CB 09/06/20 moved to Global with PRE(AddEx)
!TagIncAllDllX BYTE,STATIC
!OmtIncAllDups BYTE,STATIC
!TagIncAllDups BYTE,STATIC
    CODE 
    IF GLO:IsSubtracting THEN RETURN 0.
    SymName=CLIP(ExportQ.symbol)       
    CASE SymName
    OF   'DllGetVersion'            !CB I'm sure we do not want this in a LIB  
    OROF 'DllRegisterServer'        !Only called by RegSrv
    OROF 'DllUnregisterServer'      !Only called by RegSrv
    OROF 'DllInstall'               !Only called by RegSrv
    OROF 'DllCanUnloadNow'          !OLE calls not you
    OROF 'DllGetActivationFactory'  !COM calls
    OROF 'DllGetClassObject'        !OLE calls
          IF AddEx:OmtIncAllDllX THEN
             IF AddEx:TagIncAllDllX THEN SetSearchFlagInExportQ(SearchFlag::Found,1,0).  !1,0=SyncCnt,NoPut
             RETURN 2-AddEx:OmtIncAllDllX
          END
          SkipIt = 1 
          CASE MESSAGE('This Export Symbol is normally not statically linked so should not be included '&|
                '|in your LIB. It is exported by many DLLs and typically is dynamically loaded.'&|
                '|If you include it you likely will have duplicate errors from the linker.'&|
                '||     Name: '&SymName, |
            SymName, , 'Omit Symbol|Omit ALL|Include|Include ALL|Tag All')
          OF 1                          ! Omit Symbol
          OF 2 ; AddEx:OmtIncAllDllX=1  ! Omit ALL
          OF 3 ; SkipIt=0               ! Include
          OF 4 ; SkipIt=0 ; AddEx:OmtIncAllDllX=2                         ! Include ALL
          OF 5 ; SkipIt=0 ; AddEx:OmtIncAllDllX=2 ; AddEx:TagIncAllDllX=1 ! Tag ALL
          END !CASE    
         RETURN SkipIt
    END 
    SaveExportQ = ExportQ
    SaveModule  = ExportQ.Module
    ExportQ.symbol=SymName
    ExportQ.treelevel=2
    GET(ExportQ,ExportQ.symbol,ExportQ.treelevel) 
    UniqueSym=ERRORCODE()
    IF ~UniqueSym AND AddEx:OmtIncAllDups THEN
       SkipIt=2-AddEx:OmtIncAllDups
    ELSIF ~UniqueSym THEN
        CASE MESSAGE('A Symbol already exists. It can cause Linker errors to have a duplicate symbol name.'&|
             '||     Symbol: <9>'& CLIP(ExportQ.symbol) &|
              '|     Module: <9>'& CLIP(ExportQ.module) &|
             '||     Duplicates: <9>'& CLIP(SaveModule) &|
             '|','Duplicate Symbol: ' & SymName, , |
              'Omit Duplicate|Omit ALL|Include Dup|Include ALL|Tag All Dups')
        OF 1 ; SkipIt=1                             ! Omit Duplicate
        OF 2 ; SkipIt=1 ; AddEx:OmtIncAllDups = 1   ! Omit All
      ! OF 3                                        ! Include Duplicate        
        OF 4 ; AddEx:OmtIncAllDups=2                ! Include ALL         
        OF 5 ; AddEx:OmtIncAllDups=2 ; AddEx:TagIncAllDups=1 ! Tag ALL         
        END !CASE        
    END
    ExportQ = SaveExportQ
    IF ~UniqueSym AND AddEx:TagIncAllDups THEN SetSearchFlagInExportQ(SearchFlag::Found,1,0).  !1,0=SyncCnt,NoPut
    RETURN SkipIt
!=====================================================    
!Copied from LibMaker. Seems useful when using EXPORT on Declarion to get an EXP file to use instead
WriteTextEXP PROCEDURE
i           UNSIGNED,AUTO
  CODE
  IF ~FILEDIALOGA('Save EXP File', TxtFileName, |
                  'EXP File|*.EXP|Text File|*.TXT|All Files|*.*',|
                   FILE:Save+FILE:AddExtension+FILE:LongName) THEN RETURN.
   CLOSE(TxtFile)
   CREATE(TxtFile)
   OPEN(TxtFile)
   Txt:Line = 'LIBRARY ''Name'' GUI' ;  ADD(TxtFile)
   Txt:Line = 'EXPORTS' ;  ADD(TxtFile)
   LOOP i = 1 TO RECORDS(ExportQ)
      GET(ExportQ, i)
      CASE ABS(ExportQ.treelevel)
      OF 1
        Txt:Line = '; ' & ExportQ.module
        ADD(TxtFile)
      OF 2
         IF Glo.WriteFiltered AND glo.FoundCount THEN  !CB Save Filtered
            FLTQ:OrgOrder=ExportQ.OrgOrder
            GET(FilteredQ,FLTQ:OrgOrder)
            IF ERRORCODE() THEN CYCLE.
         END
         Txt:Line = '  ' & CLIP(ExportQ.Symbol) & ' @?'
         ADD(TxtFile)
      END
   END
   CLOSE(TxtFile)
   RETURN
!=============================================================
CopyList2Clip PROCEDURE()      !09/02/20 Copy one of Lists to Clip
TagNY BYTE,AUTO   !0=All 1=No 2=Yes
Fmt   BYTE,AUTO
i     LONG,AUTO
Tb    EQUATE('<9>') 
EOL   EQUATE('<13,10>') 
CB    ANY
SCnt  LONG 
LastModule  LIKE(ExportQ.module)
    CODE
    IF ?SHEET{PROP:ChoiceFEQ}=?TAB:Filtered AND glo.FoundCount THEN 
       TagNY=2
    ELSIF glo.FoundCount THEN 
        EXECUTE POPUPunder(?,'Copy Tagged|Copy UnTagged|Copy All')
        TagNY=2 
        TagNY=1
        TagNY=0
        ELSE ; RETURN 
        END 
    END              !  1                 2                     3                   4  
    Fmt=POPUPunder(?,'Symbol Name Only|Symbol -Tab- Module|Module -Tab- Symbol|Module -CR-LF- Symbols')
    IF ~Fmt THEN RETURN.
    
    LOOP i=1 TO RECORDS(ExportQ)
       GET(ExportQ, i)
       CASE ABS(ExportQ.treelevel)
       !OF 1 ExportQ.module
       OF 2
            EXECUTE TagNY
               IF ExportQ.SearchFlag = SearchFlag::Found THEN CYCLE.
               IF ExportQ.SearchFlag <> SearchFlag::Found THEN CYCLE.
            END 
            SCnt += 1
            IF Fmt=4 AND LastModule<>ExportQ.module THEN 
               LastModule = ExportQ.module
               CB=CB & LEFT('-{20} ' & CLIP(ExportQ.module) & ' -{20}'& EOL)
            END 
            CASE Fmt 
            OF 2 ; CB=CB & LEFT(CLIP(ExportQ.symbol) & TB & CLIP(ExportQ.module) & EOL)
            OF 3 ; CB=CB & LEFT(CLIP(ExportQ.module) & TB & CLIP(ExportQ.symbol) & EOL)  
            ELSE ; CB=CB & LEFT(CLIP(ExportQ.symbol) & EOL)
            END
       END
    END
    SETCLIPBOARD(CB)
    message(SCnt & ' symbols.','Copy List')
!=============================================================
!Only works with TEXT,SINGLE.  Best to put call in Event:OpenWindow but I think works after Open Window
!e.g. SetCueBanner(?Text1,'Cue Banner Text', 1/0)   ! 1=Show Cue with focus, 0=clear cue on focus 
!You MUST have a MANIFEST !You MUST have a MANIFEST !You MUST have a MANIFEST
!-------------------------------------------------------------
SetCueBanner   PROCEDURE(LONG FeqTextSL, STRING CueBannerText, BOOL OnFocusShow)
BStrCue     BSTRING 
Cue_WStr    LONG,OVER(BStrCue)      !BSTRING is Pointer to WSTR
    CODE
    BStrCue=CLIP(CueBannerText)     !BSTRING converts to UniCode.
    SendMessageW(FeqTextSL{PROP:Handle},1501h, OnFocusShow, Cue_WStr)      
    RETURN    !TB_SETCUEBANNER = EQUATE(1501h) !0x1501; //Textbox Integer   
!=============================================================
FileNamePopupClarion  PROCEDURE(*STRING OutFileName) !Right-Click on Add or Subtract
ClaDirQ QUEUE,pre(ClDrQ),static
Path        CSTRING(256) !ClDrQ:Path
        END
PUCla   CSTRING(1000)
P       SHORT,AUTO
    CODE
    IF ~RECORDS(ClaDirQ) THEN DO LoadClaDirQRtn.
    LOOP P = 1 to RECORDS(ClaDirQ)
         GET(ClaDirQ,P)
         IF P>4 AND P % 4=1 then PUCla=PUCla&'|-'.
         PUCla=choose(P=1,'',PUCla & '|') & ClDrQ:Path
    END
    P=POPUPunder(FIELD(),PUCla)
    IF P THEN
       Get(ClaDirQ,P)
       OutFileName = ClDrQ:Path
    END    
    RETURN CHOOSE(P)
LoadClaDirQRtn ROUTINE    
    data
VerNo   SHORT
Root    STRING(255)
    code
    VerNo=13
    LOOP
        VerNo -= 1 ; IF VerNo=9 THEN VerNo=91.
        Root=GetReg(REG_LOCAL_MACHINE,'SOFTWARE\SoftVelocity\Clarion' & VerNo,'root')
        IF Root AND EXISTS(Root)
           Root[1:4]=UPPER(Root[1:4])
           ClDrQ:Path = clip(Root) & '\BIN' ; ADD(ClaDirQ)
           ClDrQ:Path = clip(Root) & '\BIN\ClaRun.DLL' ; ADD(ClaDirQ)
           ClDrQ:Path = clip(Root) & '\LIB' ; ADD(ClaDirQ)
           ClDrQ:Path = clip(Root) & '\LIB\Win32.lib' ; ADD(ClaDirQ)
        end
    UNTIL VerNo=91
    IF RECORDS(ClaDirQ) THEN EXIT.
    IF POPUP('No Clarion 9.1 - 12'). 
    RETURN 0
!===================================
FileNamePopupSystem32  PROCEDURE(*STRING OutFileName) !Right-Click on Add or Subtract
PuAPI  STRING(| !'1234567890123456' &|  Must be 16 bytes to work                   '~System 32    |-' &|
                 'KERNEL32.DLL    ' &| !Top 3 first 
                 '|USER32.DLL     ' &|                 
                 '|GDI32.DLL    |-' &|
                 '|ADVAPI32.DLL   ' &|
                 '|COMCTL32.DLL   ' &|
                 '|COMDLG32.DLL   ' &|
                 '|CRYPT32.DLL    ' &|
                 '|MPR.DLL        ' &|
                 '|NETAPI32.dll   ' &|  !lower case .dll indicates not in Clarion's Win32.LIB
                 '|NTDLL.dll      ' &|
                 '|OLE32.DLL      ' &|
                 '|OLEAUT32.DLL   ' &|
                 '|OLEDLG.DLL     ' &|
                 '|PSAPI.dll      ' &|
                 '|SECUR32.dll    ' &|
                 '|SHELL32.DLL    ' &|
                 '|SHLWAPI.dll    ' &|
                 '|WINMM.DLL      ' &|
                 '|WININET.dll    ' &|
                 '|WINSPOOL.DRV   ' &|
                 '|WTSAPI32.dll   ' &|
                 '')                 
PuDLL  string(16)
P   SHORT,AUTO
    CODE
    P=POPUPunder(FIELD(),PuAPI)
    IF ~P THEN RETURN 0.
    PuDLL=sub(PuAPI,1+(P-1)*16,16)
    IF PuDLL[15:16]='|-' THEN PuDLL[15:16]=''.
    IF PuDLL[1]='|' THEN PuDLL=PuDLL[2:16].
    OutFileName=System32 & '\' & PuDLL 
    RETURN 1
!-----------------------
PopupUnder PROCEDURE(LONG UndFEQ, STRING PopMenu)!,LONG
X LONG,AUTO
Y LONG,AUTO
H LONG,AUTO
    CODE
    GETPOSITION(UndFEQ,X,Y,,H)
    IF UndFEQ{PROP:InToolBar} THEN Y -= (0{PROP:ToolBar}){PROP:Height} .
    RETURN POPUP(PopMenu,X,Y+H,1)    
!========================================================================================
ListInit PROCEDURE(SIGNED xFEQ)
   CODE    
   xFEQ{PROP:LineHeight}  = 1 + xFEQ{PROP:LineHeight} !CB was = 8  !AB 
   CASE xFEQ 
   OF ?Filtered ; xFEQ{proplist:Width,1} = xFEQ{prop:Width} - qOrdColWidth - xFEQ{proplist:Width,2}
   ELSE         ; xFEQ{proplist:Width,1} = xFEQ{prop:Width} - qOrdColWidth
   END
   xFEQ{PROP:iconlist, 1} = '~Opened.ico'
   xFEQ{PROP:iconlist, 2} = '~Closed.ico'
   xFEQ{PROP:iconlist, 3} = '~Found.ico'

  ! xFEQ{PROPStyle:TextColor   , ListStyle:None   } = COLOR:Black

   xFEQ{PROPStyle:TextColor   , ListStyle:Opened } = COLOR:Green
   xFEQ{PROPStyle:TextSelected, ListStyle:Opened } = COLOR:White
   xFEQ{PROPStyle:BackSelected, ListStyle:Opened } = xFEQ{PROPStyle:TextColor   , ListStyle:Opened }
   xFEQ{PROPSTYLE:FontSize    , ListStyle:Opened } = 10
   xFEQ{PROPSTYLE:FontStyle   , ListStyle:Opened } = FONT:Bold

   xFEQ{PROPStyle:TextColor   , ListStyle:Closed } = COLOR:Olive
   xFEQ{PROPStyle:TextSelected, ListStyle:Closed } = COLOR:White
   xFEQ{PROPStyle:BackSelected, ListStyle:Closed } = xFEQ{PROPStyle:TextColor   , ListStyle:Closed }
   xFEQ{PROPSTYLE:FontSize    , ListStyle:Closed } = 10
   xFEQ{PROPSTYLE:FontStyle   , ListStyle:Closed } = FONT:Bold

   xFEQ{PROPStyle:TextColor   , ListStyle:Found  } = COLOR:Blue
!---------------
RegExHelper PROCEDURE(STRING X, STRING Y, STRING W)
!Clarion Regular Expression Operators:<13,10><13,10>
!RxHlp STRING('^=Line Begins with next char<13,10>$=Line End' & |
!                's with prev char<13,10>  E.g. "A$" = line ends with "A"<13,10><13,10>?=Repeats 0 or 1 times<13,10>*=Re' & |
!                'peats 0 or more times<13,10>+=Repeats 1 or more.<13,10>  Prior char or [] or {{}<13,10><13,10>[]=Character S' & |
!                'et, [^]=Not Character Set, [-]=range<13,10>   E.g. CLA Label [?_A-Za-z][_:A-Za-z0-9]* <13>' & |
!                '<10>{{}=Grouping  |=Alternatives   E.g. {{Mr|Mrs|Ms}\.?  matches Mr Mr. Mrs Mrs. Ms' & |
!                ' Ms.<13,10><13,10>.=Any Character (Wildcard)    \=Escape Operators .?*+\[]{{}|^  Not used ' & |
!                'in [sets]  {900}')

RxHlp STRING('^=Line Begins with next char' &|
     '<13,10>$=Line Ends with prev char' &|
     '<13,10>  e.g. "W$" = line ends with "W"' &|
     '<13,10>' &|
     '<13,10>.=Any Character (Wildcard)    ' &|
     '<13,10>' &|
     '<13,10>?=Repeats 0 or 1 time' &|
     '<13,10>*=Repeats 0 or more times' &|
     '<13,10>+=Repeats 1 or more times' &|
     '<13,10>  Affects Prior Char, or [set], or {{group}' &|
     '<13,10>' &|
     '<13,10>[ ]=Character Set e.g. [0-9ABCDEF] Hex Digit' &|
     '<13,10>[^]=Negated   Set e.g. [^0-9ABCDEF] Not Hex' &|
     '<13,10>[-]=Range e.g. [A-Za-z0-9] Letters and Numbers' &|
     '<13,10>    e.g. CLA Label [?_A-Za-z][_:A-Za-z0-9]*' &|
     '<13,10>    For "-" to be in [set] it must be First' &|
     '<13,10>    For "^" to be in [set] it must NOT be First' &|
     '<13,10>' &|
     '<13,10>{{}=Grouping  ' &|
     '<13,10>| =Alternatives   ' &|
     '<13,10>   e.g. {{Mr|Mrs|Ms}\.?  matches Mr Mr. Mrs Mrs. Ms Ms.' &|
     '<13,10>' &|
     '<13,10>\=Escape Operators .?*+\[]{{}|^  Not used in [sets] '&|
     '<13,10><13,10>See Help on MATCH() for more. {900}' )
     
RxWindow  WINDOW('Clarion MATCH Regular Expression Operators'),AT(,,240,180),GRAY,SYSTEM,FONT('Segoe UI',10,,FONT:regular), |
            RESIZE
        TEXT,AT(2,2),USE(RxHlp),FULL,VSCROLL,FLAT,FONT('Consolas',10,,FONT:Bold)
    END
    CODE 
    OPEN(RxWindow)
    SETPOSITION(0,X+W+5,Y)
    ?RxHlp{PROP:Color}=80000018H   !COLOR:InfoBackground
    ?RxHlp{PROP:FontColor}=80000017H   !COLOR:InfoText
    ACCEPT ; END
    CLOSE(RxWindow)
    RETURN
!=========================================================================
ExistsFile PROCEDURE(STRING Nm)!BOOL !Is it a File and not Folder?
CN CSTRING(261),AUTO
FA LONG,AUTO
INVALID_FA EQUATE(-1)
FA_DIR     EQUATE(10h) 
    CODE
    CN=CLIP(Nm)
    FA=GetFileAttributes(CN)  
    RETURN CHOOSE(FA<>Invalid_FA AND ~BAND(FA,FA_DIR))
!=========================================================================    
ExistsFolder PROCEDURE(STRING Nm)!BOOL !Is it a Folder and not File?
CN CSTRING(261),AUTO
FA LONG,AUTO
INVALID_FA EQUATE(-1)
FA_DIR     EQUATE(10h) 
    CODE
    CN=CLIP(Nm)
    FA=GetFileAttributes(CN)
    RETURN CHOOSE(FA<>Invalid_FA AND BAND(FA,FA_DIR))
!=========================================================================
DB   PROCEDURE(STRING xMsg)
Prfx EQUATE('MGLibMkr: ')
sz   CSTRING(SIZE(Prfx)+SIZE(xMsg)+3),AUTO
  CODE 
  sz=Prfx & CLIP(xMsg) & '<13,10>'
  OutputDebugString(sz)
!=========================================================================
WinSxSPicker    PROCEDURE(*STRING OutFileName)!,BOOL
RetBool     BOOL
WinSxSDirBS PSTRING(256)
Dir1Q       QUEUE(FILE:Queue),PRE(Dir1Q).   !The current Dir (Dir1Loading) files

Pz       LONG,DIM(4),STATIC
FolderQ  QUEUE(FILE:Queue),PRE(FldQ),STATIC
NameLwrBS    PSTRING(256)
         END ! FldQ:Name  FldQ:ShortName(8.3?)  FldQ:Date  FldQ:Time  FldQ:Size  FldQ:SizeU  FldQ:Attrib  FldQ:NameLwr                                                                                                                                         FldQ:SizeU  ULONG,OVER(FldQ:Size) !Over 2GB will show negative w/o UNSigned SizeU
FindQ    QUEUE(),PRE(FndQ),STATIC
Name           STRING(255)
Date           LONG
Time           LONG
Size           LONG
Path           STRING(255)
NameLwr        STRING(255)
FldNameLwrBS   LIKE(FldQ:NameLwrBS)
         END ! FndQ:Name  FndQ:ShortName(8.3?)  FndQ:Date  FndQ:Time  FndQ:Size  FndQ:SizeU  FndQ:Attrib  FndQ:Path  FndQ:NameLwr
Dll2Find    STRING(32),STATIC
x86Only     BYTE(1),STATIC
LoadSys32   BYTE(1),STATIC
ChoiceFiles LONG,STATIC
ChoiceFolds LONG,STATIC

About STRING('The Windows WinSxS Folder holds multiple versions of DLLs that Windows loads based on the Manifest. The ' &|
     'files are each in a sub directory with a name that is about a foot long.' &|
     '<13,10>' &|
     '<13,10>WinSxS is where Version 6 of the Common Controls DLL (ComCtl32.DLL) is located. That DLL has all the ' &|
     'new features like Visual Styles. If you want to link to the newest procedures you need the V 6 DLL. E.g. the ' &|
     'new Task Dialog is in ComCtl32.DLL. The ComCtl32.DLL in System32 is the older Version 5.' &|
     '<13,10>' &|
     '<13,10>ComCtl32 V6 is found in the below folders. I would take the 6.18362.1 version. ' &|
     '<13,10>' &|
     '<13,10>C:\WINDOWS\WinSxS\' &|
     '<13,10>x86_microsoft.windows.common-controls_6595b64144ccf1df_6.0.18362.1_none_19851cfc38cbb817\' &|
     '<13,10>x86_microsoft.windows.common-controls_6595b64144ccf1df_6.0.18362.295_none_2e70e394278c3b98\' &|
     '<13,10>x86_microsoft.windows.common-controls_6595b64144ccf1df_6.0.18362.356_none_2e71e654278b50d2\  ' &|
     '<13,10>' &|
     '<13,10>FYI ... the LIB file is just for the Linker to have a list of Procedure Names with the correct case, ' &|
     'i.e. it just has names. <13,10>' )
     
SXSWindow WINDOW('WinSxS Folder Search (Windows Side-by-Side)'),AT(,,343,194),GRAY,SYSTEM, |
            FONT('Segoe UI',9,,FONT:regular),RESIZE
        PROMPT('&DLL File to Find:'),AT(2,4),USE(?Dll:Pmt)
        ENTRY(@s32),AT(55,3,121),USE(Dll2Find)
        BUTTON('&Find Files'),AT(185,2,42,13),USE(?FindBtn),DEFAULT
        BUTTON('&Cancel'),AT(236,2,42,13),USE(?CancelBtn),STD(STD:Close)
        CHECK('System&32'),AT(287,4),USE(LoadSys32),SKIP,TIP('Load files from System32 and WinSxS')
        SHEET,AT(2,20),FULL,USE(?SHEET1),NOSHEET,BELOW
            TAB('  File Search  '),USE(?TAB:Files)
                STRING('Double-Click on a File below to Select it and load into Lib Maker.  Ctrl+C t' & |
                        'o Copy Name + Folder.'),AT(6,37),USE(?Instrx)
                LIST,AT(1,50),FULL,USE(?ListFindQ),VSCROLL,FONT('Consolas'),FROM(FindQ), |
                        FORMAT('54L(2)|M~File Name~@s255@37R(2)|M~Date~C(0)@d1@37R(2)|M~Time~C(0)@T4' & |
                        '@42R(2)|M~Size~C(0)@n13@100L(2)~Folder~L(1)@s255@'),ALRT(CtrlC)
            END
            TAB(' SxS Folders  '),USE(?TAB:Folds)
                CHECK('x86 Only'),AT(6,37),USE(x86Only),SKIP,TIP('"x86" are the folders with 32-bit ' & |
                        'stuff.<13,10>Amd 64 are 64-bit DLL and always excluded.')
                ENTRY(@s255),AT(62,37,267,11),USE(WinSxSDirBS),SKIP,TRN,FLAT
                LIST,AT(1,50),FULL,USE(?ListFolders),VSCROLL,FONT('Consolas'),TIP('Double-click on l' & |
                        'ine to load contents of folder into file search list.<13,10>Ctrl+C to Copy Name.'), |
                        FROM(FolderQ),FORMAT('37R(2)|M~Date~C(0)@d1@#3#37R(2)|M~Time~C(0)@T4@120L(2)' & |
                        '~Name~L(1)@s255@#1#'),ALRT(CtrlC)
            END
            TAB(' About  '),USE(?TAB:About)
                TEXT,AT(1,37,,152),FULL,USE(About),FLAT,VSCROLL,READONLY
            END
        END
    END
DOO CLASS
LoadFolders     PROCEDURE
FindInFolders   PROCEDURE
AddDir1Q        PROCEDURE(STRING DirBS)
Load1Sxs        PROCEDURE
SetText         PROCEDURE
	END	
    CODE
    WinSxSDirBS=WindowsDir & '\WinSxS\'
    IF ~Dll2Find THEN Dll2Find='ComCtl32.DLL'.
    OPEN(SXSWindow)
    IF Pz[3]>9 THEN SETPOSITION(0,Pz[1],Pz[2],Pz[3],Pz[4]).
    ?SHEET1{PROP:TabSheetStyle}=1
    IF ~RECORDS(FolderQ) THEN 
        DOO.LoadFolders()
    ELSE
        ?ListFindQ{PROP:Selected}=ChoiceFiles
        ?ListFolders{PROP:Selected}=ChoiceFolds
    END
    DOO.SetText()
    ACCEPT
       CASE ACCEPTED() 
       OF ?FindBtn ; DOO.FindInFolders()
       OF ?x86Only ; DOO.LoadFolders()
       OF ?ListFindQ
            GET(FindQ,CHOICE(?ListFindQ))
            IF KEYCODE()=MouseLeft2 THEN
                FldQ:NameLwrBS = FndQ:FldNameLwrBS
                GET(FolderQ,FldQ:NameLwrBS)
                ?ListFolders{PROP:Selected}=POINTER(FolderQ) 
                RetBool = 1 
                OutFileName=CLIP(FindQ:Path) & FindQ:Name
                BREAK
            END
       OF ?ListFolders
            GET(FolderQ,CHOICE(?ListFolders))
            IF KEYCODE()=MouseLeft2 THEN DOO.Load1Sxs().
       END
       CASE FIELD()
       OF ?ListFindQ 
           GET(FindQ,CHOICE(?ListFindQ))
           IF KEYCODE()=CtrlC THEN SETCLIPBOARD(CLIP(FindQ.Name) &' '& Findq.Path).
       OF ?ListFolders 
           GET(FolderQ,CHOICE(?ListFindQ))
           IF KEYCODE()=CtrlC THEN SETCLIPBOARD(FolderQ.Name).
       END
    END !ACCEPT 
    ChoiceFiles = CHOICE(?ListFindQ)
    ChoiceFolds = CHOICE(?ListFolders)
    GETPOSITION(0,Pz[1],Pz[2],Pz[3],Pz[4])
    CLOSE(SXSWindow)
    RETURN RetBool
!--------------------    
DOO.SetText PROCEDURE
    CODE
    ?TAB:Folds{PROP:Text}=' SxS Folders (' & RECORDS(FolderQ) &') '
DOO.LoadFolders PROCEDURE
QNdx   LONG,AUTO
    CODE
    FREE(FolderQ) ; DISPLAY
    DIRECTORY(FolderQ,WinSxSDirBS & CHOOSE(~x86Only,'*.*','x86_*.*'),ff_:NORMAL+ff_:DIRECTORY)
    LOOP QNdx = RECORDS(FolderQ) TO 1 BY -1      !--Keep directories, Delete files
         GET(FolderQ,QNdx)
         IF ~BAND(FldQ:Attrib,FF_:Directory) OR FldQ:Name='.' OR FldQ:Name='..'   !. or .. Dirs
            DELETE(FolderQ) ; CYCLE
         END
         FldQ:NameLwrBS=LOWER(CLIP(FldQ:Name)) &'\'
         PUT(FolderQ)
         CASE FldQ:NameLwrBS[1:4]
         OF 'amd6'  !amd64 always 64-bit
            DELETE(FolderQ) ; CYCLE
         END   
    END !LOOP Delete files    
    SORT(FolderQ , FldQ:NameLwrBS )  !  FldQ:Name , FldQ:ShortName , FldQ:Date , FldQ:Time , FldQ:Size , FldQ:Attrib )
    DOO.SetText()
    RETURN    
DOO.FindInFolders PROCEDURE   !TODO Progress 
QNdx   LONG,AUTO
    CODE
    IF ~Dll2Find THEN 
        SELECT(?FindBtn) ; BEEP ; RETURN
    END
    FREE(FindQ) ; FREE(Dir1Q)
    LOOP QNdx = 1 TO RECORDS(FolderQ)
         GET(FolderQ,QNdx) 
         DIRECTORY(Dir1Q,WinSxSDirBS & FldQ:NameLwrBS & Dll2Find, ff_:NORMAL)
         IF RECORDS(Dir1Q) THEN 
            DOO.AddDir1Q(WinSxSDirBS & FldQ:NameLwrBS)
            FREE(Dir1Q)
            ?ListFolders{PROP:Selected}=QNdx
         END          
    END
    SORT(FindQ, FndQ:NameLwr, FndQ:FldNameLwrBS, FndQ:Date, FndQ:Time) 
    IF LoadSys32 THEN
       DIRECTORY(Dir1Q,System32 &'\'& Dll2Find, ff_:NORMAL)
       DOO.AddDir1Q(System32 &'\') 
       FREE(Dir1Q)
    END
    SELECT(?ListFindQ,1)
    RETURN
DOO.AddDir1Q PROCEDURE(STRING DirBS)
XD LONG,AUTO
    CODE
    LOOP XD=1 TO RECORDS(Dir1Q)
         GET(Dir1Q,XD)
         FindQ :=: Dir1Q
         FndQ:NameLwr = LOWER(FndQ:Name)
         FndQ:Path = DirBS
         FndQ:FldNameLwrBS = FldQ:NameLwrBS
         ADD(FindQ)
    END           
    RETURN
DOO.Load1Sxs PROCEDURE
    CODE
    FREE(FindQ) ; FREE(Dir1Q)
    DIRECTORY(Dir1Q,WinSxSDirBS & FldQ:NameLwrBS & '*.*', ff_:NORMAL)
    DOO.AddDir1Q(WinSxSDirBS & FldQ:NameLwrBS)
    SELECT(?ListFindQ,1)
