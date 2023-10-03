#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Fileversion=23.1003.929.19
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_Fileversion_First_Increment=y
#AutoIt3Wrapper_Res_Fileversion_Use_Template=%YY.%MO%DD.%HH%MI.%SE
#AutoIt3Wrapper_Res_Description=NAPS2 Wrapper
#AutoIt3Wrapper_Res_ProductName=NAPS2 Wrapper
#AutoIt3Wrapper_Icon=Res\Icon.ico
#AutoIt3Wrapper_Outfile_x64=..\Init.NAPS2.exe
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.16.1
 Author:         BiatuAutMiahn[@outlook.com]

#ce ----------------------------------------------------------------------------
#include <ProcessConstants.au3>
#include <WinAPIProc.au3>
#include <Security.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <GuiComboBox.au3>
#include <WinAPIProc.au3>
#include <Array.au3>
#include <File.au3>

Global Const $g_sDataDir=@ScriptDir
Global Const $g_sSessMagic=_RandStr()
Global Const $g_sLogPath=$g_sDataDir&'\'&@ScriptName&".log"
Global $bgNoLog=False, $g_iLogConsole = False
Global $g_oCOMError, $g_oCOMErrorDef, $g_iCOMError=0, $g_iCOMErrorExt=0, $g_sCOMError="", $g_sCOMErrorFunc="", $g_bCOMErrorLog=True
Global $g_oCOMErrorDef = ObjEvent("AutoIt.Error")
Global $g_oCOMError = ObjEvent("AutoIt.Error", "_COMErrorFunc")

Global Const $VERSION = "23.1003.929.19"
Global Const $sAlias="WrapNAPS"
Global $sTitle=$sAlias&" v"&$VERSION
Global $sBaseDir=@ScriptDir
If Not @Compiled Then $sBaseDir=@ScriptDir&"\.."
Global $bInstall,$bStay,$bRemember
Global $iScanner,$iScanSrc
Global $sRegKey="HKCU\Software\InfinitySys\Apps\WrapNAPS"
Global $sDest=@YEAR&'.'&@MON&'.'&@MDAY&','&@HOUR&@MIN&@SEC&'- Scan.pdf'
Global $sSaveDir="Documents\Scans"
Global $sOneDrivePath=@UserProfileDir&"\OneDrive_"; redacted CPNI
Global $sDestPath=$sOneDrivePath&'\'&$sSaveDir
Global $aScanners[1][3]
$aScanners[0][0]=0

Global $sNoOdWarn="The Lasting Change OneDrive folder cannot be found. OneDrive may not setup properly."&@LF
$sNoOdWarn&="Choose one of the following options below:"&@LF&@LF
$sNoOdWarn&="[Cancel]   : Abort the installation."&@LF
$sNoOdWarn&="[Try Again]: Check Again. (After logging into OneDrive)"&@LF
$sNoOdWarn&="[Continue] : Choose a destination folder for scans."&@LF&@LF
$sNoOdWarn&="NOTE: We are not responsible for data loss as a result of not storing data in a reliable location such as OneDrive."&@LF

Global $sMsgPostInstall='Success! A shortcut named "Scan to PDF" has been added to your desktop.'&@LF&@LF
$sMsgPostInstall&="NOTE: Scans will be stored in:"&@LF

Global $aErrMsgs[]=[ _
    "The selected scanner could not be found.", _
    "The selected scanner is offline.", _
    "No scanning device was found.", _
    "No pages are in the feeder.", _
    "The scanner's cover is open.", _
    "The scanner has a paper jam.", _
    "The selected scanner is busy.", _
    "The scanner's cover is open.", _
    "The scanner has a paper jam.", _
    "The scanner is warming up.", _
    "An error occurred with the scanning driver.", _
    "An error occurred when trying to save the file.", _
    "An unknown error occurred during the batch scan.", _
    "Batch scan stopped due to error.", _
    "The scanner is warming up.", _
    "No scanned pages to export.", _
    "The selected driver is not supported on this system.", _
    "The selected scanner does not support using duplex. If your scanner is supposed to support duplex, try using a different driver.", _
    "The selected scanner does not support using a feeder. If your scanner does have a feeder, try using a different driver." _
]


; # Installer ======================================================================================================================
If $CmdLine[0]<>0 Then
    $bInstall=False
    For $i=1 To $CmdLine[0]
        If $CmdLine[$i]<>"~!Install" Then ContinueLoop
        $bInstall=True
        ExitLoop
    Next
    If $bInstall Then
        $sTitle=$sAlias&" Installer v"&$VERSION
        Local $iRet
        Local $sNewPath
        Local $bAdmin=True;_IsUACAdmin()
        If Not $bAdmin Then
            $iRet=MsgBox(48+4,$sTitle,"Elevated permissions required for installation, would you like to attempt to elevate now?")
            If $iRet<>6 Then Exit 2
            Local $iPid
            If @Compiled Then
                $iPid=ShellExecute(@ScriptFullPath, "~!Install", @WorkingDir, "runas")
            Else
                $iPid=ShellExecute(@AutoItExe, '/AutoIt3ExecuteScript "' & @ScriptFullPath & '" ~!Install', @WorkingDir, "runas")
            EndIf
            $hProc=_WinAPI_OpenProcess($PROCESS_QUERY_INFORMATION, 0, $iPid)
            While ProcessExists($iPid)
                Sleep(125)
            WEnd
            Exit _WinAPI_GetExitCodeProcess($hProc)
        EndIf
        While Not _isDir($sOneDrivePath)
            $iRet=MsgBox(48+6,$sTitle,$sNoOdWarn)
            ConsoleWrite($iRet&@CRLF)
            Switch $iRet
                Case 2; Abort
                    Exit 2
                Case 11; Continue
                    While 1
                        $sNewPath=FileSelectFolder($sTitle&" - Choose Scan Folder",@UserProfileDir,7)
                        If $sNewPath<>'' Then ExitLoop
                        $iRet=MsgBox(48+4,$sTitle,"No directory was selected, do you want to abort the installation?")
                        If $iRet==6 Then Exit 2
                    WEnd
                    If $sNewPath<>'' Then ExitLoop
            EndSwitch
        WEnd
        If $sNewPath<>'' Then
            $sDestPath=$sNewPath
        EndIf
        $iRet=RegWrite($sRegKey,"ScanPath","REG_SZ",$sDestPath)
        If $iRet<>1  Then
            MsgBox(16,"","Failed to update registry! (Error: "&$iRet&')'&@LF&"Exiting.")
            Exit 1
        EndIf
        If FileCreateShortcut($sBaseDir&"\Init.NAPS2.exe",@DesktopDir&"\Scan to PDF.lnk",@ScriptDir)<>1 Then
            MsgBox(16,"","Failed to create desktop shortcut! Exiting.")
            Exit 1
        EndIf
        MsgBox(64,$sTitle,$sMsgPostInstall&'"'&$sDestPath&'"')
        Exit 0
    EndIf
EndIf
; ==================================================================================================================================

$sDestPath=RegRead($sRegKey,"ScanPath")
If @error Then
    MsgBox(16,$sTitle,"Please reinstall, ScanPath not configured.")
    Exit 1
EndIf

Local $iScale=1
Local $iGuiWidth=256+128
Local $iGuiComL=4
Local $iGuiComGrpL=8
Local $iGuiComBtnH=25
Local $iGuiComOptH=16
Local $iGuiComOptW=128
Local $iGuiScanOptT=52
Local $iGuiOptT=16+2
Local $iGuiBtnW=($iGuiWidth/3)-5


getScanners()
$hMain = GUICreate($sTitle, $iGuiWidth*$iScale, 192)
GUISetFont(10*$iScale, 400, 0, "Consolas")
$idScanner=GUICtrlCreateCombo("",$iGuiComL*$iScale,8*$iScale,($iGuiWidth-8)*$iScale,25,0x0003)
$hScanner=GUICtrlGetHandle($idScanner)
GUICtrlCreateGroup("Select a scan source", $iGuiComL*$iScale, 33*$iScale, ($iGuiWidth-8)*$iScale, 126*$iScale)
$idRadFlat = GUICtrlCreateRadio("Flatbed", ($iGuiComL+$iGuiComGrpL)*$iScale, $iGuiScanOptT*$iScale, $iGuiComOptW*$iScale, $iGuiComOptH*$iScale)
$idRadFeed = GUICtrlCreateRadio("Document Feeder", ($iGuiComL+$iGuiComGrpL)*$iScale, ($iGuiScanOptT+$iGuiComOptH)*$iScale, $iGuiComOptW*$iScale, $iGuiComOptH*$iScale)
$idRadDplx = GUICtrlCreateRadio("Duplex Tray", ($iGuiComL+$iGuiComGrpL)*$iScale, ($iGuiScanOptT+($iGuiComOptH*2))*$iScale, $iGuiComOptW*$iScale, $iGuiComOptH*$iScale)
GUICtrlCreateGroup("Options", $iGuiComL*$iScale, 101*$iScale, ($iGuiWidth-8)*$iScale, 58*$iScale)
$idChkRemb = GUICtrlCreateCheckbox("Remember my last choices", ($iGuiComL+$iGuiComGrpL)*$iScale, (101+$iGuiOptT)*$iScale, ($iGuiWidth-32)*$iScale, $iGuiComOptH*$iScale)
$idChkStay = GUICtrlCreateCheckbox("Stay Open (Batch Scan)", ($iGuiComL+$iGuiComGrpL)*$iScale, (101+($iGuiOptT*2))*$iScale, ($iGuiWidth-32)*$iScale, $iGuiComOptH*$iScale)
$idBtnScan = GUICtrlCreateButton("Scan", $iGuiComL, 163, $iGuiBtnW, $iGuiComBtnH)
$idBtnNAPS = GUICtrlCreateButton("NAPS2", $iGuiComL+$iGuiBtnW+4, 163, $iGuiBtnW, $iGuiComBtnH)
$idBtnExit = GUICtrlCreateButton("Exit", $iGuiComL+($iGuiBtnW*2)+8, 163, $iGuiBtnW, $iGuiComBtnH)
GUICtrlSetState($idScanner,$GUI_DISABLE)
GUICtrlSetState($idRadDplx,$GUI_DISABLE)
GUICtrlSetState($idRadFeed,$GUI_DISABLE)
GUICtrlSetState($idRadFlat,$GUI_DISABLE)
GUICtrlSetState($idBtnScan,$GUI_DISABLE)
If RegRead($sRegKey,"Remember") Then
    $bRemember=1
    GUICtrlSetState($idChkRemb,$GUI_CHECKED)
EndIf
_rescan()
GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")
GUISetState(@SW_SHOW,$hMain)

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE, $idBtnExit
			_Exit(0)
        Case $idBtnNAPS
            $aPos=MouseGetPos()
            ToolTip("Starting NAPS2...",$aPos[0]+10,$aPos[1]+10)
            $hTimeout=TimerInit()
            $iPid=Run($sBaseDir&"\App\NAPS2.exe",$sBaseDir,@SW_SHOW)
            While Sleep(10)
                $aPos=MouseGetPos()
                ToolTip("Starting NAPS2...",$aPos[0]+10,$aPos[1]+10)
                $aWnd=_WinAPI_EnumProcessWindows($iPid,1)
                If Not IsArray($aWnd) Then ContinueLoop
                If $aWnd[0][0] Then ExitLoop
            WEnd
            ToolTip("Starting NAPS2...Done",$aPos[0]+10,$aPos[1]+10)
            Sleep(250)
            ToolTip('')
            If Not $bStay Then _Exit(0)
        Case $idBtnScan
            $sDest=@YEAR&'.'&@MON&'.'&@MDAY&','&@HOUR&@MIN&@SEC&'- Scan.pdf'
            _GenProfile()
            $sProfile=''
            Switch $iScanSrc
                Case 0x1
                    $sProfile="Auto_Glass"
                Case 0x2
                    $sProfile="Auto_Feeder"
                Case 0x4
                    $sProfile="Auto_Duplex"
            EndSwitch
            $aPos=MouseGetPos()
            ;ConsoleWrite($sBaseDir&"\App\NAPS2.Console.exe --progress -p "&$sProfile&' --output "'&$sDestPath&'\'&$sDest&'"'&@CRLF)
            If Not _isDir($sDestPath) Then
                DirCreate($sDestPath)
            EndIf
            $sExecPath=$sBaseDir&"\App\NAPS2.Console.exe"
            If StringInStr($sExecPath,' ') Then $sExecPath='"'&$sExecPath&'"'
            $iPid=Run($sExecPath&" --progress -p "&$sProfile&' --output "'&$sDestPath&'\'&$sDest&'"',$sBaseDir,@SW_HIDE,0x10008)
            ToolTip("Scanning...",$aPos[0]+10,$aPos[1]+10)
            Local $sOutput=''
            Do
                If StdoutRead($iPid,1) Then
                    $sOutput&=StdoutRead($iPid)
                EndIf
                $aPos=MouseGetPos()
                ToolTip("Scanning...",$aPos[0]+10,$aPos[1]+10)
                Sleep(10)
            Until Not ProcessExists($iPid)
            If $sOutput<>'' Then
                $hLog=FileOpen($sBaseDir&"\LastScan.log",2)
                FileWrite($hLog,$sOutput)
                FileClose($hLog)
                ConsoleWrite('"""'&$sOutput&'"""'&@CRLF)
                For $iIdx=0 To UBound($aErrMsgs,1)-1
                    If StringInStr($sOutput,$aErrMsgs[$iIdx]) Then
                        ToolTip("Scanning...Failed",$aPos[0]+10,$aPos[1]+10)
                        Sleep(250)
                        ToolTip('')
                        MsgBox(16,$sTitle,$aErrMsgs[$iIdx])
                        ContinueLoop 2
                    EndIf
                Next
            EndIf
            ToolTip("Scanning...Done",$aPos[0]+10,$aPos[1]+10)
            Sleep(500)
            ToolTip('')
            If Not $bStay Then _Exit(0)
            ;-p Auto --output "'&@UserProfileDir&'\Documents\Scans\'&$sDest&'"  --progress',@ScriptDir)
	EndSwitch
WEnd
;~ If BitAND($iScanSrc,0x1) Then GUICtrlSetState($idRadFlat,$GUI_CHECKED)
;~             If BitAND($iScanSrc,0x2) Then GUICtrlSetState($idRadFeed,$GUI_CHECKED)
;~             If BitAND($iScanSrc,0x4) Then GUICtrlSetState($idRadDplx,$GUI_CHECKED)
; # Functions ======================================================================================================================

Func _Exit($iCode)
    _savePrefs()
    Exit $iCode
EndFunc

Func _savePrefs()
    If Not $bRemember Then
        RegWrite($sRegKey,"Remember","REG_DWORD",$bRemember)
        Return
    EndIf
    If $iScanner==-1 Then Return
    RegWrite($sRegKey,"LastScanner","REG_SZ",$aScanners[$iScanner][1])
    RegWrite($sRegKey,"Remember","REG_DWORD",$bRemember)
    RegWrite($sRegKey&"\DocSources",$aScanners[$iScanner][1],"REG_DWORD",$iScanSrc)
EndFunc

Func _rescan()
    _GUICtrlComboBox_BeginUpdate($hScanner)
    _GUICtrlComboBox_ResetContent($hScanner)
    If $aScanners[0][0]<>0 Then
        $iScanner=-1
        $sLastScanner=RegRead($sRegKey,"LastScanner")
        If @error Then $sLastScanner=-1
        For $i=1 To $aScanners[0][0]
            _GUICtrlComboBox_AddString($hScanner,$aScanners[$i][0])
            If $sLastScanner<>-1 Then
                If $sLastScanner=$aScanners[$i][1] Then $iScanner=$i
            EndIf
        Next
        _GUICtrlComboBox_SetCueBanner($hScanner,"Select Scanner")
        If $iScanner<>-1 Then
            _GUICtrlComboBox_SetCurSel($hScanner,$iScanner-1)
            GUICtrlSetState($idRadFlat,(BitAND($aScanners[$iScanner][2],0x1)?$GUI_ENABLE:$GUI_DISABLE))
            GUICtrlSetState($idRadFeed,(BitAND($aScanners[$iScanner][2],0x2)?$GUI_ENABLE:$GUI_DISABLE))
            GUICtrlSetState($idRadDplx,(BitAND($aScanners[$iScanner][2],0x4)?$GUI_ENABLE:$GUI_DISABLE))
            If $sLastScanner<>-1 Then
                ConsoleWrite($iScanner&@CRLF)
                $iScanSrc=RegRead($sRegKey&"\DocSources",$aScanners[$iScanner][1])
                If BitAND($iScanSrc,0x1) Then GUICtrlSetState($idRadFlat,$GUI_CHECKED)
                If BitAND($iScanSrc,0x2) Then GUICtrlSetState($idRadFeed,$GUI_CHECKED)
                If BitAND($iScanSrc,0x4) Then GUICtrlSetState($idRadDplx,$GUI_CHECKED)
            EndIf
            If $iScanSrc<>0 Then GUICtrlSetState($idBtnScan,$GUI_ENABLE)
        EndIf
        GUICtrlSetState($idScanner,$GUI_ENABLE)
    Else
        _GUICtrlComboBox_AddString($hScanner,"No Scanners Available")
        _GUICtrlComboBox_SetCurSel($hScanner,0)
        GUICtrlSetState($idScanner,$GUI_DISABLE)
        GUICtrlSetState($idBtnScan,$GUI_DISABLE)
        GUICtrlSetState($idRadDplx,$GUI_DISABLE)
        GUICtrlSetState($idRadFeed,$GUI_DISABLE)
        GUICtrlSetState($idRadFlat,$GUI_DISABLE)
    EndIf
    _GUICtrlComboBox_EndUpdate($hScanner)
EndFunc

Func _isDir($sPath)
    If Not FileExists($sPath) Then Return SetError(1,0,0)
    If Not StringInStr(FileGetAttrib($sPath),'d') Then Return SetError(1,1,0)
    Return SetError(0,0,1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _IsUACAdmin
; Description ...: Determines if process has Admin privileges and whether running under UAC.
; Syntax ........: _IsUACAdmin()
; Parameters ....: None
; Return values .: Success          - 1 - User has full Admin rights (Elevated Admin w/ UAC)
;                  Failure          - 0 - User is not an Admin, sets @extended:
;                                   | 0 - User cannot elevate
;                                   | 1 - User can elevate
; Author ........: Erik Pilsits
; Modified ......:
; Remarks .......: THE GOOD STUFF: returns 0 w/ @extended = 1 > UAC Protected Admin
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _IsUACAdmin()
    ; check elevation
    If StringRegExp(@OSVersion, "_(XP|20(0|3))") Or (Not _IsUACEnabled()) Then ; XP, XPe, 2000, 2003 > no UAC
        ; no UAC available or turned off
        If IsAdmin() Then
            Return SetExtended(0, 1)
        Else
            Return SetExtended(0, 0)
        EndIf
    Else
        ; check UAC elevation
        ;
        ; get process token groups information
        Local $hToken = _Security__OpenProcessToken(_WinAPI_GetCurrentProcess(), $TOKEN_QUERY)
        Local $tTI = _Security__GetTokenInformation($hToken, $TOKENGROUPS)
        _WinAPI_CloseHandle($hToken)
        ;
        Local $pTI = DllStructGetPtr($tTI)
        Local $cbSIDATTR = DllStructGetSize(DllStructCreate("ptr;dword"))
        Local $count = DllStructGetData(DllStructCreate("dword", $pTI), 1)
        Local $pGROUP1 = DllStructGetPtr(DllStructCreate("dword;STRUCT;ptr;dword;ENDSTRUCT", $pTI), 2)
        Local $tGROUP, $sGROUP = ""
        ;
        ; S-1-5-32-544 > BUILTINAdministrators > $SID_ADMINISTRATORS
        ; S-1-16-8192  > Mandatory LabelMedium Mandatory Level (Protected Admin) > $SID_MEDIUM_MANDATORY_LEVEL
        ; S-1-16-12288 > Mandatory LabelHigh Mandatory Level (Elevated Admin) > $SID_HIGH_MANDATORY_LEVEL
        ; SE_GROUP_USE_FOR_DENY_ONLY = 0x10
        ;
        ; check SIDs
        Local $inAdminGrp = False, $denyAdmin = False, $elevatedAdmin = False, $sSID
        For $i = 0 To $count - 1
            $tGROUP = DllStructCreate("ptr;dword", $pGROUP1 + ($cbSIDATTR * $i))
            $sSID = _Security__SidToStringSid(DllStructGetData($tGROUP, 1))
            If StringInStr($sSID, "S-1-5-32-544") Then
                ; member of Administrators group
                $inAdminGrp = True
                ; check for deny attribute
                If (BitAND(DllStructGetData($tGROUP, 2), 0x10) = 0x10) Then $denyAdmin = True
            ElseIf StringInStr($sSID, "S-1-16-12288") Then
                $elevatedAdmin = True
            EndIf
        Next
        ;
        If $inAdminGrp Then
            ; check elevated
            If $elevatedAdmin Then
                ; check deny status
                If $denyAdmin Then
                    ; protected Admin CANNOT elevate
                    Return SetExtended(0, 0)
                Else
                    ; elevated Admin
                    Return SetExtended(1, 1)
                EndIf
            Else
                ; protected Admin
                Return SetExtended(1, 0)
            EndIf
        Else
            ; not an Admin
            Return SetExtended(0, 0)
        EndIf
    EndIf
EndFunc   ;==>_IsUACAdmin

Func _IsUACEnabled()
    Return (RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "EnableLUA") = 1)
EndFunc   ;==>_IsUACEnabled

Func getScanners()
    Local $oWia=ObjCreate("WIA.DeviceManager")
    If Not IsObj($oWia) Then  Return SetError(1,0,0)
    Local $oDevs=$oWia.DeviceInfos
    If Not IsObj($oDevs) Then Return SetError(1,1,0)
    If $oDevs.Count=0 Then Return SetError(1,2,0)
    Local $iMax=0
    For $oDev In $oDevs
        If Not IsObj($oDev) Then ContinueLoop
        If $oDev.Type<>1 Then ContinueLoop
        $oScanner = $oDev.Connect()
        If $g_iCOMError<>0 Then
            _Log("Cannot connect to scanner.")
            ContinueLoop
        EndIf
        If Not IsObj($oScanner) Then ContinueLoop
        $iMax=UBound($aScanners)
        ReDim $aScanners[$iMax+1][3]
        $aScanners[$iMax][0]=$oDev.Properties("Name").Value
        $aScanners[$iMax][1]=$oDev.DeviceID
        $aScanners[$iMax][2]=$oScanner.Properties("Document Handling Capabilities").Value
    Next
    $aScanners[0][0]=$iMax
    Return SetError(0,0,1)
EndFunc

Func WM_COMMAND($hWnd, $iMsg, $wParam, $lParam)
    If $hWnd<>$hMain Then Return $GUI_RUNDEFMSG
    Local $iCode = BitShift($wParam, 16)
    Local $iId = BitAND($wParam, 0xFFFF)
    ;Local $hCtrl = $lParam
    Switch $iCode
        Case $CBN_SELCHANGE
            If $iId<>$idScanner Then Return $GUI_RUNDEFMSG
            $iScanner=_GUICtrlComboBox_GetCurSel($hScanner)+1
            GUICtrlSetState($idRadFlat,(BitAND($aScanners[$iScanner][2],0x1)?$GUI_ENABLE:$GUI_DISABLE))
            GUICtrlSetState($idRadFeed,(BitAND($aScanners[$iScanner][2],0x2)?$GUI_ENABLE:$GUI_DISABLE))
            GUICtrlSetState($idRadDplx,(BitAND($aScanners[$iScanner][2],0x4)?$GUI_ENABLE:$GUI_DISABLE))
            $iScanSrc=RegRead($sRegKey&"\DocSources",$aScanners[$iScanner][1])
            If BitAND($iScanSrc,0x1) Then GUICtrlSetState($idRadFlat,$GUI_CHECKED)
            If BitAND($iScanSrc,0x2) Then GUICtrlSetState($idRadFeed,$GUI_CHECKED)
            If BitAND($iScanSrc,0x4) Then GUICtrlSetState($idRadDplx,$GUI_CHECKED)
        Case $BN_CLICKED
            Switch $iId
                Case $idRadFlat
                    If BitAnd(GUICtrlRead($iId),$GUI_CHECKED) = $GUI_CHECKED Then $iScanSrc=0x1
                    ConsoleWrite("iScanSrc:"&$iScanSrc&@CRLF)
                    If $iScanner<>-1 Then GUICtrlSetState($idBtnScan,$GUI_ENABLE)
                Case $idRadFeed
                    If BitAnd(GUICtrlRead($iId),$GUI_CHECKED) = $GUI_CHECKED Then $iScanSrc=0x2
                    ConsoleWrite("iScanSrc:"&$iScanSrc&@CRLF)
                    If $iScanner<>-1 Then GUICtrlSetState($idBtnScan,$GUI_ENABLE)
                Case $idRadDplx
                    If BitAnd(GUICtrlRead($iId),$GUI_CHECKED) = $GUI_CHECKED Then $iScanSrc=0x4
                    ConsoleWrite("iScanSrc:"&$iScanSrc&@CRLF)
                    If $iScanner<>-1 Then GUICtrlSetState($idBtnScan,$GUI_ENABLE)
                Case $idChkRemb
                    $bRemember=BitAnd(GUICtrlRead($iId),$GUI_CHECKED) = $GUI_CHECKED
                    ConsoleWrite("Remb:"&$bRemember&@CRLF)
                Case $idChkStay
                    $bStay=BitAnd(GUICtrlRead($iId),$GUI_CHECKED) = $GUI_CHECKED
                    ConsoleWrite("Stay:"&$bStay&@CRLF)
            EndSwitch
    EndSwitch
    Return $GUI_RUNDEFMSG
EndFunc   ;==>MY_WM_COMMAND

Func _Base64Encode($sInput)
    $sInput=Binary($sInput)
    Local $tInput=DllStructCreate("byte["&BinaryLen($sInput)&']')
    DllStructSetData($tInput,1,$sInput)
    Local $tStruct=DllStructCreate("int")
    Local $aCall=DllCall("Crypt32.dll","int","CryptBinaryToString","ptr",DllStructGetPtr($tInput),"int",DllStructGetSize($tInput),"int",1,"ptr",0,"ptr",DllStructGetPtr($tStruct))
    If @error Or Not $aCall[0] Then Return SetError(1,0,0) ; error calculating the length of the buffer needed
    Local $tOut=DllStructCreate("char["&DllStructGetData($tStruct,1)&']')
    $aCall=DllCall("Crypt32.dll","int","CryptBinaryToString","ptr",DllStructGetPtr($tInput),"int",DllStructGetSize($tInput),"int",1,"ptr",DllStructGetPtr($tOut),"ptr",DllStructGetPtr($tStruct))
    If @error Or Not $aCall[0] Then Return SetError(2,0,0); error encoding
    Return DllStructGetData($tOut,1)
EndFunc   ;==>_Base64Encode

Func _Base64Decode($sInput)
    Local $tStruct=DllStructCreate("int")
    Local $aCall=DllCall("Crypt32.dll","int","CryptStringToBinary","str",$sInput,"int",0,"int",1,"ptr",0,"ptr",DllStructGetPtr($tStruct,1),"ptr",0,"ptr",0)
    If @error Or Not $aCall[0] Then Return SetError(1,0,0) ; error calculating the length of the buffer needed
    Local $tData=DllStructCreate("byte["&DllStructGetData($tStruct,1)&']')
    $aCall = DllCall("Crypt32.dll","int","CryptStringToBinary","str",$sInput,"int",0,"int",1,"ptr",DllStructGetPtr($tData),"ptr",DllStructGetPtr($tStruct,1),"ptr",0,"ptr",0)
    If @error Or Not $aCall[0] Then Return SetError(2,0,0); error decoding
    Return DllStructGetData($tData,1)
EndFunc   ;==>_Base64Decode

Func _GenProfile()
    If $iScanner=-1 Then Return
    $sProfiles=$sBaseDir&"\Data\Profiles.xml"
    Local $aProfiles
    _FileReadToArray($sProfiles,$aProfiles)
    ;_ArrayDIsplay($aProfiles)
    For $i=0 To UBound($aProfiles,1)-1
        If StringRegExp($aProfiles[$i],"<ID>[^\<]*</ID>") Then
            ConsoleWrite($aScanners[$iScanner][1]&@CRLF)
            $aProfiles[$i]="        <ID>"&$aScanners[$iScanner][1]&"</ID>"
        EndIf
        If StringRegExp($aProfiles[$i],"<Name>[^\<]*</Name>") Then
            $aProfiles[$i]="        <Name>"&$aScanners[$iScanner][0]&"</Name>"
        EndIf
    Next
    ;_ArrayDIsplay($aProfiles)
    _FileWriteFromArray($sProfiles,$aProfiles,1)
EndFunc

Func _COMErrorReset()
    $g_iCOMError=0
    $g_iCOMErrorExt=0
    $g_sCOMError=''
EndFunc

Func _COMErrorFunc()
    If Not IsObj($g_oCOMError) Then Return
    $g_iCOMError=1
    $g_iCOMErrorExt=$g_oCOMError.number
    $g_sCOMError=$g_oCOMError.windescription&" (0x"&Hex($g_iCOMErrorExt)&")"
    If $g_bCOMErrorLog Then _Log($g_sCOMError,$g_sCOMErrorFunc)
EndFunc

Func _Log($sStr,$sFunc="Main")
    If $bgNoLog Then Return
    Local $sStamp=@YEAR&'.'&@MON&'.'&@MDAY&','&@HOUR&':'&@MIN&':'&@SEC&':'&@MSEC
    Local $sErr="+["&$g_sSessMagic&'|'&$sStamp&"|"&@ComputerName&"|"&@UserName&"|"&@ScriptName&"|"&$sFunc&"]: "&$sStr
	If @Compiled And Not $g_iLogConsole Then
		If Not FileWriteLine($g_sLogPath,$sErr) Then
            MsgBox(16,"Error", "Cannot write to log"&@LF&$g_sLogPath&@LF&"Exiting.")
            Exit 0
        EndIf
    Else
		ConsoleWrite($sErr&@CRLF)
	EndIf
    Return
EndFunc

;
; Generate Random 16 digit Alphanumeric String
; UEZ, modified by Biatu
;
Func _RandStr()
    Local $sRet = "", $aTmp[3], $iLen = 16
    For $i = 1 To $iLen
        $aTmp[0] = Chr(Random(65, 90, 1)) ;A-Z
        $aTmp[1] = Chr(Random(97, 122, 1)) ;a-z
        $aTmp[2] = Chr(Random(48, 57, 1)) ;0-9
        $sRet &= $aTmp[Random(0, 2, 1)]
    Next
    Return $sRet
EndFunc
