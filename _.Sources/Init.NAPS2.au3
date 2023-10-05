#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Res\Icon.ico
#AutoIt3Wrapper_Outfile_x64=..\Init.NAPS2.exe
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Description=NAPS2 Wrapper
#AutoIt3Wrapper_Res_Fileversion=1.2310.513.1517
#AutoIt3Wrapper_Res_ProductName=NAPS2 Wrapper
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Run_After=echo %fileversion%>..\VERSION
#AutoIt3Wrapper_Run_Au3Stripper=y
#AutoIt3Wrapper_Res_Fileversion_First_Increment=y
#AutoIt3Wrapper_Res_FileVersion_AutoIncrement=y
#AutoIt3Wrapper_Res_Fileversion_Use_Template=1.%YY%MO.%DD%HH.%MI%SE
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
#include <Misc.au3>
#include <WinAPISys.au3>
#include <WinAPIFiles.au3>
#include <WinAPIProc.au3>
#include <WinAPIError.au3>

#include "Includes\WinHttp.au3"

Global Const $VERSION = "1.2310.513.1517"
Global Const $g_sSessMagic=_RandStr()
Global Const $sAlias="WrapNAPS"
Global $sBaseDir=@LocalAppDataDir&"\Programs\NAPS2"
If Not @Compiled Then $sBaseDir=@ScriptDir&"\.."
Global $sSnapsDir=$sBaseDir&"\Snaps"
Global $sLogPath=$sBaseDir&'\'&@ScriptName&".log"
Global $bgNoLog=False, $g_iLogConsole = False
Global $g_oCOMError, $g_oCOMErrorDef, $g_iCOMError=0, $g_iCOMErrorExt=0, $g_sCOMError="", $g_sCOMErrorFunc="", $g_bCOMErrorLog=True
Global $g_oCOMErrorDef = ObjEvent("AutoIt.Error")
Global $g_oCOMError = ObjEvent("AutoIt.Error", "_COMErrorFunc")

Global $sTitle=$sAlias&" v"&$VERSION

If _WinAPI_GetVersion()<10 Then
    MsgBox(16,$sTitle,"This program does support this version of windows!")
    Exit 1
EndIf

Global $bInstall,$bStay,$bRemember=1
Global $iScanner,$iScanSrc
Global $sRegKey="HKCU\Software\InfinitySys\Apps\WrapNAPS"
Global $sDest=@YEAR&'.'&@MON&'.'&@MDAY&','&@HOUR&@MIN&@SEC&'- Scan.pdf'
Global $sInitialSaveDir=@UserProfileDir&"\Documents"
;Global $sOneDrivePath=@UserProfileDir&"\OneDrive"; redacted CPNI
Global $sDestPath;=$sOneDrivePath&'\'&$sSaveDir
Global $aScanners[1][3]
$aScanners[0][0]=0

; Update Variables
Global $bCheckUpdate=True
Global $g_iProgPerc, $g_iProgBytes, $g_iProgTotal, $g_iUpdateStat
Global $hWINHTTP_STATUS_CALLBACK = DllCallbackRegister("__WINHTTP_STATUS_CALLBACK", "none", "handle;dword_ptr;dword;ptr;dword")

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
        Switch $CmdLine[$i]
            Case "~!Install"
                $bInstall=True
                ExitLoop
            Case "~!Update"
                _Update(1)
                Exit @Extended
            Case "~!Recovery"
                _UpdateRecovery()
                 Exit @Extended
        EndSwitch
    Next
    If $bInstall Then
        $sTitle=$sAlias&" Installer v"&$VERSION
        Local $iRet
        Local $sNewPath
        $sDestPath=FileSelectFolder($sTitle&" - Choose Scan Folder",$sInitialSaveDir,7)
        $iRet=RegWrite($sRegKey,"ScanPath","REG_SZ",$sDestPath)
        If $iRet<>1  Then
            MsgBox(16,"","Failed to update registry! (Error: "&$iRet&')'&@LF&"Exiting.")
            Exit 1
        EndIf
        If FileCreateShortcut($sBaseDir&"\Init.NAPS2.exe",@DesktopDir&"\Scan to PDF.lnk",$sBaseDir)<>1 Then
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
    $sDestPath=FileSelectFolder($sTitle&" - Choose Scan Folder",$sInitialSaveDir,7)
    $iRet=RegWrite($sRegKey,"ScanPath","REG_SZ",$sDestPath)
    If $iRet<>1  Then
        MsgBox(16,"","Failed to update registry! (Error: "&$iRet&')'&@LF&"Exiting.")
        Exit 1
    EndIf
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


$hMain = GUICreate($sTitle, $iGuiWidth*$iScale, 192)
GUISetFont(10*$iScale, 400, 0, "Consolas")
$idScanner=GUICtrlCreateCombo("",$iGuiComL*$iScale,8*$iScale,($iGuiWidth-8)*$iScale,25,0x0003)
$hScanner=GUICtrlGetHandle($idScanner)
GUICtrlCreateGroup("Document Source", $iGuiComL*$iScale, 33*$iScale, ($iGuiWidth-8)*$iScale, 126*$iScale)
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
GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")
GUISetState(@SW_SHOW,$hMain)
_rescan()
AdlibRegister("_DoUpdate")


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
            _UpdProfile()
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
                _Log("NAPS2.Console:"&@LF&$sOutput)
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
	EndSwitch
WEnd

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
    _GUICtrlComboBox_AddString($hScanner,"Detecting Scanners...")
    _GUICtrlComboBox_SetCurSel($hScanner,0)
    GUICtrlSetState($idScanner,$GUI_DISABLE)
    GUICtrlSetState($idBtnScan,$GUI_DISABLE)
    GUICtrlSetState($idRadDplx,$GUI_DISABLE)
    GUICtrlSetState($idRadFeed,$GUI_DISABLE)
    GUICtrlSetState($idRadFlat,$GUI_DISABLE)
    _GUICtrlComboBox_EndUpdate($hScanner)
    Sleep(500)
    getScanners()
    _udpScanners()
EndFunc

Func _udpScanners()
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
        If $aScanners[0][0]=1 Then $iScanner=1
        If $iScanner<>-1 Then
            _GUICtrlComboBox_SetCurSel($hScanner,$iScanner-1)
            GUICtrlSetState($idRadFlat,(BitAND($aScanners[$iScanner][2],0x1)?$GUI_ENABLE:$GUI_DISABLE))
            GUICtrlSetState($idRadFeed,(BitAND($aScanners[$iScanner][2],0x2)?$GUI_ENABLE:$GUI_DISABLE))
            GUICtrlSetState($idRadDplx,(BitAND($aScanners[$iScanner][2],0x4)?$GUI_ENABLE:$GUI_DISABLE))
            If $sLastScanner<>-1 Then
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

Func getScanners()
    Local $oWia=ObjCreate("WIA.DeviceManager")
    If Not IsObj($oWia) Then Return SetError(1,0,0)
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
                    ;ConsoleWrite("iScanSrc:"&$iScanSrc&@CRLF)
                    If $iScanner<>-1 Then GUICtrlSetState($idBtnScan,$GUI_ENABLE)
                Case $idRadFeed
                    If BitAnd(GUICtrlRead($iId),$GUI_CHECKED) = $GUI_CHECKED Then $iScanSrc=0x2
                    ;ConsoleWrite("iScanSrc:"&$iScanSrc&@CRLF)
                    If $iScanner<>-1 Then GUICtrlSetState($idBtnScan,$GUI_ENABLE)
                Case $idRadDplx
                    If BitAnd(GUICtrlRead($iId),$GUI_CHECKED) = $GUI_CHECKED Then $iScanSrc=0x4
                    ;ConsoleWrite("iScanSrc:"&$iScanSrc&@CRLF)
                    If $iScanner<>-1 Then GUICtrlSetState($idBtnScan,$GUI_ENABLE)
                Case $idChkRemb
                    $bRemember=BitAnd(GUICtrlRead($iId),$GUI_CHECKED) = $GUI_CHECKED
                    ;ConsoleWrite("Remb:"&$bRemember&@CRLF)
                Case $idChkStay
                    $bStay=BitAnd(GUICtrlRead($iId),$GUI_CHECKED) = $GUI_CHECKED
                    ;ConsoleWrite("Stay:"&$bStay&@CRLF)
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

;~ Func _GenProfile()
;~     If $iScanner=-1 Then Return
;~     $sProfiles=$sBaseDir&"\Data\Profiles.xml"
;~     Local $aProfiles
;~     _FileReadToArray($sProfiles,$aProfiles)
;~     ;_ArrayDIsplay($aProfiles)
;~     For $i=0 To UBound($aProfiles,1)-1
;~         If StringRegExp($aProfiles[$i],"<ID>[^\<]*</ID>") Then
;~             ConsoleWrite($aScanners[$iScanner][1]&@CRLF)
;~             $aProfiles[$i]="        <ID>"&$aScanners[$iScanner][1]&"</ID>"
;~         EndIf
;~         If StringRegExp($aProfiles[$i],"<Name>[^\<]*</Name>") Then
;~             $aProfiles[$i]="        <Name>"&$aScanners[$iScanner][0]&"</Name>"
;~         EndIf
;~     Next
;~     ;_ArrayDIsplay($aProfiles)
;~     _FileWriteFromArray($sProfiles,$aProfiles,1)
;~ EndFunc

Func _UpdProfile()
    Local $bWNPF
    Local $bWNPD
    Local $bWNPG
    Local $sNewProfile
    Local $sProfilePath="Profiles.xml"
    Local $bExists=FileExists($sProfilePath)
    Local $hFile
    ;FileSetPos($hFile,0)
    If $bExists Then
        $bWNPF=0
        $bWNPD=0
        $bWNPG=0
        $hFile=FileOpen($sProfilePath,0)
        If $hFile=-1 And $bExists Then
            _Log("Error: '"&$sProfilePath&"' exists, but cannot be opened for reading.")
            _MsgNoProfUpdate()
            Return SetError(1,1,0)
        EndIf
        _Log("Info: Reading NAPS2 Profiles.")
        Local $sData=FileRead($hFile)
        FileClose($hFile)
        Local $aProfiles=StringRegExp($sData,"<ScanProfile>((?:.|\s)*?)<\/ScanProfile>",3); Returns array of ScanProfile(s)
        If Not @error Then
            For $sProfile In $aProfiles
                $sDisplayName=_XmlGetField($sProfile,"DisplayName")
                ; This profile is nameless, skip.
                If @error Or $sDisplayName='' Then
                    _Log("Warning: Skipping the '"&$sDisplayName&"' profile, it doesn't have a name.")
                    ContinueLoop
                EndIf
                ; If this profile isn't intended for us, ignore it. We dont to modify profiles we don't need.
                If Not StringInStr($sAlias&"_Feeder"&$sAlias&"_Duplex"&$sAlias&"_Glass",$sDisplayName) Then ContinueLoop
                ; Replace the Device ID, and Device Name in the Profile.
                $sNewProfile=$sProfile
                $sId=_XmlGetField($sNewProfile,"ID")
                If Not @Error Then
                    $sNewProfile=StringReplace($sNewProfile,"<ID>"&$sId&"</ID>","<ID>"&$aScanners[$iScanner][1]&"</ID>")
                    If Not @extended Then
                        _Log("Error: Couldn't update NAPS2 profile's Device ID property. ('<ID>"&$sId&"</ID>' -> '<ID>"&$aScanners[$iScanner][1]&"</ID>')","_UpdProfile")
                        _MsgNoProfUpdate()
                        Return SetError(1,2,0)
                    EndIf
                Else
                    _Log("Error: NAPS2 profile does not contain a Device ID property.","_UpdProfile")
                    _MsgNoProfUpdate()
                    Return SetError(1,3,0)
                EndIf
                $sName=_XmlGetField($sNewProfile,"Name")
                If Not @Error Then
                    $sNewProfile=StringReplace($sNewProfile,"<Name>"&$sName&"</Name>","<Name>"&$aScanners[$iScanner][0]&"</Name>")
                    If Not @extended Then
                        _Log("Error: Couldn't update NAPS2 profile's Device Name property. ('<Name>"&$sName&"</Name>' -> '<Name>"&$aScanners[$iScanner][0]&"</Name>')","_UpdProfile")
                        _MsgNoProfUpdate()
                        Return SetError(1,4,0)
                    EndIf
                Else
                    _Log("Error: NAPS2 profile does not contain a Device Name property.","_UpdProfile")
                    _MsgNoProfUpdate()
                    Return SetError(1,5,0)
                EndIf
                ; Replace profile in xml.
                $sData=StringReplace($sData,$sProfile,$sNewProfile)
                If Not @extended Then
                    _Log("Error: Couldn't replace old profile string.","_UpdProfile")
                    _MsgNoProfUpdate()
                    Return SetError(1,6,0)
                EndIf
                Switch $sDisplayName
                    Case $sAlias&"_Feeder"
                        $bWNPF=1
                    Case $sAlias&"_Duplex"
                        $bWNPD=1
                    Case $sAlias&"_Glass"
                        $bWNPG=1
                EndSwitch
            Next
        Else
            _Log("Warning: No NAPS2 Profiles Exist.")
        EndIf
        If Not $bWNPF Or Not $bWNPD Or Not $bWNPG Then
            Local $sProfiles=''
            If Not $bWNPF Then
                _Log("Warning: The "&$sAlias&"_Feeder profile was not found, adding.")
                $sProfiles&=_GenProfile(0)
            EndIf
            If Not $bWNPD Then
                _Log("Warning: The "&$sAlias&"_Duplex profile was not found, adding.")
                $sProfiles&=_GenProfile(1)
            EndIf
            If Not $bWNPG Then
                _Log("Warning: The "&$sAlias&"_Glass profile was not found, adding.")
                $sProfiles&=_GenProfile(2)
            EndIf
            ;ConsoleWrite($sProfiles&@CRLF)
            $sData=StringReplace($sData,"</ArrayOfScanProfile>",$sProfiles&"</ArrayOfScanProfile>")
            If Not @extended Then
                _Log("Error: Couldn't append missing NAPS2 profiles.","_UpdProfile")
                _MsgNoProfUpdate()
                Return SetError(1,8,0)
            EndIf
        EndIf
        Local $bBackup=1
        If Not FileCopy($sProfilePath,$sProfilePath&".bak",1) Then
            $bBackup=0
            _Log("Warning: Cannot create a backup copy of '"&$sProfilePath&"'")
        EndIf
        $hFile=FileOpen($sProfilePath,2)
        If $hFile=-1 Then
            _Log("Error: '"&$sProfilePath&"' cannot be opened for writing.")
            If $bBackup Then
                If Not FileDelete($sProfilePath) Then
                    _Log("Warning: Cannot restore backup, cannot delete '"&$sProfilePath&"'")
                EndIf
                If Not FileCopy($sProfilePath&".bak",$sProfilePath,1) Then
                    _Log("Warning: Cannot restore backup, cannot overwrite '"&$sProfilePath&"'")
                EndIf
                If Not FileDelete($sProfilePath&".bak") Then
                    _Log("Warning: Cannot delete backup copy.")
                EndIf
            EndIf
            _MsgNoProfUpdate()
            Return SetError(1,9,0)
        EndIf
        If Not FileWrite($hFile,$sData) Then
            _Log("Error: Couldn't write NAPS2 Profiles.xml","_UpdProfile")
            Return SetError(1,10,0)
        EndIf
    Else
        ; Profiles.xml does not exist, creating default.
        $hFile=FileOpen($sProfilePath,2)
        If $hFile=-1 Then
            If $bExists Then
                _Log("Error: '"&$sProfilePath&"' exists, but cannot be opened for writing.")
            Else
                _Log("Error: '"&$sProfilePath&"' cannot be opened for writing.")
            EndIf
            _MsgNoProfUpdate()
            Return SetError(1,1,0)
        EndIf
        Local $sData='<?xml version="1.0" encoding="utf-8"?>'&@CRLF
        $sData&='<ArrayOfScanProfile>'&@CRLF
        For $i=0 To 2
            $sData&=_GenProfile($i)
        Next
        $sData&='</ArrayOfScanProfile>'&@CRLF
        FileWrite($hFile,$sData)
        FileClose($hFile)
    EndIf
    Return SetError(0,0,1)
EndFunc

Func _GenProfile($iType)
    Local $sData
    Local $aIntFields=StringSplit("Version|IconID|Brightness|Contrast|Quality|BlankPageWhiteThreshold|BlankPageCoverageThreshold|WiaDelayBetweenScansSeconds",'|')
    Local $aIntData=StringSplit("2|0|0|0|75|70|25|2",'|')
    Local $aBoolFields=StringSplit("MaxQuality|IsDefault|UseNativeUI|AfterScanScale|EnableAutoSave|AutoDeskew|BrightnessContrastAfterScan|ForcePageSize|ForcePageSizeCrop|ExcludeBlankPages|WiaOffsetWidth|WiaRetryOnFailure|WiaDelayBetweenScans|FlipDuplexedPages",'|')
    Local $aBoolData=StringSplit("0|0|0|0|0|0|0|0|0|0|0|0|2|0",'|')
    Local $aStrFields=StringSplit("DriverName|AfterScanScale|BitDepth|PageAlign|PageSize|Resolution|TwainImpl|WiaVersion",'|')
    Local $aStrData=StringSplit("wia|OneToOne|C24Bit|Right|Letter|Dpi300|Default|Default",'|')
    Local $aTmplFields=StringSplit("CustomPageSizeName|CustomPageSize|AutoSaveSettings|KeyValueOptions",'|')
    Local $aType[]=["Feeder","Duplex","Glass"]
    $sData="  <ScanProfile>"&@CRLF
    For $j=1 To $aIntFields[0]
        $sData&="    <"&$aIntFields[$j]&'>'&$aIntData[$j]&"</"&$aIntFields[$j]&'>'&@CRLF
    Next
    For $j=1 To $aBoolFields[0]
        $sData&="    <"&$aBoolFields[$j]&'>'&($aBoolData[$j]==1?"true":"false")&"</"&$aBoolFields[$j]&'>'&@CRLF
    Next
    For $j=1 To $aStrFields[0]
        $sData&="    <"&$aStrFields[$j]&'>'&$aStrData[$j]&"</"&$aStrFields[$j]&'>'&@CRLF
    Next
    For $j=1 To $aTmplFields[0]
        $sData&="    <"&$aTmplFields[$j]&' p3:nil="true" xmlns:p3="http://www.w3.org/2001/XMLSchema-instance" />'&@CRLF
    Next
    $sData&="    <PaperSource>"&$aType[$iType]&"</PaperSource>"&@CRLF
    $sData&="    <DisplayName>"&$sAlias&'_'&$aType[$iType]&"</DisplayName>"&@CRLF
    $sData&="    <Device>"&@CRLF
    $sData&="      <ID>"&$aScanners[$iScanner][1]&"</ID>"&@CRLF
    $sData&="      <Name>"&$aScanners[$iScanner][0]&"</Name>"&@CRLF
    $sData&="    </Device>"&@CRLF
    $sData&="  </ScanProfile>"&@CRLF
    Return $sData
EndFunc

Func _XmlGetField(ByRef $sXml,$sField)
    $sRet=StringRegExp($sXml,'<'&$sField&">([^<]*?)<\/"&$sField&'>',1)
    If @Error Then Return SetError(1,1,0)
    Return SetError(0,0,$sRet[0])
EndFunc

Func _MsgNoProfUpdate()
    MsgBox(16,$sTitle,"Error: Could not update NAPS2 profiles.xml. Please contact your system administrator or see log for details.")
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
    Local $sStamp=@YEAR&'.'&@MON&'.'&@MDAY&','&@HOUR&':'&@MIN&':'&@SEC&'.'&@MSEC
    Local $sErr="+["&$g_sSessMagic&'|'&$sStamp&"|"&@ComputerName&"|"&@UserName&"|"&@ScriptName&"|"&$sFunc&"]: "&$sStr
	If @Compiled And Not $g_iLogConsole Then
		If Not FileWriteLine($sLogPath,$sErr) Then
            MsgBox(16,"Error", "Cannot write to log"&@LF&$sLogPath&@LF&"Exiting.")
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

Func __Update_SecureGet($sSrv,$sUri,$fCallback=-1)
    If $fCallback<>-1 Then Call($fCallback,0,0,0)
    Local $hOpen=_WinHttpOpen()
    If @error Then Return SetError(1,1,0)
    _WinHttpSetStatusCallback($hOpen, $hWINHTTP_STATUS_CALLBACK)
    If $fCallback<>-1 Then Call($fCallback,1,0,0)
    Local $hConnect=_WinHttpConnect($hOpen,$sSrv,$INTERNET_DEFAULT_HTTPS_PORT)
    If @error Then
        _WinHttpCloseHandle($hOpen)
        Return SetError(1,2,0)
    EndIf
    Local $hRequest=_WinHttpOpenRequest($hConnect,"GET",$sUri,Default, Default, Default, BitOR($WINHTTP_FLAG_SECURE, $WINHTTP_FLAG_ESCAPE_DISABLE))
    If @error Then
        _WinHttpCloseHandle($hConnect)
        _WinHttpCloseHandle($hOpen)
        Return SetError(1,3,0)
    EndIf
	_WinHttpSetOption($hRequest, $WINHTTP_OPTION_DECOMPRESSION, $WINHTTP_DECOMPRESSION_FLAG_ALL)
	_WinHttpSetOption($hRequest, $WINHTTP_OPTION_UNSAFE_HEADER_PARSING,1)
    _WinHttpSendRequest($hRequest)
    If @error Then
        _WinHttpCloseHandle($hRequest)
        _WinHttpCloseHandle($hConnect)
        _WinHttpCloseHandle($hOpen)
        Return SetError(1,4,0)
    EndIf
    _WinHttpReceiveResponse($hRequest)
    If @error Then
        _WinHttpCloseHandle($hRequest)
        _WinHttpCloseHandle($hConnect)
        _WinHttpCloseHandle($hOpen)
        Return SetError(1,5,0)
    EndIf
    ; See if there is data to read
    Local $sChunk, $sData, $bAvail=_WinHttpQueryDataAvailable($hRequest)
    Local $iTotal=@Extended
    If $bAvail Then
        If $fCallback<>-1 Then Call($fCallback,2,0,$iTotal)
        ; Read
        While 1
            $sChunk = _WinHttpReadData($hRequest,2)
            If @error Then ExitLoop
            $sData &= BinaryToString($sChunk)
            If $fCallback<>-1 Then Call($fCallback,2,StringLen($sData),$iTotal)
        WEnd
    Else
        _WinHttpCloseHandle($hRequest)
        _WinHttpCloseHandle($hConnect)
        _WinHttpCloseHandle($hOpen)
        Return SetError(1,6,0)
    EndIf
    _WinHttpCloseHandle($hRequest)
    _WinHttpCloseHandle($hConnect)
    _WinHttpCloseHandle($hOpen)
    Return SetError(0,0,$sData)
EndFunc

Func __Update_CmdGetResult($sCmd,$sWorkDir=@ScriptDir,$iRetMode=0)
    Local $iPid,$hProc
    Local $aRet[3]
    $iPid=Run($sCmd,$sWorkDir,@SW_HIDE,6)
    $hProc=_WinAPI_OpenProcess($PROCESS_QUERY_INFORMATION,0,$iPid,0)
    If @error Then Return SetError(1,1,0)
    Local $vPeek,$vChunk,$bBreak=0
    Local $vStdOut
    Local $vStdErr
    Do
        $vPeek=StdoutRead($iPid,1,1)
        If BinaryLen($vPeek) Then
            $vChunk=StdoutRead($iPid,0,1)
            If @error Then $bBreak=1
            $vStdOut&=$vChunk
            $vChunk=StdErrRead($iPid,0,1)
            If @error Then $bBreak=1
            $vStdErr&=$vChunk
        EndIf
    Until Not ProcessExists($iPid) Or Not $bBreak
    $aRet[0]=_WinAPI_GetExitCodeProcess($hProc)
    If $iRetMode=0 Then
        $aRet[1]=$vStdOut
        $aRet[2]=$vStdErr
    ElseIf $iRetMode=1 Then
        Local $aStdOut,$aStdErr[]=[0],$iMax
        $sStdOut=StringReplace(BinaryToString($vStdOut),@CRLF,@LF)
        Do
            $sStdOut=StringReplace($sStdOut,@LF&@LF,@LF)
        Until Not StringInStr($sStdOut,@LF&@LF)
        $aRet[1]=StringSplit($sStdOut,@LF)
        $sStdErr=StringReplace(BinaryToString($vStdErr),@CRLF,@LF)
        Do
            $sStdErr=StringReplace($sStdErr,@LF&@LF,@LF)
        Until Not StringInStr($sStdErr,@LF&@LF)
        $aRet[1]=StringSplit($sStdOut,@LF)
        $aRet[2]=StringSplit($sStdErr,@LF)
    EndIf
    Return SetError(0,0,$aRet)
EndFunc

Func _Update($bPost=0)
    Local $sTitle=$sAlias&" Update"
    _Log("Current Version: "&$VERSION,"_Update")
    If $bPost Then
        _Log("################### Update Stage 2 ###################","_Update")
        ; I don't trust this, we should create a dir .\Snap\xx.xx.xx.xx.exe, then hardlink to .\Init.NAPS2.exe
        ; Get Version of main exec.
        $sBaseSnap=$sBaseDir&"\Init.NAPS2.exe"
        If Not FileExists($sBaseSnap) Then
            MsgBox(16,$sTitle,"Error: Update Failed, see log for details.")
            MsgBox(16,$sTitle,'Error: "Init.NAPS2.exe" could not be found. Update Failed.')
            Return SetError(0,11,1)
        EndIf
        Local $vCurVer=FileGetVersion($sBaseSnap)
        If @error Then
            MsgBox(16,$sTitle,"Error: Update Failed, see log for details.")
            _Log("Error: Could not retrieve FileVersion for Init.NAPS.exe","_Update")
            Return SetError(0,12,1)
        EndIf
        DirCreate($sSnapsDir)
        If Not _isDir($sSnapsDir) Then
            MsgBox(16,$sTitle,"Error: Update Failed, see log for details.")
            _Log("Error: Failed to create snaphot directory.","_Update")
            Return SetError(0,13,1)
        EndIf
        $sCurSnap=$sSnapsDir&'\WrapNAPS2_v'&$vCurVer&".exe"
        If Not FileExists($sCurSnap) Then
            FileCopy($sBaseSnap,$sCurSnap)
        EndIf
;~         Local $hToken = _WinAPI_OpenProcessToken(BitOR($TOKEN_ADJUST_PRIVILEGES, $TOKEN_QUERY))
;~         Local $aAdjust
;~         _WinAPI_AdjustTokenPrivileges($hToken, $SE_CREATE_SYMBOLIC_LINK_NAME, $SE_PRIVILEGE_ENABLED, $aAdjust)
;~         If @error Or @extended Then
;~             MsgBox(16,$sTitle,"Error: Update Failed, see log for details.")
;~             _Log("Error: Cannot grant SeCreateSymbolicLinkPrivilege.","_Update")
;~             Return SetError(0,14,1)
;~         EndIf
        FileDelete($sBaseSnap)
        If Not FileCopy($sSnapsDir&'\WrapNAPS2_v'&$VERSION&".exe",$sBaseSnap,1) Then
            MsgBox(16,$sTitle,"Error: Update Failed, see log for details.")
            _Log("Error: Cannot copy update","_Update")
             If FileCopy($sCurSnap,$sBaseSnap) Then _Log("Restored original Init.NAPS2.exe","_Update")
            _Log("Copy "&$sSnapsDir&'\WrapNAPS2_v'&$VERSION&".exe over Init.NAPS2.exe to manually update.","_Update")
            If Not FileExists($sBaseSnap) Then
                MsgBox(16,$sTitle,"Error: Cannot recover from update failure. Please contact your system administrator/developer.")
                Exit 1
            EndIf
        EndIf
        Run($sBaseSnap,$sBaseDir,@SW_SHOW)
;~         If Not _WinAPI_CreateSymbolicLink($sBaseSnap,$sSnapsDir&'\WrapNAPS2_v'&$VERSION&".exe") Then
;~             MsgBox(16,$sTitle,"Error: Update Failed, see log for details.")
;~             _Log("Error: Cannot create Symbolic Link","_Update")
;~             _Log("Restored original Init.NAPS2.exe","_Update")
;~              FileCopy($sCurSnap,$sBaseSnap)
;~             _Log("Copy "&$sSnapsDir&'\WrapNAPS2_v'&$VERSION&".exe over Init.NAPS2.exe to manually update.","_Update")
;~             Return SetError(0,15,1)
;~         EndIf
        Return SetError(0,2,1)
    EndIf
    If Not $bCheckUpdate Then
        _Log("Update check disabled","_Update")
        Return SetError(0,1,1)
    EndIf
    _Log("Checking for updates.","_Update")
    Local $iRet,$sRet
    $sRet=__Update_SecureGet("raw.githubusercontent.com","/BiatuAutMiahn/WrapNAPS/main/VERSION")
    If @error Then
        $iRet=@Extended
        Switch $iRet
            Case 1
                _Log("Error initializeing WinHttp","_Update")
            Case 2
                _Log("Error connecting to Update Server.","_Update")
            Case 3
                _Log("Error creating update request","_Update")
            Case 4
                _Log("Error sending update request","_Update")
            Case 5
                _Log("Error recieveing update request response","_Update")
            Case 6
                _Log("Error while recieving update data","_Update")
        EndSwitch
        Return SetError(1,$iRet,0)
    EndIf
    $sRet=StringStripWS(BinaryToString($sRet),7)
    _Log("Server Returned: "&$sRet,"_Update")
    If $sRet="404: Not Found" Then
        _Log("Error: Recieved HTTP Error 404 while checking for update","_Update")
        Return SetError(1,7,0)
    EndIf
    Local $vVer=_VersionCompare($VERSION,$sRet)
    If @error Then
        _Log("Error during update version comparison."&$sRet,"_Update")
        Return SetError(1,8,0)
    EndIf
    _Log("Upstream Version: "&$sRet,"_Update")
    If $vVer=0 Then
        _Log("Up to date!","_Update")
        Return SetError(0,3,1); No update available.
    EndIf
    If $vVer=1 Then
        _Log("Warning: upstream version is older than self.","_Update")
        Return SetError(1,9,0); We are newer than upstream
    EndIf
    If $vVer<>-1 Then
        _Log("Error, unexpected update variable! ("&$vVer&')',"_Update")
        Return SetError(1,10,0); _VersionCompare returned undocumented result.
    EndIf
    ; Prompt user for update.
    Local $iRet=MsgBox(32+4+65536,$sTitle,"An update is available!"&@LF&@LF&"Current version: "&$VERSION&@LF&"New version: "&$sRet&@LF&@LF&"Would you like to update now?",0,$hMain)
    If $iRet<>6 Then Return SetError(0,4,1)
    ; Download new version
    _Log("Downloading new version...","_Update")
    AdlibRegister("__UpdateProgWatch",10)
    $vUpdate=__Update_SecureGet("raw.githubusercontent.com","/BiatuAutMiahn/WrapNAPS/main/Init.NAPS2.exe","__UpdateProgCallback")
    If @error Then
        $iRet=@Extended
        Switch $iRet
            Case 1
                _Log("Error initializeing WinHttp","_Update")
            Case 2
                _Log("Error connecting to Update Server.","_Update")
            Case 3
                _Log("Error creating update request","_Update")
            Case 4
                _Log("Error sending update request","_Update")
            Case 5
                _Log("Error recieveing update request response","_Update")
            Case 6
                _Log("Error while recieving update data","_Update")
        EndSwitch
        AdlibUnregister("__UpdateProgWatch")
        MsgBox(16,$sTitle,"Error: Failed to download update. See log for details.")
        Return SetError(1,$iRet,0)
    EndIf    ; Execute new version with ~!Update
    AdlibUnregister("__UpdateProgWatch")
    DirCreate($sSnapsDir)
    If Not _isDir($sSnapsDir) Then
        MsgBox(16,$sTitle,"Error: Update Failed, see log for details.")
        _Log("Error: Failed to create snaphot directory.","_Update")
        Return SetError(0,13,1)
    EndIf
    ;If @error Then MsgBox(64,"Meh",@Error)
    $sUpdate=$sSnapsDir&'\WrapNAPS2_v'&$sRet&".exe"
    $hFile=FileOpen($sUpdate,2+8+16)
    If $hFile=-1 Then
        MsgBox(16,$sTitle,"Error: Update Failed, see log for details.")
        _Log('Error: Failed create file: "'&$sUpdate&'"',"_Update")
        Return SetError(0,14,1)
    EndIf
    If Not FileWrite($hFile,BinaryToString($vUpdate)) Then
        MsgBox(16,$sTitle,"Error: Update Failed, see log for details.")
        _Log('Error: Cannot write to file: "'&$sUpdate&'"',"_Update")
        Return SetError(0,14,1)
    EndIf
    FileClose($hFile)
    Run($sUpdate&" ~!Update",$sBaseDir,@SW_SHOW)
    Exit 0
EndFunc

Func __UpdateProgCallback($iStatus,$iBytes,$iTotal)
    $g_iUpdateStat=$iStatus
    $g_iProgBytes=$iBytes
    $g_iProgTotal=$iTotal
    $g_iProgPerc=Round(($iBytes/$g_iProgTotal)*100)
EndFunc

Func __UpdateProgWatch()
    Local $sMsg="Updating: "
    Switch $g_iUpdateStat
        Case 0
            $sMsg&="Initializing"
        Case 1
            $sMsg&="Connecting"
        Case 2
            $sMsg&="Downloading ("&$g_iProgPerc&"%,"&$g_iProgBytes&'\'&$g_iProgTotal&')'
    EndSwitch
    $aPos=MouseGetPos()
    ToolTip($sMsg,$aPos[0]+16,$aPos[1]+16,$sAlias&" Update",0,4)
    _Log($sMsg,"__UpdateProgWatch")
EndFunc

Func _UpdateRecovery()
    MsgBox(16,$sTitle,"Not Yet Implemented")
    Exit 1
EndFunc

Func _DoUpdate()
    AdlibUnRegister("_DoUpdate")
    _Update(0)
EndFunc

Func __WINHTTP_STATUS_CALLBACK($hInternet, $iContext, $iInternetStatus, $pStatusInformation, $iStatusInformationLength)
    #forceref $hInternet, $iContext, $pStatusInformation, $iStatusInformationLength
    ;ConsoleWrite("!->Current status of the connection: " & $iInternetStatus & " " & @TAB & " ")
    ; Interpret the status
    Local $sStatus
    Switch $iInternetStatus
        ;Case $WINHTTP_CALLBACK_STATUS_CLOSING_CONNECTION
        ;    $sStatus = "Closing the connection to the server"
        ;Case $WINHTTP_CALLBACK_STATUS_CONNECTED_TO_SERVER
        ;    $sStatus = "Successfully connected to the server."
        ;Case $WINHTTP_CALLBACK_STATUS_CONNECTING_TO_SERVER
        ;    $sStatus = "Connecting to the server."
        ;Case $WINHTTP_CALLBACK_STATUS_CONNECTION_CLOSED
        ;    $sStatus = "Successfully closed the connection to the server."
        ;Case $WINHTTP_CALLBACK_STATUS_DATA_AVAILABLE
        ;    $sStatus = "Data is available to be retrieved with WinHttpReadData."
        ;Case $WINHTTP_CALLBACK_STATUS_HANDLE_CREATED
        ;    $sStatus = "An HINTERNET handle has been created."
        ;Case $WINHTTP_CALLBACK_STATUS_HANDLE_CLOSING
        ;    $sStatus = "This handle value has been terminated."
        ;Case $WINHTTP_CALLBACK_STATUS_HEADERS_AVAILABLE
        ;    $sStatus = "The response header has been received and is available with WinHttpQueryHeaders."
        Case $WINHTTP_CALLBACK_STATUS_INTERMEDIATE_RESPONSE
            $sStatus = "Received an intermediate (100 level) status code message from the server."
        ;Case $WINHTTP_CALLBACK_STATUS_NAME_RESOLVED
        ;    $sStatus = "Successfully found the IP address of the server."
        ;Case $WINHTTP_CALLBACK_STATUS_READ_COMPLETE
        ;    $sStatus = "Data was successfully read from the server."
        ;Case $WINHTTP_CALLBACK_STATUS_RECEIVING_RESPONSE
        ;    $sStatus = "Waiting for the server to respond to a request."
        ;Case $WINHTTP_CALLBACK_STATUS_REDIRECT
        ;    $sStatus = "An HTTP request is about to automatically redirect the request."
        Case $WINHTTP_CALLBACK_STATUS_REQUEST_ERROR
            $sStatus = "An error occurred while sending an HTTP request."
        ;Case $WINHTTP_CALLBACK_STATUS_REQUEST_SENT
        ;    $sStatus = "Successfully sent the information request to the server."
        ;Case $WINHTTP_CALLBACK_STATUS_RESOLVING_NAME
        ;    $sStatus = "Looking up the IP address of a server name."
        ;Case $WINHTTP_CALLBACK_STATUS_RESPONSE_RECEIVED
        ;    $sStatus = "Successfully received a response from the server."
        Case $WINHTTP_CALLBACK_STATUS_SECURE_FAILURE
            $sStatus = "One or more errors were encountered while retrieving a Secure Sockets Layer (SSL) certificate from the server."
        ;Case $WINHTTP_CALLBACK_STATUS_SENDING_REQUEST
        ;    $sStatus = "Sending the information request to the server."
        ;Case $WINHTTP_CALLBACK_STATUS_SENDREQUEST_COMPLETE
        ;    $sStatus = "The request completed successfully."
        ;Case $WINHTTP_CALLBACK_STATUS_WRITE_COMPLETE
        ;    $sStatus = "Data was successfully written to the server."
    EndSwitch
    ; Print it
    If $sStatus<>'' Then _Log($sStatus,"__Update_SecureGet")
EndFunc    ;==>__WINHTTP_STATUS_CALLBACK
