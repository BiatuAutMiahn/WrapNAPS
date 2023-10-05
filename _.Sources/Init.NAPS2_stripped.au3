Global Const $UBOUND_ROWS = 1
Global Const $UBOUND_COLUMNS = 2
Global Const $NUMBER_DOUBLE = 3
Global $__g_vEnum, $__g_vExt = 0
Global Const $tagOSVERSIONINFO = 'struct;dword OSVersionInfoSize;dword MajorVersion;dword MinorVersion;dword BuildNumber;dword PlatformId;wchar CSDVersion[128];endstruct'
Func _WinAPI_GetVersion()
Local $tOSVI = DllStructCreate($tagOSVERSIONINFO)
DllStructSetData($tOSVI, 1, DllStructGetSize($tOSVI))
Local $aCall = DllCall('kernel32.dll', 'bool', 'GetVersionExW', 'struct*', $tOSVI)
If @error Or Not $aCall[0] Then Return SetError(@error, @extended, 0)
Return Number(DllStructGetData($tOSVI, 2) & "." & DllStructGetData($tOSVI, 3), $NUMBER_DOUBLE)
EndFunc
Func __EnumWindowsProc($hWnd, $bVisible)
Local $aCall
If $bVisible Then
$aCall = DllCall("user32.dll", "bool", "IsWindowVisible", "hwnd", $hWnd)
If Not $aCall[0] Then
Return 1
EndIf
EndIf
__Inc($__g_vEnum)
$__g_vEnum[$__g_vEnum[0][0]][0] = $hWnd
$aCall = DllCall("user32.dll", "int", "GetClassNameW", "hwnd", $hWnd, "wstr", "", "int", 4096)
$__g_vEnum[$__g_vEnum[0][0]][1] = $aCall[2]
Return 1
EndFunc
Func __Inc(ByRef $aData, $iIncrement = 100)
Select
Case UBound($aData, $UBOUND_COLUMNS)
If $iIncrement < 0 Then
ReDim $aData[$aData[0][0] + 1][UBound($aData, $UBOUND_COLUMNS)]
Else
$aData[0][0] += 1
If $aData[0][0] > UBound($aData) - 1 Then
ReDim $aData[$aData[0][0] + $iIncrement][UBound($aData, $UBOUND_COLUMNS)]
EndIf
EndIf
Case UBound($aData, $UBOUND_ROWS)
If $iIncrement < 0 Then
ReDim $aData[$aData[0] + 1]
Else
$aData[0] += 1
If $aData[0] > UBound($aData) - 1 Then
ReDim $aData[$aData[0] + $iIncrement]
EndIf
EndIf
Case Else
Return 0
EndSelect
Return 1
EndFunc
Func _WinAPI_EnumProcessThreads($iPID = 0)
If Not $iPID Then $iPID = @AutoItPID
Local $hSnapshot = DllCall('kernel32.dll', 'handle', 'CreateToolhelp32Snapshot', 'dword', 0x00000004, 'dword', 0)
If @error Or Not $hSnapshot[0] Then Return SetError(@error + 10, @extended, 0)
Local Const $tagTHREADENTRY32 = 'dword Size;dword Usage;dword ThreadID;dword OwnerProcessID;long BasePri;long DeltaPri;dword Flags'
Local $tTHREADENTRY32 = DllStructCreate($tagTHREADENTRY32)
Local $aRet[101] = [0]
$hSnapshot = $hSnapshot[0]
DllStructSetData($tTHREADENTRY32, 'Size', DllStructGetSize($tTHREADENTRY32))
Local $aCall = DllCall('kernel32.dll', 'bool', 'Thread32First', 'handle', $hSnapshot, 'struct*', $tTHREADENTRY32)
While Not @error And $aCall[0]
If DllStructGetData($tTHREADENTRY32, 'OwnerProcessID') = $iPID Then
__Inc($aRet)
$aRet[$aRet[0]] = DllStructGetData($tTHREADENTRY32, 'ThreadID')
EndIf
$aCall = DllCall('kernel32.dll', 'bool', 'Thread32Next', 'handle', $hSnapshot, 'struct*', $tTHREADENTRY32)
WEnd
DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hSnapshot)
If Not $aRet[0] Then Return SetError(1, 0, 0)
__Inc($aRet, -1)
Return $aRet
EndFunc
Func _WinAPI_EnumProcessWindows($iPID = 0, $bVisible = True)
Local $aThreads = _WinAPI_EnumProcessThreads($iPID)
If @error Then Return SetError(@error, @extended, 0)
Local $hEnumProc = DllCallbackRegister('__EnumWindowsProc', 'bool', 'hwnd;lparam')
Dim $__g_vEnum[101][2] = [[0]]
For $i = 1 To $aThreads[0]
DllCall('user32.dll', 'bool', 'EnumThreadWindows', 'dword', $aThreads[$i], 'ptr', DllCallbackGetPtr($hEnumProc), 'lparam', $bVisible)
If @error Then
ExitLoop
EndIf
Next
DllCallbackFree($hEnumProc)
If Not $__g_vEnum[0][0] Then Return SetError(11, 0, 0)
__Inc($__g_vEnum, -1)
Return $__g_vEnum
EndFunc
Global Const $BN_CLICKED = 0
Global Const $GUI_EVENT_CLOSE = -3
Global Const $GUI_RUNDEFMSG = 'GUI_RUNDEFMSG'
Global Const $GUI_CHECKED = 1
Global Const $GUI_ENABLE = 64
Global Const $GUI_DISABLE = 128
Global Const $WM_COMMAND = 0x0111
Global Const $CBM_FIRST = 0x1700
Global Const $CB_ADDSTRING = 0x143
Global Const $CB_GETCURSEL = 0x147
Global Const $CB_RESETCONTENT = 0x14B
Global Const $CB_SETCUEBANNER =($CBM_FIRST + 3)
Global Const $CB_SETCURSEL = 0x14E
Global Const $CBN_SELCHANGE = 1
Func _SendMessage($hWnd, $iMsg, $wParam = 0, $lParam = 0, $iReturn = 0, $wParamType = "wparam", $lParamType = "lparam", $sReturnType = "lresult")
Local $aCall = DllCall("user32.dll", $sReturnType, "SendMessageW", "hwnd", $hWnd, "uint", $iMsg, $wParamType, $wParam, $lParamType, $lParam)
If @error Then Return SetError(@error, @extended, "")
If $iReturn >= 0 And $iReturn <= 4 Then Return $aCall[$iReturn]
Return $aCall
EndFunc
Func _WinAPI_MultiByteToWideChar($vText, $iCodePage = 0, $iFlags = 0, $bRetString = False)
Local $sTextType = ""
If IsString($vText) Then $sTextType = "str"
If(IsDllStruct($vText) Or IsPtr($vText)) Then $sTextType = "struct*"
If $sTextType = "" Then Return SetError(1, 0, 0)
Local $aCall = DllCall("kernel32.dll", "int", "MultiByteToWideChar", "uint", $iCodePage, "dword", $iFlags, $sTextType, $vText, "int", -1, "ptr", 0, "int", 0)
If @error Or Not $aCall[0] Then Return SetError(@error + 10, @extended, 0)
Local $iOut = $aCall[0]
Local $tOut = DllStructCreate("wchar[" & $iOut & "]")
$aCall = DllCall("kernel32.dll", "int", "MultiByteToWideChar", "uint", $iCodePage, "dword", $iFlags, $sTextType, $vText, "int", -1, "struct*", $tOut, "int", $iOut)
If @error Or Not $aCall[0] Then Return SetError(@error + 20, @extended, 0)
If $bRetString Then Return DllStructGetData($tOut, 1)
Return $tOut
EndFunc
Global Const $__COMBOBOXCONSTANT_WM_SETREDRAW = 0x000B
Func _GUICtrlComboBox_AddString($hWnd, $sText)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Return _SendMessage($hWnd, $CB_ADDSTRING, 0, $sText, 0, "wparam", "wstr")
EndFunc
Func _GUICtrlComboBox_BeginUpdate($hWnd)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Return _SendMessage($hWnd, $__COMBOBOXCONSTANT_WM_SETREDRAW, False) = 0
EndFunc
Func _GUICtrlComboBox_EndUpdate($hWnd)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Return _SendMessage($hWnd, $__COMBOBOXCONSTANT_WM_SETREDRAW, True) = 0
EndFunc
Func _GUICtrlComboBox_GetCurSel($hWnd)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Return _SendMessage($hWnd, $CB_GETCURSEL)
EndFunc
Func _GUICtrlComboBox_ResetContent($hWnd)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
_SendMessage($hWnd, $CB_RESETCONTENT)
EndFunc
Func _GUICtrlComboBox_SetCueBanner($hWnd, $sText)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Local $tText = _WinAPI_MultiByteToWideChar($sText)
Return _SendMessage($hWnd, $CB_SETCUEBANNER, 0, $tText, 0, "wparam", "struct*") = 1
EndFunc
Func _GUICtrlComboBox_SetCurSel($hWnd, $iIndex = -1)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Return _SendMessage($hWnd, $CB_SETCURSEL, $iIndex)
EndFunc
Func _VersionCompare($sVersion1, $sVersion2)
If $sVersion1 = $sVersion2 Then Return 0
Local $sSubVersion1 = "", $sSubVersion2 = ""
If StringIsAlpha(StringRight($sVersion1, 1)) Then
$sSubVersion1 = StringRight($sVersion1, 1)
$sVersion1 = StringTrimRight($sVersion1, 1)
EndIf
If StringIsAlpha(StringRight($sVersion2, 1)) Then
$sSubVersion2 = StringRight($sVersion2, 1)
$sVersion2 = StringTrimRight($sVersion2, 1)
EndIf
Local $aVersion1 = StringSplit($sVersion1, ".,"), $aVersion2 = StringSplit($sVersion2, ".,")
Local $iPartDifference =($aVersion1[0] - $aVersion2[0])
If $iPartDifference < 0 Then
ReDim $aVersion1[UBound($aVersion2)]
$aVersion1[0] = UBound($aVersion1) - 1
For $i =(UBound($aVersion1) - Abs($iPartDifference)) To $aVersion1[0]
$aVersion1[$i] = "0"
Next
ElseIf $iPartDifference > 0 Then
ReDim $aVersion2[UBound($aVersion1)]
$aVersion2[0] = UBound($aVersion2) - 1
For $i =(UBound($aVersion2) - Abs($iPartDifference)) To $aVersion2[0]
$aVersion2[$i] = "0"
Next
EndIf
For $i = 1 To $aVersion1[0]
If StringIsDigit($aVersion1[$i]) And StringIsDigit($aVersion2[$i]) Then
If Number($aVersion1[$i]) > Number($aVersion2[$i]) Then
Return SetExtended(2, 1)
ElseIf Number($aVersion1[$i]) < Number($aVersion2[$i]) Then
Return SetExtended(2, -1)
ElseIf $i = $aVersion1[0] Then
If $sSubVersion1 > $sSubVersion2 Then
Return SetExtended(3, 1)
ElseIf $sSubVersion1 < $sSubVersion2 Then
Return SetExtended(3, -1)
EndIf
EndIf
Else
If $aVersion1[$i] > $aVersion2[$i] Then
Return SetExtended(1, 1)
ElseIf $aVersion1[$i] < $aVersion2[$i] Then
Return SetExtended(1, -1)
EndIf
EndIf
Next
Return SetExtended(Abs($iPartDifference), 0)
EndFunc
Global Const $INTERNET_DEFAULT_PORT = 0
Global Const $INTERNET_DEFAULT_HTTPS_PORT = 443
Global Const $ICU_ESCAPE = 0x80000000
Global Const $WINHTTP_FLAG_ASYNC = 0x10000000
Global Const $WINHTTP_FLAG_ESCAPE_DISABLE = 0x00000040
Global Const $WINHTTP_FLAG_SECURE = 0x00800000
Global Const $WINHTTP_ACCESS_TYPE_NO_PROXY = 1
Global Const $WINHTTP_NO_PROXY_NAME = ""
Global Const $WINHTTP_NO_PROXY_BYPASS = ""
Global Const $WINHTTP_NO_REFERER = ""
Global Const $WINHTTP_DEFAULT_ACCEPT_TYPES = 0
Global Const $WINHTTP_NO_ADDITIONAL_HEADERS = ""
Global Const $WINHTTP_NO_REQUEST_DATA = ""
Global Const $WINHTTP_OPTION_CALLBACK = 1
Global Const $WINHTTP_OPTION_RESOLVE_TIMEOUT = 2
Global Const $WINHTTP_OPTION_CONNECT_TIMEOUT = 3
Global Const $WINHTTP_OPTION_CONNECT_RETRIES = 4
Global Const $WINHTTP_OPTION_SEND_TIMEOUT = 5
Global Const $WINHTTP_OPTION_RECEIVE_TIMEOUT = 6
Global Const $WINHTTP_OPTION_RECEIVE_RESPONSE_TIMEOUT = 7
Global Const $WINHTTP_OPTION_HANDLE_TYPE = 9
Global Const $WINHTTP_OPTION_READ_BUFFER_SIZE = 12
Global Const $WINHTTP_OPTION_WRITE_BUFFER_SIZE = 13
Global Const $WINHTTP_OPTION_PARENT_HANDLE = 21
Global Const $WINHTTP_OPTION_EXTENDED_ERROR = 24
Global Const $WINHTTP_OPTION_SECURITY_FLAGS = 31
Global Const $WINHTTP_OPTION_URL = 34
Global Const $WINHTTP_OPTION_SECURITY_KEY_BITNESS = 36
Global Const $WINHTTP_OPTION_PROXY = 38
Global Const $WINHTTP_OPTION_USER_AGENT = 41
Global Const $WINHTTP_OPTION_CONTEXT_VALUE = 45
Global Const $WINHTTP_OPTION_CLIENT_CERT_CONTEXT = 47
Global Const $WINHTTP_OPTION_REQUEST_PRIORITY = 58
Global Const $WINHTTP_OPTION_HTTP_VERSION = 59
Global Const $WINHTTP_OPTION_DISABLE_FEATURE = 63
Global Const $WINHTTP_OPTION_CODEPAGE = 68
Global Const $WINHTTP_OPTION_MAX_CONNS_PER_SERVER = 73
Global Const $WINHTTP_OPTION_MAX_CONNS_PER_1_0_SERVER = 74
Global Const $WINHTTP_OPTION_AUTOLOGON_POLICY = 77
Global Const $WINHTTP_OPTION_SERVER_CERT_CONTEXT = 78
Global Const $WINHTTP_OPTION_ENABLE_FEATURE = 79
Global Const $WINHTTP_OPTION_WORKER_THREAD_COUNT = 80
Global Const $WINHTTP_OPTION_PASSPORT_COBRANDING_TEXT = 81
Global Const $WINHTTP_OPTION_PASSPORT_COBRANDING_URL = 82
Global Const $WINHTTP_OPTION_CONFIGURE_PASSPORT_AUTH = 83
Global Const $WINHTTP_OPTION_SECURE_PROTOCOLS = 84
Global Const $WINHTTP_OPTION_ENABLETRACING = 85
Global Const $WINHTTP_OPTION_PASSPORT_SIGN_OUT = 86
Global Const $WINHTTP_OPTION_REDIRECT_POLICY = 88
Global Const $WINHTTP_OPTION_MAX_HTTP_AUTOMATIC_REDIRECTS = 89
Global Const $WINHTTP_OPTION_MAX_HTTP_STATUS_CONTINUE = 90
Global Const $WINHTTP_OPTION_MAX_RESPONSE_HEADER_SIZE = 91
Global Const $WINHTTP_OPTION_MAX_RESPONSE_DRAIN_SIZE = 92
Global Const $WINHTTP_OPTION_CONNECTION_INFO = 93
Global Const $WINHTTP_OPTION_SPN = 96
Global Const $WINHTTP_OPTION_GLOBAL_PROXY_CREDS = 97
Global Const $WINHTTP_OPTION_GLOBAL_SERVER_CREDS = 98
Global Const $WINHTTP_OPTION_REJECT_USERPWD_IN_URL = 100
Global Const $WINHTTP_OPTION_USE_GLOBAL_SERVER_CREDENTIALS = 101
Global Const $WINHTTP_OPTION_UNSAFE_HEADER_PARSING = 110
Global Const $WINHTTP_OPTION_DECOMPRESSION = 118
Global Const $WINHTTP_OPTION_USERNAME = 0x1000
Global Const $WINHTTP_OPTION_PASSWORD = 0x1001
Global Const $WINHTTP_OPTION_PROXY_USERNAME = 0x1002
Global Const $WINHTTP_OPTION_PROXY_PASSWORD = 0x1003
Global Const $WINHTTP_DECOMPRESSION_FLAG_ALL = 0x00000003
Global Const $WINHTTP_AUTOLOGON_SECURITY_LEVEL_MEDIUM = 0
Global Const $WINHTTP_AUTOLOGON_SECURITY_LEVEL_LOW = 1
Global Const $WINHTTP_AUTOLOGON_SECURITY_LEVEL_HIGH = 2
Global Const $WINHTTP_CALLBACK_STATUS_INTERMEDIATE_RESPONSE = 0x00008000
Global Const $WINHTTP_CALLBACK_STATUS_SECURE_FAILURE = 0x00010000
Global Const $WINHTTP_CALLBACK_STATUS_REQUEST_ERROR = 0x00200000
Global Const $WINHTTP_CALLBACK_FLAG_ALL_NOTIFICATIONS = 0xFFFFFFFF
Global Const $hWINHTTPDLL__WINHTTP = DllOpen("winhttp.dll")
DllOpen("winhttp.dll")
Func _WinHttpCloseHandle($hInternet)
Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpCloseHandle", "handle", $hInternet)
If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
Return 1
EndFunc
Func _WinHttpConnect($hSession, $sServerName, $iServerPort = Default)
Local $aURL = _WinHttpCrackUrl($sServerName), $iScheme = 0
If @error Then
__WinHttpDefault($iServerPort, $INTERNET_DEFAULT_PORT)
Else
$sServerName = $aURL[2]
$iServerPort = $aURL[3]
$iScheme = $aURL[1]
EndIf
Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "handle", "WinHttpConnect", "handle", $hSession, "wstr", $sServerName, "dword", $iServerPort, "dword", 0)
If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
_WinHttpSetOption($aCall[0], $WINHTTP_OPTION_CONTEXT_VALUE, $iScheme)
Return $aCall[0]
EndFunc
Func _WinHttpCrackUrl($sURL, $iFlag = Default)
__WinHttpDefault($iFlag, $ICU_ESCAPE)
Local $tURL_COMPONENTS = DllStructCreate("dword StructSize;" & "ptr SchemeName;" & "dword SchemeNameLength;" & "int Scheme;" & "ptr HostName;" & "dword HostNameLength;" & "word Port;" & "ptr UserName;" & "dword UserNameLength;" & "ptr Password;" & "dword PasswordLength;" & "ptr UrlPath;" & "dword UrlPathLength;" & "ptr ExtraInfo;" & "dword ExtraInfoLength")
DllStructSetData($tURL_COMPONENTS, 1, DllStructGetSize($tURL_COMPONENTS))
Local $tBuffers[6]
Local $iURLLen = StringLen($sURL)
For $i = 0 To 5
$tBuffers[$i] = DllStructCreate("wchar[" & $iURLLen + 1 & "]")
Next
DllStructSetData($tURL_COMPONENTS, "SchemeNameLength", $iURLLen)
DllStructSetData($tURL_COMPONENTS, "SchemeName", DllStructGetPtr($tBuffers[0]))
DllStructSetData($tURL_COMPONENTS, "HostNameLength", $iURLLen)
DllStructSetData($tURL_COMPONENTS, "HostName", DllStructGetPtr($tBuffers[1]))
DllStructSetData($tURL_COMPONENTS, "UserNameLength", $iURLLen)
DllStructSetData($tURL_COMPONENTS, "UserName", DllStructGetPtr($tBuffers[2]))
DllStructSetData($tURL_COMPONENTS, "PasswordLength", $iURLLen)
DllStructSetData($tURL_COMPONENTS, "Password", DllStructGetPtr($tBuffers[3]))
DllStructSetData($tURL_COMPONENTS, "UrlPathLength", $iURLLen)
DllStructSetData($tURL_COMPONENTS, "UrlPath", DllStructGetPtr($tBuffers[4]))
DllStructSetData($tURL_COMPONENTS, "ExtraInfoLength", $iURLLen)
DllStructSetData($tURL_COMPONENTS, "ExtraInfo", DllStructGetPtr($tBuffers[5]))
Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpCrackUrl", "wstr", $sURL, "dword", $iURLLen, "dword", $iFlag, "struct*", $tURL_COMPONENTS)
If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
Local $aRet[8] = [DllStructGetData($tBuffers[0], 1), DllStructGetData($tURL_COMPONENTS, "Scheme"), DllStructGetData($tBuffers[1], 1), DllStructGetData($tURL_COMPONENTS, "Port"), DllStructGetData($tBuffers[2], 1), DllStructGetData($tBuffers[3], 1), DllStructGetData($tBuffers[4], 1), DllStructGetData($tBuffers[5], 1)]
Return $aRet
EndFunc
Func _WinHttpOpen($sUserAgent = Default, $iAccessType = Default, $sProxyName = Default, $sProxyBypass = Default, $iFlag = Default)
__WinHttpDefault($sUserAgent, __WinHttpUA())
__WinHttpDefault($iAccessType, $WINHTTP_ACCESS_TYPE_NO_PROXY)
__WinHttpDefault($sProxyName, $WINHTTP_NO_PROXY_NAME)
__WinHttpDefault($sProxyBypass, $WINHTTP_NO_PROXY_BYPASS)
__WinHttpDefault($iFlag, 0)
Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "handle", "WinHttpOpen", "wstr", $sUserAgent, "dword", $iAccessType, "wstr", $sProxyName, "wstr", $sProxyBypass, "dword", $iFlag)
If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
If $iFlag = $WINHTTP_FLAG_ASYNC Then _WinHttpSetOption($aCall[0], $WINHTTP_OPTION_CONTEXT_VALUE, $WINHTTP_FLAG_ASYNC)
Return $aCall[0]
EndFunc
Func _WinHttpOpenRequest($hConnect, $sVerb = Default, $sObjectName = Default, $sVersion = Default, $sReferrer = Default, $sAcceptTypes = Default, $iFlags = Default)
__WinHttpDefault($sVerb, "GET")
__WinHttpDefault($sObjectName, "")
__WinHttpDefault($sVersion, "HTTP/1.1")
__WinHttpDefault($sReferrer, $WINHTTP_NO_REFERER)
__WinHttpDefault($iFlags, $WINHTTP_FLAG_ESCAPE_DISABLE)
Local $pAcceptTypes
If $sAcceptTypes = Default Or Number($sAcceptTypes) = -1 Then
$pAcceptTypes = $WINHTTP_DEFAULT_ACCEPT_TYPES
Else
Local $aTypes = StringSplit($sAcceptTypes, ",", 2)
Local $tAcceptTypes = DllStructCreate("ptr[" & UBound($aTypes) + 1 & "]")
Local $tType[UBound($aTypes)]
For $i = 0 To UBound($aTypes) - 1
$tType[$i] = DllStructCreate("wchar[" & StringLen($aTypes[$i]) + 1 & "]")
DllStructSetData($tType[$i], 1, $aTypes[$i])
DllStructSetData($tAcceptTypes, 1, DllStructGetPtr($tType[$i]), $i + 1)
Next
$pAcceptTypes = DllStructGetPtr($tAcceptTypes)
EndIf
Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "handle", "WinHttpOpenRequest", "handle", $hConnect, "wstr", StringUpper($sVerb), "wstr", $sObjectName, "wstr", StringUpper($sVersion), "wstr", $sReferrer, "ptr", $pAcceptTypes, "dword", $iFlags)
If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
Return $aCall[0]
EndFunc
Func _WinHttpQueryDataAvailable($hRequest)
Local $sReadType = "dword*"
If BitAND(_WinHttpQueryOption(_WinHttpQueryOption(_WinHttpQueryOption($hRequest, $WINHTTP_OPTION_PARENT_HANDLE), $WINHTTP_OPTION_PARENT_HANDLE), $WINHTTP_OPTION_CONTEXT_VALUE), $WINHTTP_FLAG_ASYNC) Then $sReadType = "ptr"
Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpQueryDataAvailable", "handle", $hRequest, $sReadType, 0)
If @error Then Return SetError(1, 0, 0)
Return SetExtended($aCall[2], $aCall[0])
EndFunc
Func _WinHttpQueryOption($hInternet, $iOption)
Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpQueryOption", "handle", $hInternet, "dword", $iOption, "ptr", 0, "dword*", 0)
If @error Or $aCall[0] Then Return SetError(1, 0, "")
Local $iSize = $aCall[4]
Local $tBuffer
Switch $iOption
Case $WINHTTP_OPTION_CONNECTION_INFO, $WINHTTP_OPTION_PASSWORD, $WINHTTP_OPTION_PROXY_PASSWORD, $WINHTTP_OPTION_PROXY_USERNAME, $WINHTTP_OPTION_URL, $WINHTTP_OPTION_USERNAME, $WINHTTP_OPTION_USER_AGENT, $WINHTTP_OPTION_PASSPORT_COBRANDING_TEXT, $WINHTTP_OPTION_PASSPORT_COBRANDING_URL
$tBuffer = DllStructCreate("wchar[" & $iSize + 1 & "]")
Case $WINHTTP_OPTION_PARENT_HANDLE, $WINHTTP_OPTION_CALLBACK, $WINHTTP_OPTION_SERVER_CERT_CONTEXT
$tBuffer = DllStructCreate("ptr")
Case $WINHTTP_OPTION_CONNECT_TIMEOUT, $WINHTTP_AUTOLOGON_SECURITY_LEVEL_HIGH, $WINHTTP_AUTOLOGON_SECURITY_LEVEL_LOW, $WINHTTP_AUTOLOGON_SECURITY_LEVEL_MEDIUM, $WINHTTP_OPTION_CONFIGURE_PASSPORT_AUTH, $WINHTTP_OPTION_CONNECT_RETRIES, $WINHTTP_OPTION_EXTENDED_ERROR, $WINHTTP_OPTION_HANDLE_TYPE, $WINHTTP_OPTION_MAX_CONNS_PER_1_0_SERVER, $WINHTTP_OPTION_MAX_CONNS_PER_SERVER, $WINHTTP_OPTION_MAX_HTTP_AUTOMATIC_REDIRECTS, $WINHTTP_OPTION_RECEIVE_RESPONSE_TIMEOUT, $WINHTTP_OPTION_RECEIVE_TIMEOUT, $WINHTTP_OPTION_RESOLVE_TIMEOUT, $WINHTTP_OPTION_SECURITY_FLAGS, $WINHTTP_OPTION_SECURITY_KEY_BITNESS, $WINHTTP_OPTION_SEND_TIMEOUT
$tBuffer = DllStructCreate("int")
Case $WINHTTP_OPTION_CONTEXT_VALUE
$tBuffer = DllStructCreate("dword_ptr")
Case Else
$tBuffer = DllStructCreate("byte[" & $iSize & "]")
EndSwitch
$aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpQueryOption", "handle", $hInternet, "dword", $iOption, "struct*", $tBuffer, "dword*", $iSize)
If @error Or Not $aCall[0] Then Return SetError(2, 0, "")
Return DllStructGetData($tBuffer, 1)
EndFunc
Func _WinHttpReadData($hRequest, $iMode = Default, $iNumberOfBytesToRead = Default, $pBuffer = Default)
__WinHttpDefault($iMode, 0)
__WinHttpDefault($iNumberOfBytesToRead, 8192)
Local $tBuffer, $vOutOnError = ""
If $iMode = 2 Then $vOutOnError = Binary($vOutOnError)
Switch $iMode
Case 1, 2
If $pBuffer And $pBuffer <> Default Then
$tBuffer = DllStructCreate("byte[" & $iNumberOfBytesToRead & "]", $pBuffer)
Else
$tBuffer = DllStructCreate("byte[" & $iNumberOfBytesToRead & "]")
EndIf
Case Else
$iMode = 0
If $pBuffer And $pBuffer <> Default Then
$tBuffer = DllStructCreate("char[" & $iNumberOfBytesToRead & "]", $pBuffer)
Else
$tBuffer = DllStructCreate("char[" & $iNumberOfBytesToRead & "]")
EndIf
EndSwitch
Local $sReadType = "dword*"
If BitAND(_WinHttpQueryOption(_WinHttpQueryOption(_WinHttpQueryOption($hRequest, $WINHTTP_OPTION_PARENT_HANDLE), $WINHTTP_OPTION_PARENT_HANDLE), $WINHTTP_OPTION_CONTEXT_VALUE), $WINHTTP_FLAG_ASYNC) Then $sReadType = "ptr"
Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpReadData", "handle", $hRequest, "struct*", $tBuffer, "dword", $iNumberOfBytesToRead, $sReadType, 0)
If @error Or Not $aCall[0] Then Return SetError(1, 0, "")
If Not $aCall[4] Then Return SetError(-1, 0, $vOutOnError)
If $aCall[4] < $iNumberOfBytesToRead Then
Switch $iMode
Case 0
Return SetExtended($aCall[4], StringLeft(DllStructGetData($tBuffer, 1), $aCall[4]))
Case 1
Return SetExtended($aCall[4], BinaryToString(BinaryMid(DllStructGetData($tBuffer, 1), 1, $aCall[4]), 4))
Case 2
Return SetExtended($aCall[4], BinaryMid(DllStructGetData($tBuffer, 1), 1, $aCall[4]))
EndSwitch
Else
Switch $iMode
Case 0, 2
Return SetExtended($aCall[4], DllStructGetData($tBuffer, 1))
Case 1
Return SetExtended($aCall[4], BinaryToString(DllStructGetData($tBuffer, 1), 4))
EndSwitch
EndIf
EndFunc
Func _WinHttpReceiveResponse($hRequest)
Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpReceiveResponse", "handle", $hRequest, "ptr", 0)
If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
Return 1
EndFunc
Func _WinHttpSendRequest($hRequest, $sHeaders = Default, $vOptional = Default, $iTotalLength = Default, $iContext = Default)
__WinHttpDefault($sHeaders, $WINHTTP_NO_ADDITIONAL_HEADERS)
__WinHttpDefault($vOptional, $WINHTTP_NO_REQUEST_DATA)
__WinHttpDefault($iTotalLength, 0)
__WinHttpDefault($iContext, 0)
Local $pOptional = 0, $iOptionalLength = 0
If @NumParams > 2 Then
Local $tOptional
$iOptionalLength = BinaryLen($vOptional)
$tOptional = DllStructCreate("byte[" & $iOptionalLength & "]")
If $iOptionalLength Then $pOptional = DllStructGetPtr($tOptional)
DllStructSetData($tOptional, 1, $vOptional)
EndIf
If Not $iTotalLength Or $iTotalLength < $iOptionalLength Then $iTotalLength += $iOptionalLength
Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpSendRequest", "handle", $hRequest, "wstr", $sHeaders, "dword", 0, "ptr", $pOptional, "dword", $iOptionalLength, "dword", $iTotalLength, "dword_ptr", $iContext)
If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
Return 1
EndFunc
Func _WinHttpSetOption($hInternet, $iOption, $vSetting, $iSize = Default)
If $iSize = Default Then $iSize = -1
If IsBinary($vSetting) Then
$iSize = DllStructCreate("byte[" & BinaryLen($vSetting) & "]")
DllStructSetData($iSize, 1, $vSetting)
$vSetting = $iSize
$iSize = DllStructGetSize($vSetting)
EndIf
Local $sType
Switch $iOption
Case $WINHTTP_OPTION_AUTOLOGON_POLICY, $WINHTTP_OPTION_CODEPAGE, $WINHTTP_OPTION_CONFIGURE_PASSPORT_AUTH, $WINHTTP_OPTION_CONNECT_RETRIES, $WINHTTP_OPTION_CONNECT_TIMEOUT, $WINHTTP_OPTION_DISABLE_FEATURE, $WINHTTP_OPTION_ENABLE_FEATURE, $WINHTTP_OPTION_ENABLETRACING, $WINHTTP_OPTION_MAX_CONNS_PER_1_0_SERVER, $WINHTTP_OPTION_MAX_CONNS_PER_SERVER, $WINHTTP_OPTION_MAX_HTTP_AUTOMATIC_REDIRECTS, $WINHTTP_OPTION_MAX_HTTP_STATUS_CONTINUE, $WINHTTP_OPTION_MAX_RESPONSE_DRAIN_SIZE, $WINHTTP_OPTION_MAX_RESPONSE_HEADER_SIZE, $WINHTTP_OPTION_READ_BUFFER_SIZE, $WINHTTP_OPTION_RECEIVE_TIMEOUT, $WINHTTP_OPTION_RECEIVE_RESPONSE_TIMEOUT, $WINHTTP_OPTION_REDIRECT_POLICY, $WINHTTP_OPTION_REJECT_USERPWD_IN_URL, $WINHTTP_OPTION_REQUEST_PRIORITY, $WINHTTP_OPTION_RESOLVE_TIMEOUT, $WINHTTP_OPTION_SECURE_PROTOCOLS, $WINHTTP_OPTION_SECURITY_FLAGS, $WINHTTP_OPTION_SECURITY_KEY_BITNESS, $WINHTTP_OPTION_SEND_TIMEOUT, $WINHTTP_OPTION_SPN, $WINHTTP_OPTION_USE_GLOBAL_SERVER_CREDENTIALS, $WINHTTP_OPTION_WORKER_THREAD_COUNT, $WINHTTP_OPTION_WRITE_BUFFER_SIZE, $WINHTTP_OPTION_DECOMPRESSION, $WINHTTP_OPTION_UNSAFE_HEADER_PARSING
$sType = "dword*"
$iSize = 4
Case $WINHTTP_OPTION_CALLBACK, $WINHTTP_OPTION_PASSPORT_SIGN_OUT
$sType = "ptr*"
$iSize = 4
If @AutoItX64 Then $iSize = 8
If Not IsPtr($vSetting) Then Return SetError(3, 0, 0)
Case $WINHTTP_OPTION_CONTEXT_VALUE
$sType = "dword_ptr*"
$iSize = 4
If @AutoItX64 Then $iSize = 8
Case $WINHTTP_OPTION_PASSWORD, $WINHTTP_OPTION_PROXY_PASSWORD, $WINHTTP_OPTION_PROXY_USERNAME, $WINHTTP_OPTION_USER_AGENT, $WINHTTP_OPTION_USERNAME
$sType = "wstr"
If(IsDllStruct($vSetting) Or IsPtr($vSetting)) Then Return SetError(3, 0, 0)
If $iSize < 1 Then $iSize = StringLen($vSetting)
Case $WINHTTP_OPTION_CLIENT_CERT_CONTEXT, $WINHTTP_OPTION_GLOBAL_PROXY_CREDS, $WINHTTP_OPTION_GLOBAL_SERVER_CREDS, $WINHTTP_OPTION_HTTP_VERSION, $WINHTTP_OPTION_PROXY
$sType = "ptr"
If Not(IsDllStruct($vSetting) Or IsPtr($vSetting)) Then Return SetError(3, 0, 0)
Case Else
Return SetError(1, 0, 0)
EndSwitch
If $iSize < 1 Then
If IsDllStruct($vSetting) Then
$iSize = DllStructGetSize($vSetting)
Else
Return SetError(2, 0, 0)
EndIf
EndIf
Local $aCall
If IsDllStruct($vSetting) Then
$aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpSetOption", "handle", $hInternet, "dword", $iOption, $sType, DllStructGetPtr($vSetting), "dword", $iSize)
Else
$aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpSetOption", "handle", $hInternet, "dword", $iOption, $sType, $vSetting, "dword", $iSize)
EndIf
If @error Or Not $aCall[0] Then Return SetError(4, 0, 0)
Return 1
EndFunc
Func _WinHttpSetStatusCallback($hInternet, $hInternetCallback, $iNotificationFlags = Default)
__WinHttpDefault($iNotificationFlags, $WINHTTP_CALLBACK_FLAG_ALL_NOTIFICATIONS)
Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "ptr", "WinHttpSetStatusCallback", "handle", $hInternet, "ptr", DllCallbackGetPtr($hInternetCallback), "dword", $iNotificationFlags, "ptr", 0)
If @error Then Return SetError(1, 0, 0)
Return $aCall[0]
EndFunc
Func __WinHttpDefault(ByRef $vInput, $vOutput)
If $vInput = Default Or Number($vInput) = -1 Then $vInput = $vOutput
EndFunc
Func __WinHttpUA()
Local Static $sUA = "Mozilla/5.0 " & __WinHttpSysInfo() & " WinHttp/" & __WinHttpVer() & " (WinHTTP/5.1) like Gecko"
Return $sUA
EndFunc
Func __WinHttpSysInfo()
Local $sDta = FileGetVersion("kernel32.dll")
$sDta = "(Windows NT " & StringLeft($sDta, StringInStr($sDta, ".", 1, 2) - 1)
If StringInStr(@OSArch, "64") And Not @AutoItX64 Then $sDta &= "; WOW64"
$sDta &= ")"
Return $sDta
EndFunc
Func __WinHttpVer()
Return "1.6.4.1"
EndFunc
Global Const $VERSION = "1.2310.512.5402"
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
Global $sDestPath
Global $aScanners[1][3]
$aScanners[0][0]=0
Global $bCheckUpdate=True
Global $g_iProgPerc, $g_iProgBytes, $g_iProgTotal, $g_iUpdateStat
Global $hWINHTTP_STATUS_CALLBACK = DllCallbackRegister("__WINHTTP_STATUS_CALLBACK", "none", "handle;dword_ptr;dword;ptr;dword")
Global $sMsgPostInstall='Success! A shortcut named "Scan to PDF" has been added to your desktop.'&@LF&@LF
$sMsgPostInstall&="NOTE: Scans will be stored in:"&@LF
Global $aErrMsgs[]=[ "The selected scanner could not be found.", "The selected scanner is offline.", "No scanning device was found.", "No pages are in the feeder.", "The scanner's cover is open.", "The scanner has a paper jam.", "The selected scanner is busy.", "The scanner's cover is open.", "The scanner has a paper jam.", "The scanner is warming up.", "An error occurred with the scanning driver.", "An error occurred when trying to save the file.", "An unknown error occurred during the batch scan.", "Batch scan stopped due to error.", "The scanner is warming up.", "No scanned pages to export.", "The selected driver is not supported on this system.", "The selected scanner does not support using duplex. If your scanner is supposed to support duplex, try using a different driver.", "The selected scanner does not support using a feeder. If your scanner does have a feeder, try using a different driver." ]
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
$sDestPath=FileSelectFolder($sTitle&" - Choose Scan Folder",$sInitialSaveDir,7)
$iRet=RegWrite($sRegKey,"ScanPath","REG_SZ",$sDestPath)
If $iRet<>1 Then
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
$sDestPath=RegRead($sRegKey,"ScanPath")
If @error Then
$sDestPath=FileSelectFolder($sTitle&" - Choose Scan Folder",$sInitialSaveDir,7)
$iRet=RegWrite($sRegKey,"ScanPath","REG_SZ",$sDestPath)
If $iRet<>1 Then
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
GUICtrlCreateGroup("Document Source", $iGuiComL*$iScale, 33*$iScale,($iGuiWidth-8)*$iScale, 126*$iScale)
$idRadFlat = GUICtrlCreateRadio("Flatbed",($iGuiComL+$iGuiComGrpL)*$iScale, $iGuiScanOptT*$iScale, $iGuiComOptW*$iScale, $iGuiComOptH*$iScale)
$idRadFeed = GUICtrlCreateRadio("Document Feeder",($iGuiComL+$iGuiComGrpL)*$iScale,($iGuiScanOptT+$iGuiComOptH)*$iScale, $iGuiComOptW*$iScale, $iGuiComOptH*$iScale)
$idRadDplx = GUICtrlCreateRadio("Duplex Tray",($iGuiComL+$iGuiComGrpL)*$iScale,($iGuiScanOptT+($iGuiComOptH*2))*$iScale, $iGuiComOptW*$iScale, $iGuiComOptH*$iScale)
GUICtrlCreateGroup("Options", $iGuiComL*$iScale, 101*$iScale,($iGuiWidth-8)*$iScale, 58*$iScale)
$idChkRemb = GUICtrlCreateCheckbox("Remember my last choices",($iGuiComL+$iGuiComGrpL)*$iScale,(101+$iGuiOptT)*$iScale,($iGuiWidth-32)*$iScale, $iGuiComOptH*$iScale)
$idChkStay = GUICtrlCreateCheckbox("Stay Open (Batch Scan)",($iGuiComL+$iGuiComGrpL)*$iScale,(101+($iGuiOptT*2))*$iScale,($iGuiWidth-32)*$iScale, $iGuiComOptH*$iScale)
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
If $iScanner<>-1 Then GUICtrlSetState($idBtnScan,$GUI_ENABLE)
Case $idRadFeed
If BitAnd(GUICtrlRead($iId),$GUI_CHECKED) = $GUI_CHECKED Then $iScanSrc=0x2
If $iScanner<>-1 Then GUICtrlSetState($idBtnScan,$GUI_ENABLE)
Case $idRadDplx
If BitAnd(GUICtrlRead($iId),$GUI_CHECKED) = $GUI_CHECKED Then $iScanSrc=0x4
If $iScanner<>-1 Then GUICtrlSetState($idBtnScan,$GUI_ENABLE)
Case $idChkRemb
$bRemember=BitAnd(GUICtrlRead($iId),$GUI_CHECKED) = $GUI_CHECKED
Case $idChkStay
$bStay=BitAnd(GUICtrlRead($iId),$GUI_CHECKED) = $GUI_CHECKED
EndSwitch
EndSwitch
Return $GUI_RUNDEFMSG
EndFunc
Func _UpdProfile()
Local $bWNPF
Local $bWNPD
Local $bWNPG
Local $sNewProfile
Local $sProfilePath="Profiles.xml"
Local $bExists=FileExists($sProfilePath)
Local $hFile
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
Local $aProfiles=StringRegExp($sData,"<ScanProfile>((?:.|\s)*?)<\/ScanProfile>",3)
If Not @error Then
For $sProfile In $aProfiles
$sDisplayName=_XmlGetField($sProfile,"DisplayName")
If @error Or $sDisplayName='' Then
_Log("Warning: Skipping the '"&$sDisplayName&"' profile, it doesn't have a name.")
ContinueLoop
EndIf
If Not StringInStr($sAlias&"_Feeder"&$sAlias&"_Duplex"&$sAlias&"_Glass",$sDisplayName) Then ContinueLoop
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
Func _RandStr()
Local $sRet = "", $aTmp[3], $iLen = 16
For $i = 1 To $iLen
$aTmp[0] = Chr(Random(65, 90, 1))
$aTmp[1] = Chr(Random(97, 122, 1))
$aTmp[2] = Chr(Random(48, 57, 1))
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
Local $sChunk, $sData, $bAvail=_WinHttpQueryDataAvailable($hRequest)
Local $iTotal=@Extended
If $bAvail Then
If $fCallback<>-1 Then Call($fCallback,2,0,$iTotal)
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
Func _Update($bPost=0)
Local $sTitle=$sAlias&" Update"
_Log("Current Version: "&$VERSION,"_Update")
If $bPost Then
_Log("################### Update Stage 2 ###################","_Update")
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
Return SetError(0,3,1)
EndIf
If $vVer=1 Then
_Log("Warning: upstream version is older than self.","_Update")
Return SetError(1,9,0)
EndIf
If $vVer<>-1 Then
_Log("Error, unexpected update variable! ("&$vVer&')',"_Update")
Return SetError(1,10,0)
EndIf
Local $iRet=MsgBox(32+4+65536,$sTitle,"An update is available!"&@LF&@LF&"Current version: "&$VERSION&@LF&"New version: "&$sRet&@LF&@LF&"Would you like to update now?",0,$hMain)
If $iRet<>6 Then Return SetError(0,4,1)
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
EndIf
AdlibUnregister("__UpdateProgWatch")
DirCreate($sSnapsDir)
If Not _isDir($sSnapsDir) Then
MsgBox(16,$sTitle,"Error: Update Failed, see log for details.")
_Log("Error: Failed to create snaphot directory.","_Update")
Return SetError(0,13,1)
EndIf
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
Local $sStatus
Switch $iInternetStatus
Case $WINHTTP_CALLBACK_STATUS_INTERMEDIATE_RESPONSE
$sStatus = "Received an intermediate (100 level) status code message from the server."
Case $WINHTTP_CALLBACK_STATUS_REQUEST_ERROR
$sStatus = "An error occurred while sending an HTTP request."
Case $WINHTTP_CALLBACK_STATUS_SECURE_FAILURE
$sStatus = "One or more errors were encountered while retrieving a Secure Sockets Layer (SSL) certificate from the server."
EndSwitch
_Log($sStatus,"__Update_SecureGet")
EndFunc
