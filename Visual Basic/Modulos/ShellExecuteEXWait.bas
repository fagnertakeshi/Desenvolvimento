Attribute VB_Name = "ShellExecuteEXWait"
'Autor: Isnard (Consulta na Internet -> http://www.codecomments.com/archive293-2005-9-617106.html)

Option Explicit
Option Base 1

Public Enum ENUM_SHELLOPERATION
esoOpen
esoExplore
esoPrint
esofind
esoProperties
esoEdit
End Enum

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function ShellExecuteEx Lib "shell32.dll" Alias _
"ShellExecuteExA" (ByRef shellInfo As SHELLEXECUTEINFO) As Long

Private Type SHELLEXECUTEINFO
cbSize As Long
fMask As Long
hwnd As Long
lpVerb As String
lpFile As String
lpParameters As String
lpDirectory As String
nShow As Long
hInstApp As Long
' Optional fields
lpIDList As Long
lpClass As String
hkeyClass As Long
dwHotKey As Long
hIcon As Long
hProcess As Long
End Type


Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_DDEWAIT = &H100

Private Const INVALID_HANDLE_VALUE = -1
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const SE_ERR_ACCESSDENIED = 5
Private Const SE_ERR_ASSOCINCOMPLETE = 27
Private Const SE_ERR_DDEBUSY = 30
Private Const SE_ERR_DDEFAIL = 29
Private Const SE_ERR_DDETIMEOUT = 28
Private Const SE_ERR_DLLNOTFOUND = 32
Private Const SE_ERR_FNF = 2
Private Const SE_ERR_NOASSOC = 31
Private Const SE_ERR_OOM = 8
Private Const SE_ERR_PNF = 3
Private Const SE_ERR_SHARE = 26

Private Const SW_HIDE = 0
Private Const SW_RESTORE = 9
Private Const SW_SHOW = 5
Private Const SW_SHOWDEFAULT = 10
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMINNOACTIVE = 7
Private Const SW_SHOWNA = 8
Private Const SW_SHOWNOACTIVATE = 4
Private Const SW_SHOWNORMAL = 1


Private Declare Function GetProcessTimes Lib "kernel32" (ByVal hProcess As Long, _
                                                         lpCreationTime As FILETIME, _
                                                         lpExitTime As FILETIME, _
                                                         lpKernelTime As FILETIME, _
                                                         lpUserTime As FILETIME) As Long

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Declare Function GetProcessVersion Lib "kernel32" (ByVal ProcessId As Long) As Long

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long _

Private Const PROCESS_TERMINATE = &H1
Private Const PROCESS_CREATE_THREAD = &H2
Private Const PROCESS_SET_SESSIONID = &H4
Private Const PROCESS_VM_OPERATION = &H8
Private Const PROCESS_VM_READ = &H10
Private Const PROCESS_VM_WRITE = &H20
Private Const PROCESS_DUP_HANDLE = &H40
Private Const PROCESS_CREATE_PROCESS = &H80
Private Const PROCESS_SET_QUOTA = &H100
Private Const PROCESS_SET_INFORMATION = &H200
Private Const PROCESS_QUERY_INFORMATION = &H400

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle _
As Long, ByVal dwMilliseconds As Long) As Long

Private Const INFINITE = &HFFFFFFFF ' dwMilliseconds TimeOut value
Private Const STATUS_WAIT_0 = &H0
Private Const STATUS_ABANDONED_WAIT_0 = &H80
Private Const STATUS_USER_APC = &HC0
Private Const STATUS_TIMEOUT = &H102
Private Const STATUS_PENDING = &H103

Private Const WAIT_FAILED = &HFFFFFFFF
Private Const WAIT_OBJECT_0 = ((STATUS_WAIT_0) + 0)

Private Const WAIT_ABANDONED = ((STATUS_ABANDONED_WAIT_0) + 0)
Private Const WAIT_ABANDONED_0 = ((STATUS_ABANDONED_WAIT_0) + 0)

Private Const WAIT_TIMEOUT = STATUS_TIMEOUT
Private Const WAIT_IO_COMPLETION = STATUS_USER_APC
Private Const STILL_ACTIVE = STATUS_PENDING


Public Function ShellAndWait(PathName, _
                             Optional WindowStyle As VbAppWinStyle = vbMinimizedFocus) As Boolean

Dim nPID As Long ' Win32 ProcessID
    
    nPID = Shell(PathName, WindowStyle)
    Call WaitProcess(nPID, True)
    ShellAndWait = (nPID <> 0)
    
End Function

Public Function ShellExecuteAndWait(ByVal veOperation As ENUM_SHELLOPERATION, _
                                    ByVal vsFile As String, _
                                    ByVal vsParameters As String, _
                                    ByVal vsDirectory As String, _
                                    Optional WindowStyle As VbAppWinStyle = vbMinimizedFocus) As Boolean


Dim nHandle As Long ' Win32 ProcessHandle
Dim si As SHELLEXECUTEINFO
Dim sOp As String
Dim nResult As Long
Dim nShowCmd As Long
    
    Select Case veOperation
        Case esoExplore: sOp = "Explore"
        Case esoOpen: sOp = "Open"
        Case esoPrint: sOp = "Print"
        Case esofind: sOp = "Find"
        Case esoProperties: sOp = "Properties"
        Case esoEdit: sOp = "Edit"
    End Select

    With si
        .cbSize = Len(si)
        .fMask = SEE_MASK_NOCLOSEPROCESS
        .nShow = WindowStyle2SW(WindowStyle)
        .lpFile = vsFile
        .lpParameters = vsParameters
        .lpVerb = sOp
    End With

    nResult = ShellExecuteEx(si)
    nHandle = si.hProcess
    
    Call WaitProcess(0, nHandle, True)
    ShellExecuteAndWait = (nResult > 32)
    
End Function


Private Function WaitProcess(ByRef rnPID As Long, _
                             ByRef rnHandle As Long, _
                             Optional Wait As Boolean = False) As Boolean

' Return True when process was created
Const C_WaitRefreshInterval = 800 ' msec
Dim hP As Long ' Win32 ProcessHandle

    WaitProcess = False

    If rnHandle = 0 Then
        hP = OpenProcess(PROCESS_QUERY_INFORMATION, 0, rnPID)
    Else
        hP = rnHandle
    End If

On Error GoTo Finally
    
    If hP <> 0 Then
        WaitProcess = True
        While Wait
        
            ' Wait with timeout in msec (INFINITE, 0 = no wait)
            Select Case WaitForSingleObject(hP, C_WaitRefreshInterval)
                Case WAIT_OBJECT_0
                    ' process is terminated
                    Wait = False
                Case WAIT_TIMEOUT
                    ' run windows messageloop to refresh screen paints (if any)
                    DoEvents
                Case Else ' WAIT_FAILED
                    Wait = False
            End Select
        Wend
    Else
        hP = INVALID_HANDLE_VALUE
    End If

Finally:
    CloseHandle hP

End Function

Private Function WindowStyle2SW(ByVal vnWS As VbAppWinStyle) As Long

    Select Case vnWS
        Case vbHide: WindowStyle2SW = SW_HIDE
        Case vbMaximizedFocus: WindowStyle2SW = SW_SHOWMAXIMIZED
        Case vbMinimizedFocus: WindowStyle2SW = SW_SHOWMINIMIZED
        Case vbMinimizedNoFocus: WindowStyle2SW = SW_SHOWMINNOACTIVE
        Case vbNormalFocus: WindowStyle2SW = SW_SHOWNORMAL
        Case vbNormalNoFocus: WindowStyle2SW = SW_SHOWNA
    End Select

End Function

