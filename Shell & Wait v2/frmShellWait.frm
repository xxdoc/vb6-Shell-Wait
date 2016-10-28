VERSION 5.00
Begin VB.Form frmShellWait 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shell & Wait Demo"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8880
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      Caption         =   "&PathName"
      ClipControls    =   0   'False
      Height          =   3315
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   8715
      Begin VB.PictureBox PictBox 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         HasDC           =   0   'False
         Height          =   3015
         Index           =   0
         Left            =   75
         ScaleHeight     =   3015
         ScaleWidth      =   8565
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   225
         Width           =   8565
         Begin VB.Frame Frame 
            Caption         =   "     &Wait (in milliseconds)"
            ClipControls    =   0   'False
            Height          =   690
            Index           =   2
            Left            =   6075
            TabIndex        =   16
            Top             =   525
            Width           =   2415
            Begin VB.TextBox TxtBox 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   150
               TabIndex        =   18
               Top             =   300
               Width           =   2115
            End
            Begin VB.CheckBox ChkBox 
               Height          =   315
               Left            =   150
               TabIndex        =   17
               Top             =   0
               Width           =   240
            End
         End
         Begin VB.Frame Frame 
            Caption         =   "Window&Style"
            ClipControls    =   0   'False
            Height          =   1665
            Index           =   1
            Left            =   75
            TabIndex        =   3
            Top             =   525
            Width           =   5865
            Begin VB.PictureBox PictBox 
               BorderStyle     =   0  'None
               ClipControls    =   0   'False
               HasDC           =   0   'False
               Height          =   1365
               Index           =   1
               Left            =   75
               ScaleHeight     =   1365
               ScaleWidth      =   5715
               TabIndex        =   4
               TabStop         =   0   'False
               Top             =   225
               Width           =   5715
               Begin VB.OptionButton OptBtn 
                  Caption         =   "vbShowDefault"
                  Height          =   240
                  Index           =   10
                  Left            =   4080
                  TabIndex        =   15
                  Top             =   705
                  Width           =   1590
               End
               Begin VB.OptionButton OptBtn 
                  Caption         =   "vbRestore"
                  Height          =   240
                  Index           =   9
                  Left            =   4080
                  TabIndex        =   14
                  Top             =   390
                  Width           =   1590
               End
               Begin VB.OptionButton OptBtn 
                  Caption         =   "vbShowNA"
                  Height          =   240
                  Index           =   8
                  Left            =   4080
                  TabIndex        =   13
                  Top             =   75
                  Width           =   1590
               End
               Begin VB.OptionButton OptBtn 
                  Caption         =   "vbShowMinNoActive"
                  Height          =   240
                  Index           =   7
                  Left            =   2040
                  TabIndex        =   12
                  Top             =   1020
                  Width           =   2040
               End
               Begin VB.OptionButton OptBtn 
                  Caption         =   "vbMinimize"
                  Height          =   240
                  Index           =   6
                  Left            =   2040
                  TabIndex        =   11
                  Top             =   705
                  Width           =   2040
               End
               Begin VB.OptionButton OptBtn 
                  Caption         =   "vbShow"
                  Height          =   240
                  Index           =   5
                  Left            =   2040
                  TabIndex        =   10
                  Top             =   390
                  Width           =   2040
               End
               Begin VB.OptionButton OptBtn 
                  Caption         =   "vbShowNoActivate"
                  Height          =   240
                  Index           =   4
                  Left            =   2040
                  TabIndex        =   9
                  Top             =   75
                  Width           =   2040
               End
               Begin VB.OptionButton OptBtn 
                  Caption         =   "vbShowMaximized"
                  Height          =   240
                  Index           =   3
                  Left            =   75
                  TabIndex        =   8
                  Top             =   1020
                  Width           =   1890
               End
               Begin VB.OptionButton OptBtn 
                  Caption         =   "vbShowMinimized"
                  Height          =   240
                  Index           =   2
                  Left            =   75
                  TabIndex        =   7
                  Top             =   705
                  Width           =   1890
               End
               Begin VB.OptionButton OptBtn 
                  Caption         =   "vbShowNormal"
                  Height          =   240
                  Index           =   1
                  Left            =   75
                  TabIndex        =   6
                  Top             =   390
                  Width           =   1890
               End
               Begin VB.OptionButton OptBtn 
                  Caption         =   "vbHide"
                  Height          =   240
                  Index           =   0
                  Left            =   75
                  TabIndex        =   5
                  Top             =   75
                  Width           =   1890
               End
            End
         End
         Begin VB.CommandButton CmdBtn 
            Caption         =   "&3. WScript.Shell.Run"
            Height          =   615
            Index           =   2
            Left            =   5775
            TabIndex        =   21
            Top             =   2325
            Width           =   2715
         End
         Begin VB.CommandButton CmdBtn 
            Caption         =   "&2. ShellW"
            Height          =   615
            Index           =   1
            Left            =   2925
            TabIndex        =   20
            Top             =   2325
            Width           =   2715
         End
         Begin VB.CommandButton CmdBtn 
            Caption         =   "&1. Shell_n_Wait"
            Height          =   615
            Index           =   0
            Left            =   75
            TabIndex        =   19
            Top             =   2325
            Width           =   2715
         End
         Begin VB.ComboBox ComBoX 
            Height          =   360
            Index           =   0
            Left            =   75
            TabIndex        =   2
            Top             =   75
            Width           =   8415
         End
      End
   End
End
Attribute VB_Name = "frmShellWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CW_USEDEFAULT      As Long = &H80000000
Private Const HWND_TOPMOST       As Long = (-1&)
Private Const SHACF_FILESYSTEM   As Long = &H1
Private Const SWP_NOSIZE         As Long = &H1
Private Const SWP_NOMOVE         As Long = &H2
Private Const SWP_NOACTIVATE     As Long = &H10
Private Const TTF_IDISHWND       As Long = &H1
Private Const TTF_SUBCLASS       As Long = &H10
Private Const TTS_NOPREFIX       As Long = &H2
Private Const WM_USER            As Long = &H400
Private Const TTM_ADDTOOLW       As Long = (WM_USER + 50)
Private Const TTM_SETMAXTIPWIDTH As Long = (WM_USER + 24)
Private Const TOOLTIPS_CLASSW    As String = "tooltips_class32"

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Type TOOLINFO
    cbSize     As Long
    uFlags     As Long
    hWnd       As Long
    uId        As Long
    Rect_      As RECT
    hInst      As Long
    lpszText   As Long
    lParam     As Long
    lpReserved As Long
End Type

Private Declare Function CreateWindowExW Lib "user32.dll" (Optional ByVal dwExStyle As Long, Optional ByVal lpClassName As Long, Optional ByVal lpWindowName As Long, Optional ByVal dwStyle As Long, Optional ByVal X As Long = CW_USEDEFAULT, Optional ByVal Y As Long = CW_USEDEFAULT, Optional ByVal nWidth As Long = CW_USEDEFAULT, Optional ByVal nHeight As Long = CW_USEDEFAULT, Optional ByVal hWndParent As Long, Optional ByVal hMenu As Long, Optional ByVal hInstance As Long, Optional ByVal lpParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function FindWindowExW Lib "user32.dll" (Optional ByVal hWndParent As Long, Optional ByVal hWndChildAfter As Long, Optional ByVal lpszClass As Long, Optional ByVal lpszWindow As Long) As Long
Private Declare Function GetShortPathNameW Lib "kernel32.dll" (ByVal lpszLongPath As Long, Optional ByVal lpszShortPath As Long, Optional ByVal cchBuffer As Long) As Long
Private Declare Function MessageBoxA Lib "user32.dll" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As VbMsgBoxStyle) As VbMsgBoxResult
Private Declare Function SendMessageW Lib "user32.dll" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SetTopMost Lib "user32.dll" Alias "SetWindowPos" (ByVal hWnd As Long, Optional ByVal hWndInsertAfter As Long = HWND_TOPMOST, Optional ByVal X As Long, Optional ByVal Y As Long, Optional ByVal cx As Long, Optional ByVal cy As Long, Optional ByVal uFlags As Long = SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE) As Long
Private Declare Function SHAutoComplete Lib "shlwapi.dll" (ByVal hWndEdit As Long, ByVal dwFlags As Long) As Long
Private Declare Function SysReAllocStringLen Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long, Optional ByVal Length As Long) As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private intWindowStyle As Integer, hWndTT As Long, TI As TOOLINFO

Private Sub ChkBox_Click()
    TxtBox(0).Enabled = Not TxtBox(0).Enabled
End Sub

Private Sub CmdBtn_Click(Index As Integer)
    On Error GoTo 1

    Select Case Index
        Case 0: MessageBoxA 0&, "Exit Code = " & Shell_n_Wait(ComBoX(0), intWindowStyle), App.Title, vbInformation

        Case 1: MessageBoxA 0&, "Return Value = " & ShellW(ComBoX(0), intWindowStyle, IIf(TxtBox(0).Enabled, CLng(TxtBox(0)), _
                            0&)) & vbCr & "Error &H" & Hex$(Err) & ": """ & Err.Description & """", App.Title, vbInformation

        Case 2: MessageBoxA 0&, "Exit Code = " & ShellWS(ComBoX(0), intWindowStyle, CBool(TxtBox(0)) And TxtBox(0).Enabled), _
                            App.Title, vbInformation
    End Select

    If g_ExitDoLoops Then Exit Sub 'Avoid loading the Form again (by referencing 1 of its controls below) when it has already unloaded

    For Index = ComBoX(0).ListCount - 1 To 0 Step -1
        If ComBoX(0) = ComBoX(0).List(Index) Then Exit Sub
    Next:  ComBoX(0).AddItem ComBoX(0), 0:        Exit Sub

1   MsgBox Err.Description, vbCritical, "Error " & Err
End Sub

Private Sub Form_Activate()
    PictBox(0).SetFocus
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim ComboEdit_hWnd As Long

    ChkBox = vbChecked
    OptBtn(1) = True
    TxtBox(0) = "&HFFFFFFFF"

    ComBoX(0).AddItem "%ComSpec% /c Title Hi %UserName%!&Ping 127.0.0.1&Echo.&Pause&Exit 1000"
    ComBoX(0).AddItem "%SystemRoot%\explorer /e,/select,%Temp%"
    ComBoX(0).AddItem "rundll32 shell32.dll,Control_RunDLL inetcpl.cpl,@0,1"    'http://www.robvanderwoude.com/rundll.php#ControlPanelApplets
    ComBoX(0).AddItem """%ProgramFiles%\..\WINDOWS\system32\calc"""
    ComBoX(0).AddItem """%CommonProgramFiles%\..\..\WINDOWS\system32\winver"""
    ComBoX(0).AddItem """%ProgramFiles%\.\Common Files\..\..%HomePath%\..\..\WINDOWS\.\system32\..\system32\cmd"" /c start /wait rundll32 shell32.dll,Control_RunDLL inetcpl.cpl,@0,1"
    If App.LogMode Then ComBoX(0).AddItem "notepad " & GetShortPathName(App.Path & "\" & App.EXEName & ".exe")
    ComBoX(0).ListIndex = 0

    ComboEdit_hWnd = FindWindowExW(ComBoX(0).hWnd, , StrPtr("Edit"))
    SHAutoComplete ComboEdit_hWnd, SHACF_FILESYSTEM

    hWndTT = CreateWindowExW(lpClassName:=StrPtr(TOOLTIPS_CLASSW), dwStyle:=TTS_NOPREFIX, hWndParent:=hWnd, hInstance:=App.hInstance)

    If hWndTT Then
        SetTopMost hWndTT
        TI.cbSize = LenB(TI)
        TI.hWnd = hWnd
        TI.hInst = App.hInstance
        TI.uFlags = TTF_IDISHWND Or TTF_SUBCLASS
        SendMessageW hWndTT, TTM_SETMAXTIPWIDTH, 0&, 240&

        AddTool ComBoX(0).hWnd, "Enter path (quoted if with spaces) to executable or document file." & vbCrLf & "May include environment variables and/or arguments."
        AddTool ComboEdit_hWnd, "Enter path (quoted if with spaces) to executable or document file." & vbCrLf & "May include environment variables and/or arguments."
        AddTool TxtBox(0).hWnd, "For ShellW, type -1 or &HFFFFFFFF (INFINITE) to wait indefinitely." & vbCrLf & "For WScript, any non-zero value will do."
        AddTool CmdBtn(0).hWnd, "Extends the native Shell function by waiting for the shelled" & vbCrLf & "program's termination without blocking other events."
        AddTool CmdBtn(1).hWnd, "Launches an executable file or registered file type;" & vbCrLf & "optionally waits for the specified duration before returning."
        AddTool CmdBtn(2).hWnd, "Runs a program in a new process."
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If hWndTT Then hWndTT = DestroyWindow(hWndTT):  Debug.Assert hWndTT
    g_ExitDoLoops = True
End Sub

Private Sub OptBtn_Click(Index As Integer)
    intWindowStyle = Index
End Sub

Private Sub TxtBox_GotFocus(Index As Integer)
    TxtBox(Index).SelStart = 0&
    TxtBox(Index).SelLength = Len(TxtBox(Index))
End Sub

Private Sub AddTool(ByVal hWndTool As Long, ByRef ToolTipText As String)
    TI.uId = hWndTool
    TI.lpszText = StrPtr(ToolTipText)
    SendMessageW hWndTT, TTM_ADDTOOLW, 0&, TI
End Sub

Private Function GetShortPathName(ByRef sLongPathName As String) As String
    SysReAllocStringLen VarPtr(GetShortPathName), , GetShortPathNameW(StrPtr(sLongPathName)) - 1&
    GetShortPathNameW StrPtr(sLongPathName), StrPtr(GetShortPathName), Len(GetShortPathName) + 1&
End Function
