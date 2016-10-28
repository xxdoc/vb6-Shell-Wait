VERSION 5.00
Begin VB.Form frmShellWait 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shell & Wait"
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   232
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   592
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      Caption         =   "&PathName"
      Height          =   3315
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   8715
      Begin VB.PictureBox PictBox 
         BorderStyle     =   0  'None
         Height          =   3015
         Index           =   0
         Left            =   75
         ScaleHeight     =   201
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   571
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   225
         Width           =   8565
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   8025
            Top             =   1275
         End
         Begin VB.Frame Frame 
            Caption         =   "     &Wait (in milliseconds)"
            Height          =   690
            Index           =   2
            Left            =   6075
            TabIndex        =   15
            Top             =   525
            Width           =   2415
            Begin VB.TextBox TxtBox 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Terminal"
                  Size            =   9
                  Charset         =   255
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   150
               TabIndex        =   17
               Top             =   300
               Width           =   2115
            End
            Begin VB.CheckBox ChkBox 
               Height          =   315
               Index           =   0
               Left            =   150
               TabIndex        =   16
               Top             =   0
               Value           =   1  'Checked
               Width           =   240
            End
         End
         Begin VB.Frame Frame 
            Caption         =   "Window&Style"
            Height          =   1665
            Index           =   1
            Left            =   75
            TabIndex        =   2
            Top             =   525
            Width           =   5865
            Begin VB.PictureBox PictBox 
               BorderStyle     =   0  'None
               Height          =   1365
               Index           =   1
               Left            =   75
               ScaleHeight     =   91
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   381
               TabIndex        =   3
               TabStop         =   0   'False
               Top             =   225
               Width           =   5715
               Begin VB.OptionButton OptBtn 
                  Caption         =   "vbShowDefault"
                  Height          =   240
                  Index           =   10
                  Left            =   4080
                  TabIndex        =   14
                  Top             =   705
                  Width           =   1590
               End
               Begin VB.OptionButton OptBtn 
                  Caption         =   "vbRestore"
                  Height          =   240
                  Index           =   9
                  Left            =   4080
                  TabIndex        =   13
                  Top             =   390
                  Width           =   1590
               End
               Begin VB.OptionButton OptBtn 
                  Caption         =   "vbShowNA"
                  Height          =   240
                  Index           =   8
                  Left            =   4080
                  TabIndex        =   12
                  Top             =   75
                  Width           =   1590
               End
               Begin VB.OptionButton OptBtn 
                  Caption         =   "vbShowMinNoActive"
                  Height          =   240
                  Index           =   7
                  Left            =   2040
                  TabIndex        =   11
                  Top             =   1020
                  Width           =   2040
               End
               Begin VB.OptionButton OptBtn 
                  Caption         =   "vbMinimize"
                  Height          =   240
                  Index           =   6
                  Left            =   2040
                  TabIndex        =   10
                  Top             =   705
                  Width           =   2040
               End
               Begin VB.OptionButton OptBtn 
                  Caption         =   "vbShow"
                  Height          =   240
                  Index           =   5
                  Left            =   2040
                  TabIndex        =   9
                  Top             =   390
                  Width           =   2040
               End
               Begin VB.OptionButton OptBtn 
                  Caption         =   "vbShowNoActivate"
                  Height          =   240
                  Index           =   4
                  Left            =   2040
                  TabIndex        =   8
                  Top             =   75
                  Width           =   2040
               End
               Begin VB.OptionButton OptBtn 
                  Caption         =   "vbShowMaximized"
                  Height          =   240
                  Index           =   3
                  Left            =   75
                  TabIndex        =   7
                  Top             =   1020
                  Width           =   1890
               End
               Begin VB.OptionButton OptBtn 
                  Caption         =   "vbShowMinimized"
                  Height          =   240
                  Index           =   2
                  Left            =   75
                  TabIndex        =   6
                  Top             =   705
                  Width           =   1890
               End
               Begin VB.OptionButton OptBtn 
                  Caption         =   "vbShowNormal"
                  Height          =   240
                  Index           =   1
                  Left            =   75
                  TabIndex        =   5
                  Top             =   390
                  Width           =   1890
               End
               Begin VB.OptionButton OptBtn 
                  Caption         =   "vbHide"
                  Height          =   240
                  Index           =   0
                  Left            =   75
                  TabIndex        =   4
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
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   75
            TabIndex        =   1
            Top             =   75
            Width           =   8415
         End
         Begin VB.CheckBox ChkBox 
            Caption         =   "Now && &Timer"
            Height          =   915
            Index           =   1
            Left            =   6075
            TabIndex        =   18
            Top             =   1275
            Width           =   2490
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
    Rect       As RECT
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

Private intWindowStyle As Integer, TT_hWnd As Long, TI As TOOLINFO

Private Sub AddTool(ByVal ToolhWnd As Long, ByRef ToolTipText As String)
    TI.uId = ToolhWnd
    TI.lpszText = StrPtr(ToolTipText)
    SendMessageW TT_hWnd, TTM_ADDTOOLW, 0&, TI
End Sub

Private Sub ChkBox_Click(Index As Integer)
    Select Case Index
        Case 0: TxtBox(0).Enabled = Not TxtBox(0).Enabled
        Case 1: Timer1 = Not Timer1
    End Select
End Sub

Private Sub CmdBtn_Click(Index As Integer)
    On Error GoTo 1
    Select Case Index
        Case 0: MessageBoxA 0&, "Exit Code = " & Shell_n_Wait(ComBoX(0), intWindowStyle), App.Title, vbInformation
        Case 1: MessageBoxA 0&, "Return Value = " & ShellW(ComBoX(0), intWindowStyle, IIf(TxtBox(0).Enabled, CLng(TxtBox(0)), _
                            0&)) & vbCr & "Error &H" & Hex$(Err) & ": """ & Err.Description & """", App.Title, vbInformation
        Case 2: MessageBoxA 0&, "Exit Code = " & ShellWS(ComBoX(0), intWindowStyle, CBool(TxtBox(0)) _
                            And TxtBox(0).Enabled), App.Title, vbInformation
    End Select

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

    OptBtn(1) = True
    TxtBox(0) = "&HFFFFFFFF"

    ComBoX(0).AddItem "%ComSpec% /c Title Hi %UserName%!&Ping 127.0.0.1&Echo.&Pause&Exit 1000"
    ComBoX(0).AddItem "%SystemRoot%\explorer /e,/select,%Temp%"
    ComBoX(0).AddItem "rundll32 shell32.dll,Control_RunDLL inetcpl.cpl,@0,1"    'http://www.robvanderwoude.com/rundll.php#ControlPanelApplets
    ComBoX(0).AddItem """%ProgramFiles%\..\WINDOWS\system32\calc"""
    ComBoX(0).AddItem """%CommonProgramFiles%\..\..\WINDOWS\system32\winver"""
    ComBoX(0).AddItem """%ProgramFiles%\.\Common Files\..\..%HomePath%\..\..\WINDOWS\.\system32\..\system32\cmd"" /c start /wait rundll32 shell32.dll,Control_RunDLL inetcpl.cpl,@0,1"
    If Not InIDE Then ComBoX(0).AddItem "notepad " & GetShortPathName(App.Path & "\" & App.EXEName & ".exe")
    ComBoX(0).ListIndex = 0

    ComboEdit_hWnd = FindWindowExW(ComBoX(0).hWnd, , StrPtr("Edit"))
    SHAutoComplete ComboEdit_hWnd, SHACF_FILESYSTEM

    TT_hWnd = CreateWindowExW(lpClassName:=StrPtr(TOOLTIPS_CLASSW), _
                              dwStyle:=TTS_NOPREFIX, _
                              hWndParent:=hWnd, _
                              hInstance:=App.hInstance)
    If TT_hWnd Then
        SetTopMost TT_hWnd
        TI.cbSize = LenB(TI)
        TI.hWnd = hWnd
        TI.hInst = App.hInstance
        TI.uFlags = TTF_IDISHWND Or TTF_SUBCLASS
        SendMessageW TT_hWnd, TTM_SETMAXTIPWIDTH, 0&, 240&

        AddTool ComBoX(0).hWnd, "Enter path (quoted if with spaces) to executable or document file." & vbCrLf & "May include environment variables and/or arguments."
        AddTool ComboEdit_hWnd, "Enter path (quoted if with spaces) to executable or document file." & vbCrLf & "May include environment variables and/or arguments."
        AddTool TxtBox(0).hWnd, "For ShellW, type -1 or &HFFFFFFFF (INFINITE) to wait indefinitely." & vbCrLf & "For WScript, any non-zero value will do."
        AddTool ChkBox(1).hWnd, "Detect process termination by constant polling via a timer." & vbCrLf & "Send app into the background by minimizing."
        AddTool CmdBtn(0).hWnd, "Extends the native Shell function by waiting for the shelled" & vbCrLf & "program's termination without blocking other events."
        AddTool CmdBtn(1).hWnd, "Runs an executable program or document; optionally waits" & vbCrLf & "for a specified amount of time before resuming execution."
        AddTool CmdBtn(2).hWnd, "Runs a program in a new process."
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If TT_hWnd Then TT_hWnd = DestroyWindow(TT_hWnd): Debug.Assert TT_hWnd
    End    'Forcibly terminate process just in case code is stuck in the loops (doesn't seem to affect WScript.Shell...)
End Sub    'This arguably is a valid use of the End statement (",)

Private Function GetShortPathName(ByRef sLongPathName As String) As String
    SysReAllocStringLen VarPtr(GetShortPathName), , GetShortPathNameW(StrPtr(sLongPathName)) - 1&
    GetShortPathNameW StrPtr(sLongPathName), StrPtr(GetShortPathName), Len(GetShortPathName) + 1&
End Function

Private Function InIDE(Optional ByRef B As Boolean = True) As Boolean
    If B Then Debug.Assert Not InIDE(InIDE) Else B = True
End Function

Private Sub OptBtn_Click(Index As Integer)
    intWindowStyle = Index
End Sub

Private Sub Timer1_Timer()                         'The timer keeps messages coming, effectively polling the target
    ChkBox(1).Caption = Now & vbNewLine & Timer    'process regularly. Turn off to see if the process' termination
End Sub                                            'can still be accurately detected. Minimize the window too.

Private Sub TxtBox_GotFocus(Index As Integer)
    TxtBox(Index).SelStart = 0&
    TxtBox(Index).SelLength = Len(TxtBox(Index))
End Sub
