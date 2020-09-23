VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000016&
   Caption         =   "Inet Checker"
   ClientHeight    =   6510
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8190
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin Project1.lvButtons_H c4 
      Height          =   375
      Left            =   2160
      TabIndex        =   22
      Top             =   5880
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "Enable Watch"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H c3 
      Height          =   375
      Left            =   6240
      TabIndex        =   21
      Top             =   5880
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "Exit"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H c2 
      Height          =   375
      Left            =   4200
      TabIndex        =   20
      Top             =   5880
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "Clear Log"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H c1 
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   5880
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "Hide"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7200
      Top             =   3600
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabMaxWidth     =   3528
      BackColor       =   -2147483626
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "  Websites/Times  "
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lv"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "   Settings    "
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         Caption         =   "Settings And Options"
         ForeColor       =   &H00FF0000&
         Height          =   5055
         Left            =   -74880
         TabIndex        =   9
         Top             =   480
         Width           =   7695
         Begin Project1.lvButtons_H lvButtons_H3 
            Height          =   375
            Left            =   3840
            TabIndex        =   28
            Top             =   4440
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Caption         =   "Remove User"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin Project1.lvButtons_H lvButtons_H2 
            Height          =   375
            Left            =   2280
            TabIndex        =   27
            Top             =   4440
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Caption         =   "Add User"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3600
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   3960
            Width           =   3855
         End
         Begin Project1.lvButtons_H lvButtons_H1 
            Height          =   375
            Left            =   240
            TabIndex        =   24
            Top             =   4440
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Caption         =   "Change Password"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin Project1.lvButtons_H c5 
            Height          =   375
            Left            =   5400
            TabIndex        =   23
            Top             =   4440
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            Caption         =   "Add/Remove Website"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin VB.ComboBox Combo1 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   1800
            Width           =   2535
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Enable Time Limit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   1200
            Width           =   1695
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Create Log File"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   3120
            Width           =   1575
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Allowed Websites Only"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   2520
            Width           =   2055
         End
         Begin VB.ListBox List1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2985
            Left            =   3600
            TabIndex        =   12
            Top             =   600
            Width           =   3855
         End
         Begin VB.Label Label11 
            Caption         =   "User Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   26
            Top             =   3720
            Width           =   1815
         End
         Begin VB.Label Label10 
            Caption         =   "Allocate Time"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "Websites Allowed"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   13
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label8 
            Caption         =   "Your Home Page"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label7 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   600
            Width           =   2775
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Info"
         ForeColor       =   &H00FF0000&
         Height          =   975
         Left            =   120
         TabIndex        =   2
         Top             =   4560
         Width           =   7695
         Begin VB.Label Label6 
            Caption         =   "Time Spent There"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4080
            TabIndex        =   8
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label5 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4080
            TabIndex        =   7
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Website Visited Longest"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label3 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   3495
         End
         Begin VB.Label Label2 
            Caption         =   "Total Time On Net"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6000
            TabIndex        =   4
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6000
            TabIndex        =   3
            Top             =   480
            Width           =   1455
         End
      End
      Begin MSComctlLib.ListView lv 
         Height          =   3975
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   7011
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483646
         BackColor       =   -2147483634
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Url Address"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Time Entered"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Time Spent"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      X1              =   120
      X2              =   8040
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      WindowList      =   -1  'True
      Begin VB.Menu mnuhideit 
         Caption         =   "Hide"
      End
      Begin VB.Menu mnuenable 
         Caption         =   "Enable Watch"
      End
      Begin VB.Menu mnuclear 
         Caption         =   "Clear Log"
      End
      Begin VB.Menu mnu 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "Edit"
      Begin VB.Menu mnutimelimit 
         Caption         =   "Enable Time Limit"
      End
      Begin VB.Menu mnuallowed 
         Caption         =   "Allowed Websites Only"
      End
      Begin VB.Menu mnucreate 
         Caption         =   "Create Log File"
      End
   End
   Begin VB.Menu mnuhide 
      Caption         =   "hide"
      Visible         =   0   'False
      Begin VB.Menu mnushow 
         Caption         =   "Show Inet Checker"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const TH32CS_SNAPHEAPLIST = &H1
Const TH32CS_SNAPPROCESS = &H2
Const TH32CS_SNAPTHREAD = &H4
Const TH32CS_SNAPMODULE = &H8
Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Const TH32CS_INHERIT = &H80000000
Const MAX_PATH As Integer = 260
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Const REG_SZ = 1
Const REG_BINARY = 3
Const HKEY_CURRENT_USER = &H80000001
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private shomepage As String

Private Type NOTIFYICONDATA
cbSize As Long
hwnd As Long
uId As Long
uFlags As Long
uCallBackMessage As Long
hIcon As Long
szTip As String * 64
End Type

'declare constants (still in declarations)
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

Private Const WM_MOUSEMOVE = &H200

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

'The following constants are used to determine the mouse input for
'the icon in the taskbar

'Left-click constants.
Private Const WM_LBUTTONDBLCLK = &H203 'Double-click
Private Const WM_LBUTTONDOWN = &H201 'Button down
Private Const WM_LBUTTONUP = &H202 'Button up

'Right-click constants.
Private Const WM_RBUTTONDBLCLK = &H206 'Double-click
Private Const WM_RBUTTONDOWN = &H204 'Button down
Private Const WM_RBUTTONUP = &H205 'Button up

'Declare the API function call.
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'Dimension a variable as the user-defined data type.
Dim nid As NOTIFYICONDATA

Private altime As Double

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

 Const SW_SHOWNORMAL = 1
 Const LB_FINDSTRINGEXACT = &H1A2
 Private Declare Function SendMessageString Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
    'retrieve nformation about the key
    lResult = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal 0, lDataBufSize)
    If lResult = 0 Then
        If lValueType = REG_SZ Then
            'Create a buffer
            strBuf = String(lDataBufSize, Chr$(0))
            'retrieve the key's content
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then
                'Remove the unnecessary chr$(0)'s
                RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
            End If
        ElseIf lValueType = REG_BINARY Then
            Dim strData As Integer
            'retrieve the key's value
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, strData, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = strData
            End If
        End If
    End If
End Function

Function GetString(hKey As Long, strPath As String, strValue As String)
    Dim ret
    'Open the key
    RegOpenKey hKey, strPath, ret
    'Get the key's content
    GetString = RegQueryStringValue(ret, strValue)
    'Close the key
    RegCloseKey ret
End Function



Private Sub c1_Click()

Me.Hide
nid.cbSize = Len(nid)
nid.hwnd = Me.hwnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Me.Icon
nid.szTip = "Inet checker" & vbNullChar

Shell_NotifyIcon NIM_ADD, nid

End Sub

Private Sub c2_Click()
Dim I As Integer

If c4.Caption = "Disable Watch" Then
 MsgBox "Please Disable Watch", vbExclamation, "Watch Enabled"
Else
 If lv.ListItems.Count > 0 Then
  Open Left$(App.Path, 1) & ":\INET_LOG_FILE_" & Format(Date, "dd_mm_yyyy") & ".txt" For Append As #1
   Print #1,
   Print #1, "LOG CLEARED" & "  " & Now
   Print #1, "Websites/Details"; Tab(40); "Time Entered"; Tab(55); "Date"; Tab(68); "Time Spent"
   Print #1,
     For I = 1 To lv.ListItems.Count
         Print #1, lv.ListItems(I); Tab(40); lv.ListItems(I).SubItems(1); Tab(55); _
            lv.ListItems(I).SubItems(2); Tab(68); lv.ListItems(I).SubItems(3)
     Next I
   Print #1,
   Print #1, "Website Visited Longest"; Tab(30); Label3.Caption
   Print #1, "Time Spent There"; Tab(30); Label5.Caption
   Print #1,
   Print #1, "Total Time On Net"; Tab(30); Label1.Caption
 Close #1
lv.ListItems.Clear
End If

End If
 

End Sub

Private Sub c3_Click()

Unload Me
End Sub

Private Sub c4_Click()
Dim itm As ListItem

If c4.Caption = "Enable Watch" Then
 Set itm = lv.ListItems.Add(, , "Watch Enabled")
 itm.SubItems(1) = Time
 itm.SubItems(2) = Date
 mnuenable.Checked = True
 
 c4.Caption = "Disable Watch"
Else
 Set itm = lv.ListItems.Add(, , "Watch Disabled")
 itm.SubItems(1) = Time
 itm.SubItems(2) = Date
 mnuenable.Checked = False
 c4.Caption = "Enable Watch"
End If

End Sub

Private Sub c5_Click()
Dim I As Integer

Form4.Show vbModal

If regpas2 <> "" Then
 If regpas2 = regpas Then
  regpas2 = ""
  If Combo2.ListCount > 0 Then
   For I = 0 To Form1.Combo2.ListCount - 1
     Form3.Combo1.AddItem Form1.Combo2.List(I)
    Next I
    Form3.Combo1.ListIndex = 0
   Form3.Show vbModal
  Else
   MsgBox "You Need At Least One User To Add To" & vbCrLf & "Please Add A User", vbInformation, "Error"
  End If
Else
 MsgBox "Incorrect Password", vbExclamation, "Error"
 regpas2 = ""
End If
End If


End Sub

Private Sub Check1_Click()

 If Check1.Value = vbChecked Then
  mnuallowed.Checked = True
  If c4.Caption = "Disable Watch" Then
   MsgBox "Please Disable Watch While You Set User Name", vbInformation, "Error"
   Check1.Value = vbUnchecked
  ElseIf Combo2.ListCount = 0 Then
   MsgBox "You Need To Add A User To Use" & vbCrLf & "Allowed Websites Only", vbInformation, "Error"
   Check1.Value = vbUnchecked
  Else
   Combo2.Enabled = False
  End If
 Else
  mnuallowed.Checked = False
  Combo2.Enabled = True
 End If

End Sub

Private Sub Check2_Click()
 
If Check2.Value = vbChecked Then
 mnucreate.Checked = True
Else
 mnucreate.Checked = False
End If

End Sub

Private Sub Check3_Click()

If Check3.Value = vbChecked Then
 Combo1.Enabled = True
 mnutimelimit.Checked = True
Else
 Combo1.Enabled = False
 Combo1.ListIndex = 0
 mnutimelimit.Checked = False
End If

End Sub










Private Sub Combo1_Click()

Dim str1() As String

str1 = Split(Combo1.Text)

altime = CDbl(TimeSerial(Val(str1(0)), Val(str1(1)), 0))



End Sub




Private Sub Combo2_Click()
Dim lines As String, found As Boolean

List1.Clear

Open App.Path & "\users.dat" For Input As #1
 Do While Not EOF(1)
  Line Input #1, lines
   If found = True Then
    If Left$(lines, 5) <> "user_" Then
     List1.AddItem lines
    Else
     Exit Do
    End If
   Else
    If lines = "user_" & Combo2.List(Combo2.ListIndex) Then
     found = True
    End If
   End If
 Loop
Close #1

End Sub



Private Sub Form_Load()
Dim I As Integer, X As Integer

For I = 0 To 23
 For X = 0 To 50 Step 10
  Combo1.AddItem I & "Hrs " & X & "Mins"
 Next X
Next I
Combo1.RemoveItem 0
Combo1.AddItem "24Hrs " & "0Mins"
Combo1.ListIndex = 0

shomepage = GetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Start Page")
If shomepage = "" Then shomepage = "http://www.google.com/"
Label7.Caption = shomepage

regpas = GetSetting(Me.Name, "latest", "pass")
 If regpas = "" Then
  Form1.Show
  Form2.Show vbModal
 End If
If f2cancel = True Then
 Unload Me
End If

checkforfile

End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim msg As Long

msg = X / Screen.TwipsPerPixelX

Select Case msg
 Case WM_RBUTTONUP
 Me.PopupMenu mnuhide 'call the popup menu
End Select

End Sub


Private Sub Form_Unload(Cancel As Integer)
Shell_NotifyIcon NIM_DELETE, nid

If Check2.Value = vbChecked Then
   createlog
Else
 If lv.ListItems.Count > 0 Then
  If MsgBox("Would You Like To Create A Log File" & vbCrLf & "Before You Exit", vbYesNo, "Create Log") = vbYes Then
   createlog
  End If
 End If
End If

Set Form1 = Nothing

End Sub


Private Sub lvButtons_H1_Click()

Form2.Label1.Caption = "Enter Password"
Form2.Label2.Caption = "Enter Change"
Form2.Caption = "Change Password"
Form2.Height = 2670
Form2.lvButtons_H1.Top = 1680
Form2.lvButtons_H2.Top = 1680
Form2.Line1.Y1 = 1440
Form2.Line1.Y2 = 1440
Form2.Text3.Visible = True
Form2.Label3.Visible = True
Form2.Show vbModal

End Sub

Private Sub mnuabout_Click()

End Sub

Private Sub lvButtons_H2_Click()
Dim response As String

response = InputBox("Enter New User", "New User")

If response <> "" Then
  Open App.Path & "\users.dat" For Append As #1
   Print #1, "user_" & response
   Print #1, shomepage
  Close #1
  List1.Clear
 List1.AddItem shomepage
  Combo2.AddItem response
 Combo2.ListIndex = Combo2.ListCount - 1
Else
 If StrPtr(response) <> 0 Then
  MsgBox "Nothing Entered", vbInformation, "Error"
 End If
End If

End Sub

Private Sub lvButtons_H3_Click()

If Combo2.ListCount > 0 Then
 If Combo2.ListIndex >= 0 Then
  If MsgBox("Are You Sure You Want To Remove The User" & vbCrLf & """" & Combo2.Text & """", vbYesNo, "Removing User") = vbYes Then
   removeuser
   Combo2.RemoveItem Combo2.ListIndex
   List1.Clear
    If Combo2.ListCount > 0 Then
     Combo2.ListIndex = 0
    Else
     Kill App.Path & "\users.dat"
    End If
  End If
 Else
  MsgBox "Please Select An Item", vbInformation, "Error"
 End If
Else
 MsgBox "No Items To Remove", vbInformation, "Error"
End If

End Sub


Private Sub mnuallowed_Click()

If Check1.Value = vbChecked Then
 Check1.Value = vbUnchecked
Else
 Check1.Value = vbChecked
End If

End Sub

Private Sub mnuclear_Click()
c2_Click
End Sub

Private Sub mnucreate_Click()

If mnucreate.Checked = True Then
 Check2.Value = vbUnchecked
 mnucreate.Checked = False
Else
 Check2.Value = vbChecked
 mnucreate.Checked = True
End If

End Sub

Private Sub mnuenable_Click()

If mnuenable.Checked = True Then
 mnuenable.Checked = False
 c4_Click
Else
 mnuenable.Checked = True
 c4_Click
End If
 
 
End Sub

Private Sub mnuexit_Click()
Unload Me
End Sub

Private Sub mnuhideit_Click()
c1_Click
End Sub

Private Sub mnushow_Click()
Form4.Show vbModal

If regpas2 <> "" Then
If regpas2 = regpas Then
 regpas2 = ""
 Shell_NotifyIcon NIM_DELETE, nid
 Form1.Show
Else
 MsgBox "Incorrect Password", vbExclamation, "Error"
 regpas2 = ""
End If
End If

End Sub

Private Sub mnutimelimit_Click()

If mnutimelimit.Checked = True Then
 Check3.Value = vbUnchecked
 mnutimelimit.Checked = False
Else
 Check3.Value = vbChecked
 mnutimelimit.Checked = True
End If

End Sub

Private Sub Timer1_Timer()
Dim r As String, online As Boolean
Dim hSnapShot As Long, uProcess As PROCESSENTRY32
Dim itm As ListItem
Static check As Integer
Static secs As Long
Static total As Long

If c4.Caption = "Enable Watch" Then
 Form1.Caption = "Inet Checker     " & Format(Now, "dddd, dd mmmm yyyy    H:MM AMPM")

Else
Form1.Caption = "Inet Checker     " & Format(Now, "dddd, dd mmmm yyyy    H:MM AMPM")

hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    uProcess.dwSize = Len(uProcess)
    r = Process32First(hSnapShot, uProcess)
    
Do While r
    If Left$(LCase(uProcess.szExeFile), 12) = "iexplore.exe" Then
      online = True
      Exit Do
    End If
  r = Process32Next(hSnapShot, uProcess)
Loop
  CloseHandle hSnapShot
  

If online = True And check = 0 Then
  Set itm = lv.ListItems.Add(, , "Internet Opened")
   itm.SubItems(1) = Time
   itm.SubItems(2) = Date
   check = 1
  
   
ElseIf online = True And check = 1 Then
  EnumWindows AddressOf EnumProc, 0
   If Mid$(urladdress, 1, 4) <> "http" Then Exit Sub
   If Left$(urladdress, InStr(9, urladdress, "/")) <> lv.ListItems(lv.ListItems.Count) And urladdress <> "" Then
     Set itm = lv.ListItems.Add(, , Left$(urladdress, InStr(9, urladdress, "/")))
     itm.SubItems(1) = Time
     itm.SubItems(2) = Date
     secs = 0
     If Check1.Value = vbChecked Then
      allowedsites
     End If
    Else
     secs = secs + 1
     total = total + 1
     lv.ListItems(lv.ListItems.Count).SubItems(3) = Format(DateAdd("s", secs, "00:00:00"), "hh:mm:ss")
     Label1.Caption = Format(DateAdd("s", total, "00:00:00"), "hh:mm:ss")
     getmostused
    If Check3.Value = vbChecked Then
     timelimit
    End If
   End If
   
ElseIf online = False And check = 1 Then
  check = 0
  Set itm = lv.ListItems.Add(, , "Internet Closed")
   itm.SubItems(1) = Time
   itm.SubItems(2) = Date
End If
End If

End Sub



Private Sub getmostused()
Dim I As Integer, urlstring As String, date1 As Date

date1 = CDate("00:00:00")

For I = 1 To lv.ListItems.Count
 If lv.ListItems(I).SubItems(3) <> "" Then
  If CDate(lv.ListItems(I).SubItems(3)) > date1 Then
   date1 = CDate(lv.ListItems(I).SubItems(3))
   urlstring = lv.ListItems(I)
  End If
 End If
Next I

Label3.Caption = urlstring
Label5.Caption = date1


End Sub





Private Sub allowedsites()
Dim num As Long

num = SendMessageString(List1.hwnd, LB_FINDSTRINGEXACT, -1, ByVal lv.ListItems(lv.ListItems.Count))


If num = -1 Then
   ShellExecute Me.hwnd, vbNullString, Label7.Caption, vbNullString, "C:\", SW_SHOWNORMAL
   MsgBox "THE SITE VISITED WAS NOT ALLOWED" & vbCrLf & "SENT BACK TO HOMEPAGE", vbCritical, "NOT ALLOWED"
  lv.ListItems.Add , , "SITE NOT ALLOWED sent to homepage"
End If
 
End Sub

Private Sub createlog()
Dim I As Integer

If lv.ListItems.Count > 0 Then
Open Left$(App.Path, 1) & ":\INET_LOG_FILE_" & Format(Date, "dd_mm_yyyy") & ".txt" For Append As #1
  Print #1,
  Print #1, "Websites/Details"; Tab(40); "Time Entered"; Tab(55); "Date"; Tab(68); "Time Spent"
  Print #1,
 For I = 1 To lv.ListItems.Count
  Print #1, lv.ListItems(I); Tab(40); lv.ListItems(I).SubItems(1); Tab(55); _
            lv.ListItems(I).SubItems(2); Tab(68); lv.ListItems(I).SubItems(3)
 Next I
 Print #1,
 Print #1, "Website Visited Longest"; Tab(30); Label3.Caption
 Print #1, "Time Spent There"; Tab(30); Label5.Caption
 Print #1,
 Print #1, "Total Time On Net"; Tab(30); Label1.Caption
Close #1
End If
  
End Sub

Private Sub timelimit()

Dim fultime As Double

fultime = CDbl(CDate(Label1.Caption))

If fultime >= altime Then
 gostartpage
End If

End Sub

Private Sub gostartpage()


If lv.ListItems(lv.ListItems.Count) <> Label7.Caption Then

  ShellExecute Me.hwnd, vbNullString, Label7.Caption, vbNullString, "C:\", SW_SHOWNORMAL
  lv.ListItems.Add , , "TIME UP sent to homepage"
  
End If



End Sub

Private Sub checkforfile()
Dim lines As String, found As Integer

If Dir$(App.Path & "\users.dat") <> "" Then
Open App.Path & "\users.dat" For Input As #1
  Do While Not EOF(1)
  Line Input #1, lines
   If Left$(lines, 5) = "user_" Then
    Combo2.AddItem Mid$(lines, 6)
   End If
  Loop
 Close #1
 Combo2.ListIndex = 0
End If
    
End Sub

Private Sub removeuser()
Dim I As Integer, lines As String, found As Boolean

Open App.Path & "\users.dat" For Input As #1
Open App.Path & "\users2.dat" For Output As #2
 Do While Not EOF(1)
  Line Input #1, lines
    
    If found = False Then
      If lines <> "user_" & Combo2.List(Combo2.ListIndex) Then
       Print #2, lines
      Else
       found = True
      End If
    Else
     If Left$(lines, 5) = "user_" Then
      found = False
      Print #2, lines
     End If
    End If
    
 Loop
Close #1
Close #2

Kill App.Path & "\users.dat"
Name App.Path & "\users2.dat" As App.Path & "\users.dat"

End Sub
