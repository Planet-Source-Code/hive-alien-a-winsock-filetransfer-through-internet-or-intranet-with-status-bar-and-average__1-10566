VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_Client 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Client"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Btn_Listen 
      Caption         =   "Listen for file"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   5160
      TabIndex        =   12
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Frame Fra_Advanced 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Advanced Settings"
      Height          =   2295
      Left            =   4680
      TabIndex        =   8
      Top             =   600
      Width           =   1935
      Begin VB.TextBox Txt_CurrentIP 
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox Txt_Port 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Text            =   "0"
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lbl_ExternalIP 
         BackStyle       =   0  'Transparent
         Caption         =   "External IP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Lbl_Port 
         BackStyle       =   0  'Transparent
         Caption         =   "Port to listen to:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Lbl_Info 
         BackStyle       =   0  'Transparent
         Caption         =   "(0 = free port, the port has to be the same in the server form)"
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1695
      End
   End
   Begin VB.Frame FraServer 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Client Settings"
      ForeColor       =   &H00000000&
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4455
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   2040
         Top             =   1080
      End
      Begin VB.TextBox Txt_File 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Text            =   "C:\"
         Top             =   600
         Width           =   4215
      End
      Begin MSComctlLib.ProgressBar FileBar 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1920
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComDlg.CommonDialog Dlg_Browser 
         Left            =   3840
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Lbl_Complete 
         BackStyle       =   0  'Transparent
         Caption         =   "Complete: 0%"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Save File to:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Lbl_FileSize 
         BackStyle       =   0  'Transparent
         Caption         =   "Filesize: 0 kb"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label Lbl_FileName 
         BackStyle       =   0  'Transparent
         Caption         =   "Filename: -"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   4095
      End
      Begin VB.Label Lbl_Averages 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "Average: 0 / KBps"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1800
         TabIndex        =   3
         Top             =   1680
         Width           =   2415
      End
   End
   Begin MSWinsockLib.Winsock Winsock_Receive 
      Left            =   1680
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "made by www.Inter-Dev.de"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   6615
   End
   Begin VB.Label Lbl_Status 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "winsock State"
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   3120
      Width           =   2415
   End
End
Attribute VB_Name = "Frm_Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DoneBytes As Long                   '# for calculating kbps
Dim DownloadingFile As Integer          '# freefile for open files



' Ronny R. Germany Berlin
' Contact me: manager@directbox.com
' Sorry for bad english... i'm german, but I become better... I swear
'
' I made this because a lot of examples on PSC are not precise enought
' There are a lot of unforseen errors
' I hope you enjoy!





Public Function GetField(Field As String, FieldPos As Long) As String

'# That 's an routine to get elements from a string


Dim FieldCounter As Long
Dim IPPositionStart As Long
Dim IPPositionEnde As Long
Dim TempPosition As Long
Dim OpenedID As String
    
    TempPosition = 1
    
    For FieldCounter = 1 To FieldPos - 1 Step 1
        IPPositionStart = InStr(TempPosition, Field, "|", vbTextCompare)
        TempPosition = IPPositionStart + 1
    Next FieldCounter
    IPPositionStart = IPPositionStart + 1
    IPPositionEnde = InStr(IPPositionStart, Field, "|", vbTextCompare)
On Error Resume Next
    If IPPositionEnde >= IPPositionStart Then
        GetField = Mid(Field, IPPositionStart, IPPositionEnde - IPPositionStart)
    End If

End Function

Private Sub Btn_Listen_Click()
On Error GoTo ErrorHandler:
        
        
        'the following routines are nessessary to beware of errors
        If Winsock_Receive.State <> sckClosed Then          '# Reset if winsock was in use
            Winsock_Receive.Close
        End If
        Winsock_Receive.Protocol = sckTCPProtocol           '# We work with TCP now
        '# Init the Winsock
        If Txt_Port.Text <> 0 Then                          '# select the port you entered
                Winsock_Receive.LocalPort = Txt_Port.Text   '# set the winsock receive port to the selected one
        Else
                Winsock_Receive.LocalPort = 0               '# in that case 0 means to select a free port
        End If
        Winsock_Receive.Listen                              '# listning on port selected above and current external and internal IP
        Lbl_Status.Caption = Winsock_Receive.State & " on port: " & Winsock_Receive.LocalPort
        
Exit Sub
        
ErrorHandler:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Command1_Click()
    End
End Sub

Private Sub Form_Load()
    Txt_CurrentIP.Text = CurrentIP(True)                    '# I didn't write the module but its nessessary when you work with winsock!
End Sub

Private Sub Label2_Click()
OpenUrl ("http://www.inter-dev.de")
End Sub

Private Sub Timer1_Timer()
'# Hehe Tricky...set a global variable to count up the bytes and calculate them to KBps every second
    Lbl_Averages.Caption = "Average: " & Format(DoneBytes / 1000, "###0.0") & " / KBps"
    DoneBytes = 0
End Sub

Private Sub Winsock_Receive_Close()
    Close #DownloadingFile                                  '# File Ready
    Winsock_Receive.Close                                   '# Close the winsock, for receiving next files?!
End Sub

Private Sub Winsock_Receive_ConnectionRequest(ByVal requestID As Long)

    '# accept the connections
    If Winsock_Receive.State <> sckClosed Then
        Winsock_Receive.Close
    End If
    Winsock_Receive.Accept requestID
    
    
    '# We use the close event to close the file afterwards
    
        
End Sub

Private Sub Winsock_Receive_DataArrival(ByVal bytesTotal As Long)
    Dim StrData As String
    Dim lNewValue As Long
    Dim Info As String
    Dim Glob_FileName As String
    
    StrData = ""                                    '# You only get filedata trought that winsock
                                                    ' so you only have to write it in the file opened before
    Winsock_Receive.GetData StrData, vbString
    
    
    '# Thats some file info send before we receive the first package
    Info = Left(StrData, 8)
    If Info = "FILEINFO" Then
        FileBar.Max = GetField(StrData, 2)
        Glob_FileName = GetField(StrData, 3)
        
        Txt_File.Text = Glob_FileName
        DownloadingFile = FreeFile
        Open App.Path & "\" & Glob_FileName For Binary Access Write As #DownloadingFile

        Exit Sub
    End If

    FileBar.Value = FileBar.Value + bytesTotal
    DoneBytes = DoneBytes + bytesTotal
    Lbl_Complete.Caption = "Complete: " & Int(100 / FileBar.Max * FileBar.Value) & " %"
    
    Put #DownloadingFile, , StrData
    DoEvents
    
    Debug.Print Len(StrData)
    
End Sub
