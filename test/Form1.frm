VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DL Test"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   235
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   329
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Pause"
      Height          =   735
      Left            =   3720
      TabIndex        =   15
      Top             =   2280
      Width           =   1215
   End
   Begin Project1.DownLoad dl1 
      Left            =   4440
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Resume"
      Enabled         =   0   'False
      Height          =   735
      Left            =   2520
      TabIndex        =   14
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   735
      Left            =   1320
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1920
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Download!"
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ColdFusion244@aol.com"
      Height          =   255
      Left            =   1320
      TabIndex        =   18
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "I can't get Resume to work correctly. If you can help, e-mail me."
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   4575
   End
   Begin VB.Label lblResume 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Resume Supported - "
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lblconnected 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connection Present - "
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblState 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connection - "
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblerr 
      AutoSize        =   -1  'True
      Caption         =   "Error Code - "
      Height          =   195
      Left            =   0
      TabIndex        =   11
      Top             =   1440
      Width           =   885
   End
   Begin VB.Label lblTS 
      AutoSize        =   -1  'True
      Caption         =   "Total Size - "
      Height          =   195
      Left            =   0
      TabIndex        =   10
      Top             =   960
      Width           =   840
   End
   Begin VB.Label lblRate 
      AutoSize        =   -1  'True
      Caption         =   "Transfer Rate - "
      Height          =   195
      Left            =   2280
      TabIndex        =   9
      Top             =   720
      Width           =   1110
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "Time Remaining - "
      Height          =   195
      Left            =   2280
      TabIndex        =   8
      Top             =   240
      Width           =   1275
   End
   Begin VB.Label lblexists 
      AutoSize        =   -1  'True
      Caption         =   "File Exists - "
      Height          =   195
      Left            =   2280
      TabIndex        =   6
      Top             =   0
      Width           =   825
   End
   Begin VB.Label lblpercent 
      AutoSize        =   -1  'True
      Caption         =   "Percent Downloaded - "
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   1635
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Status - "
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   585
   End
   Begin VB.Label lblTotalBytes 
      AutoSize        =   -1  'True
      Caption         =   "Total Bytes - "
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   930
   End
   Begin VB.Label lblCBytes 
      AutoSize        =   -1  'True
      Caption         =   "Bytes Read - "
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If dl1.CPause = True Or dl1.InDL = True Then Exit Sub
dl1.Url = "http://www.aol.com/"
'use the download below for a longer better test
'dl1.Url = "http://download.microsoft.com/download/vstudio60ent/SP5/Wideband-VB/WIN98Me/EN-US/vs6sp5vb.exe"
dl1.GetFileInformation
If dl1.FileSize <= 0 Then
Exit Sub
Else
lblconnected = "Connection Present - " & dl1.Connected
ProgressBar1.Max = dl1.FileSize
lblTotalBytes = "Total Bytes - " & dl1.FileSize
lblTS = "Total Size - " & dl1.FileSize
lblexists = "File Exists - " & dl1.FileExists
lblResume = "Resume Supported - " & dl1.AResume
dl1.DownLoad
End If
End Sub

Private Sub Command2_Click()
dl1.Cancel
End Sub

Private Sub Command3_Click()
dl1.DLResume
End Sub

Private Sub Command4_Click()
If dl1.CPause = True Then
dl1.Pause False
Command4.Caption = "Pause"
Else
dl1.Pause True
Command4.Caption = "Unpause"
End If
End Sub

Private Sub dl1_DLComplete()
MsgBox "Done"
End Sub

Private Sub dl1_DLError(lpErrorDescription As String)
MsgBox lpErrorDescription
End Sub

Private Sub dl1_StatusChange(lpStatus As String)
lblStatus = "Status - " & lpStatus
End Sub

Private Sub dl1_Percent(lPercent As Long)
lblpercent = "Percent Complete - " & lPercent
End Sub

Private Sub dl1_RecievedBytes(lnumBYTES As Long)
ProgressBar1.Value = lnumBYTES
lblCBytes = "Current Bytes Read - " & lnumBYTES
End Sub

Private Sub dl1_DLECode(lErrorCode As Long)

Select Case lErrorCode

Case 1
    lblerr = "Error Code - [1] Uknown Error!"
Case 2
    lblerr = "Error Code - [2] File Doesn't Exist!"
Case 3
    lblerr = "Error Code - [3] Server Timed Out!"
Case 4
    lblerr = "Error Code - [4] Download Was Cancelled By User!"
Case 5
    lblerr = "Error Code - [5] No Connection Present!"
Case 6
    lblerr = "Error Code - [6] Internal Server Error!"
Case 401
    lblerr = "Error Code - [401] Unauthorized Access!"
Case 403
    lblerr = "Error Code - [403] Access Denied!"
    
End Select

End Sub

Private Sub dl1_Rate(lpRate As String)
lblRate = "Transfer Rate - " & lpRate
End Sub

Private Sub dl1_TimeLeft(lpTime As String)
lblTime = "Time Remaining - " & lpTime
End Sub

Private Sub dl1_ConnectionState(strState As String)
lblState = "Connection - " & strState
End Sub
