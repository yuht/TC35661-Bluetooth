VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form_0xFD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "串口发送文件"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_0xFD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   12855
   StartUpPosition =   3  '窗口缺省
   Begin MSCommLib.MSComm MSComm1 
      Left            =   11925
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   0   'False
      RThreshold      =   1
      BaudRate        =   115200
      InputMode       =   1
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7080
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   810
      Width           =   12660
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2925
      TabIndex        =   9
      Text            =   "140"
      Top             =   360
      Width           =   1365
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   90
      TabIndex        =   7
      Top             =   8010
      Width           =   10770
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清空"
      Height          =   330
      Left            =   7695
      TabIndex        =   6
      Top             =   225
      Width           =   1590
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2.发送命令"
      Height          =   330
      Left            =   6030
      TabIndex        =   5
      Top             =   225
      Width           =   1590
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   11160
      Top             =   180
   End
   Begin VB.TextBox txt 
      Height          =   330
      Left            =   1620
      TabIndex        =   2
      Text            =   "115200"
      Top             =   330
      Width           =   1230
   End
   Begin VB.TextBox txtCOM10 
      Height          =   330
      Left            =   90
      TabIndex        =   1
      Text            =   "3"
      Top             =   330
      Width           =   1230
   End
   Begin VB.CommandButton Open 
      Caption         =   "1.打开串口和文件"
      Height          =   330
      Index           =   0
      Left            =   4365
      TabIndex        =   0
      Top             =   225
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "发送间隔 ms"
      Height          =   195
      Left            =   3015
      TabIndex        =   8
      Top             =   45
      Width           =   1140
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "波特率"
      Height          =   195
      Left            =   1845
      TabIndex        =   4
      Top             =   45
      Width           =   540
   End
   Begin VB.Label lblCom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "com口"
      Height          =   195
      Left            =   450
      TabIndex        =   3
      Top             =   45
      Width           =   465
   End
End
Attribute VB_Name = "Form_0xFD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function timeGetTime Lib "winmm.dll" () As Long
    
Dim GetPath As String, GetFile As String, GetFullFile As String
    

Function sendhex(str As String)

    Dim zzz() As String
    Dim i As Double
    zzz = Split(str, " ")
    ReDim kkk(UBound(zzz())) As Byte
    
    For i = 0 To UBound(zzz())
        kkk(i) = CByte("&h" & zzz(i))
        'Debug.Print Hex(zzz(i)),
    Next
    'Debug.Print
    Text1 = Text1 & ", 0x" & Right("0000" & Hex(UBound(zzz()) + 1), 4)
    MSComm1.Output = kkk
    
End Function

Private Sub Command1_Click()
    Dim Savetime As Double
    Dim a  As String
    Dim b As String
    

    
'    DotPos = InStrRev(GetFullFile, "\")
'    GetPath = Left(GetFullFile, DotPos)
'    GetFile = Replace(GetFullFile, GetPath, "")
'    Debug.Print GetFullFile
'    Debug.Print GetFile
'    Debug.Print GetPath
    
 
    
    Open GetFullFile For Input As #1
    
        Do While Not EOF(1)
            Line Input #1, a
            b = a
                Debug.Print a
                If InStr(a, "//") Then
                    a = Trim(Left(a, InStr(a, "//") - 1))
                End If
                Debug.Print a
'                Exit Sub
                If (Len(a)) Then
                    Text3 = "Com" & MSComm1.CommPort & "SEND: " & b & vbCrLf & Text3
                    sendhex (a)
                    Savetime = timeGetTime
                    Do
                        DoEvents
                        DoEvents
                        DoEvents
                        
                    Loop While timeGetTime < Savetime + CInt(Text2.Text)
                End If
        Loop
    Close #1
End Sub
 

Private Sub Command2_Click()
    Text3 = ""
End Sub

Private Sub MSComm1_OnComm()
    If MSComm1.CommEvent <> comEvReceive Then
        Exit Sub
    End If
    
    Timer2.Enabled = True
    
End Sub
Private Sub Open_Click(Index As Integer)

    On Error Resume Next
    MSComm1.CommPort = txtCOM10.Text
    MSComm1.Settings = txt.Text & ",n,8,1"
    
    Err.Clear
    MSComm1.PortOpen = True
    If Err Then
        MsgBox Err.Number & vbCrLf & Err.Description
        Err.Clear
        Exit Sub
    End If
    Text3 = "串口" & txtCOM10 & "已打开" & vbCrLf & Text3
    '--------------------------------------------

    
    
    GetFullFile = GetFileName(Me.hWnd)
    
    If Len(GetFullFile) = 0 Then
        Exit Sub
    End If
    Text3 = "文件已经打开，等待发送" & vbCrLf & Text3
     
End Sub

  
 

Private Sub Timer2_Timer()
    Timer2.Enabled = False
    
    Dim BytReceived() As Byte
    Dim zzz As String
  
    BytReceived = MSComm1.Input
    Dim i As Integer

    For i = 0 To UBound(BytReceived())
            zzz = zzz & " " & Right("00" & Hex(BytReceived(i)), 2)
    Next
    Text3 = "Com" & MSComm1.CommPort & " rec:" & zzz & vbCrLf & Text3
    
End Sub
