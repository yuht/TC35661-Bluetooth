VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form Form_0xFD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "串口发送文件"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12780
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
   ScaleHeight     =   8670
   ScaleWidth      =   12780
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "2.打开初始化文件"
      Height          =   330
      Left            =   6345
      TabIndex        =   13
      Top             =   315
      Width           =   1590
   End
   Begin VB.TextBox txtText4 
      Height          =   1590
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   810
      Width           =   11985
   End
   Begin VB.CommandButton cmd 
      Caption         =   "发送单条命令"
      Height          =   1440
      Left            =   12240
      TabIndex        =   11
      Top             =   855
      Width           =   420
   End
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
      Height          =   5505
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   2520
      Width           =   12030
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   2925
      TabIndex        =   9
      Text            =   "140"
      Top             =   315
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      Height          =   420
      Left            =   45
      TabIndex        =   7
      Top             =   8100
      Width           =   12660
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清空接收区"
      Height          =   4560
      Left            =   12240
      TabIndex        =   6
      Top             =   3015
      Width           =   420
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3.发送初始化文件"
      Height          =   330
      Left            =   8280
      TabIndex        =   5
      Top             =   315
      Width           =   1995
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   11160
      Top             =   180
   End
   Begin VB.TextBox txt 
      Height          =   330
      Left            =   1507
      TabIndex        =   2
      Text            =   "115200"
      Top             =   315
      Width           =   1230
   End
   Begin VB.TextBox txtCOM10 
      Height          =   330
      Left            =   90
      TabIndex        =   1
      Text            =   "3"
      Top             =   315
      Width           =   1230
   End
   Begin VB.CommandButton Open 
      Caption         =   "1.打开串口"
      Height          =   330
      Index           =   0
      Left            =   4410
      TabIndex        =   0
      Top             =   315
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
    

Function ClearUnHex(str As String) As String
    Dim i As Integer
    Dim j As Integer
    Dim charx As String
    Do
        str = Replace(str, "  ", " ")
    Loop While InStr(1, str, "  ")
    
    i = Len(str)
    For j = 1 To Len(str)
        charx = Mid(str, j, 1)
        If charx <> " " Then
            If (charx < "0" Or charx > "9") Then
                If (UCase(charx) < "A" Or UCase(charx) > "F") Then
                    Exit For
                End If
            End If
        End If
    Next
    
    ClearUnHex = Trim(Left(str, j - 1))
    

End Function



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

Private Sub cmd_Click()
    sendd (txtText4)
End Sub

Function sendd(str As String)
    Dim Savetime As Double
    Dim a  As String

    
    a = ClearUnHex(str)
    Debug.Print a
    If (Len(a)) Then
        Text3 = "Com" & MSComm1.CommPort & "SEND: " & a & Mid(InStr(str, "/"), Len(str) - InStr(str, "/")) & vbCrLf & Text3
        sendhex (a)
        Savetime = timeGetTime
        Do
            DoEvents
            DoEvents
            DoEvents
            
        Loop While timeGetTime < Savetime + CInt(Text2.Text)
    End If
End Function

Private Sub Command1_Click()

    Dim b As String
    
'    DotPos = InStrRev(GetFullFile, "\")
'    GetPath = Left(GetFullFile, DotPos)
'    GetFile = Replace(GetFullFile, GetPath, "")
'    Debug.Print GetFullFile
'    Debug.Print GetFile
'    Debug.Print GetPath
    
 
    
    Open GetFullFile For Input As #1
    
        Do While Not EOF(1)
            Line Input #1, b
            sendd (b)
        Loop
    Close #1
End Sub
 

Private Sub Command2_Click()
    Text3 = ""
End Sub

Private Sub Command3_Click()
    GetFullFile = GetFileName(Me.hWnd)
    
    If Len(GetFullFile) = 0 Then
        Exit Sub
    End If
    Text3 = "文件已经打开，等待发送" & vbCrLf & Text3
End Sub

Private Sub MSComm1_OnComm()
    If MSComm1.CommEvent <> comEvReceive Then
        Exit Sub
    End If
    
    Timer2.Enabled = True
    
End Sub
Private Sub Open_Click(Index As Integer)

    On Error Resume Next
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If
    MSComm1.CommPort = txtCOM10.Text
    MSComm1.Settings = txt.Text & ",n,8,1"
    Err.Clear
    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
    End If
    If Err Then
        MsgBox Err.Number & vbCrLf & Err.Description
        Err.Clear
        Exit Sub
    End If
    Text3 = "串口" & txtCOM10 & "已打开" & vbCrLf & Text3
    '--------------------------------------------


     
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
