VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "控制板卡测试程序"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9630
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   9630
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame3 
      Caption         =   "OUT[16....9,8...1]"
      Height          =   645
      Left            =   5175
      TabIndex        =   16
      Top             =   45
      Width           =   4380
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   240
         Index           =   7
         Left            =   135
         TabIndex        =   32
         Top             =   270
         Width           =   195
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   240
         Index           =   6
         Left            =   360
         TabIndex        =   31
         Top             =   270
         Width           =   195
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   240
         Index           =   5
         Left            =   585
         TabIndex        =   30
         Top             =   270
         Width           =   195
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   240
         Index           =   4
         Left            =   810
         TabIndex        =   29
         Top             =   270
         Width           =   195
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   240
         Index           =   3
         Left            =   1215
         TabIndex        =   28
         Top             =   270
         Width           =   195
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   240
         Index           =   2
         Left            =   1440
         TabIndex        =   27
         Top             =   270
         Width           =   195
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   240
         Index           =   1
         Left            =   1665
         TabIndex        =   26
         Top             =   270
         Width           =   195
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   240
         Index           =   0
         Left            =   1890
         TabIndex        =   25
         Top             =   270
         Width           =   195
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   240
         Index           =   7
         Left            =   2250
         TabIndex        =   24
         Top             =   270
         Width           =   195
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   240
         Index           =   6
         Left            =   2475
         TabIndex        =   23
         Top             =   270
         Width           =   195
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   240
         Index           =   5
         Left            =   2700
         TabIndex        =   22
         Top             =   270
         Width           =   195
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   240
         Index           =   4
         Left            =   2925
         TabIndex        =   21
         Top             =   270
         Width           =   195
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   240
         Index           =   3
         Left            =   3240
         TabIndex        =   20
         Top             =   270
         Width           =   195
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   240
         Index           =   2
         Left            =   3465
         TabIndex        =   19
         Top             =   270
         Width           =   195
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   240
         Index           =   1
         Left            =   3690
         TabIndex        =   18
         Top             =   270
         Width           =   195
      End
      Begin VB.CheckBox Check1 
         Height          =   240
         Index           =   0
         Left            =   3915
         TabIndex        =   17
         Top             =   270
         Width           =   195
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "控制"
      Height          =   2580
      Left            =   3375
      TabIndex        =   11
      Top             =   45
      Width           =   1725
      Begin VB.CommandButton Command_Green 
         Caption         =   "绿"
         Height          =   360
         Left            =   405
         TabIndex        =   15
         Top             =   315
         Width           =   990
      End
      Begin VB.CommandButton Command_RED 
         Caption         =   "红"
         Height          =   360
         Left            =   405
         TabIndex        =   14
         Top             =   855
         Width           =   990
      End
      Begin VB.CommandButton Command_Red_Green 
         Caption         =   "红绿"
         Height          =   360
         Left            =   405
         TabIndex        =   13
         Top             =   1440
         Width           =   990
      End
      Begin VB.CommandButton Command_OFF 
         Caption         =   "关闭"
         Height          =   360
         Left            =   405
         TabIndex        =   12
         Top             =   1980
         Width           =   990
      End
   End
   Begin MSWinsockLib.Winsock WinsockClient 
      Left            =   1350
      Top             =   45
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "主机"
      Height          =   2580
      Left            =   45
      TabIndex        =   3
      Top             =   45
      Width           =   3255
      Begin VB.CommandButton Command3 
         Caption         =   "断开"
         Height          =   360
         Left            =   1755
         TabIndex        =   10
         Top             =   1395
         Width           =   990
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1080
         TabIndex        =   8
         Text            =   "50000"
         Top             =   855
         Width           =   1680
      End
      Begin VB.CommandButton Command2 
         Caption         =   "连接"
         Height          =   360
         Left            =   540
         TabIndex        =   5
         Top             =   1395
         Width           =   990
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1080
         TabIndex        =   4
         Text            =   "192.168.0.88"
         Top             =   405
         Width           =   1680
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "状态:未连接"
         Height          =   195
         Left            =   630
         TabIndex        =   9
         Top             =   2070
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "端口:"
         Height          =   195
         Index           =   1
         Left            =   405
         TabIndex        =   7
         Top             =   900
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "主机名"
         Height          =   195
         Index           =   0
         Left            =   405
         TabIndex        =   6
         Top             =   450
         Width           =   540
      End
   End
   Begin VB.TextBox textget 
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   5175
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   765
      Width           =   4380
   End
   Begin VB.TextBox Textsend 
      Height          =   1140
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3465
      Width           =   4785
   End
   Begin VB.CommandButton Command1 
      Caption         =   "退出"
      Height          =   360
      Left            =   4860
      TabIndex        =   0
      Top             =   2790
      Width           =   990
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Command_close_Click()
    Dim sockByte(7) As Byte
    sockByte(0) = &H5A
    sockByte(1) = &H1
    sockByte(2) = &H2
    sockByte(3) = &H2
    sockByte(4) = &H0
    sockByte(5) = &H0
    sockByte(6) = &HB9
    sockByte(7) = &HB8
    
    WinsockClient.SendData sockByte
    
    
End Sub

Private Sub Check1_Click(Index As Integer)
    Dim i As Integer
    Dim B As Byte
    Dim C As Byte
    B = 0
    C = 0
    For i = 0 To 7
        If (Check1(i).Value = Checked) Then
            B = B Or 2 ^ i
        End If
        If (Check2(i).Value = Checked) Then
            C = C Or 2 ^ i
        End If
    Next
    
    Dim sockByte(7) As Byte
    Dim CRCByte(4) As Byte
    Dim crcRes As String
    Dim crclh(1) As Byte
'    Dim i As Integer
    sockByte(0) = &H5A
    sockByte(1) = &H1
    sockByte(2) = &H2
    sockByte(3) = &H2
    sockByte(4) = C
    sockByte(5) = B
    sockByte(6) = 0
    sockByte(7) = 0
    
    For i = 0 To 4
        CRCByte(i) = sockByte(i + 1)
    Next

    Call CRC16(CRCByte, crclh)
    
    sockByte(6) = crclh(0)
    sockByte(7) = crclh(1)
    
    Debug.Print Hex(crclh(0))
    Debug.Print Hex(crclh(1))
    
    Call WinsockClient.SendData(sockByte)
    
End Sub

Private Sub Check2_Click(Index As Integer)
    Call Check1_Click(Index)
End Sub

Private Sub Command_Green_Click()
    Dim sockByte(7) As Byte
    sockByte(0) = &H5A
    sockByte(1) = &H1
    sockByte(2) = &H2
    sockByte(3) = &H2
    sockByte(4) = &H0
    sockByte(5) = &H1
    sockByte(6) = &H78
    sockByte(7) = &H78
    
    WinsockClient.SendData sockByte
    
End Sub

Private Sub Command_OFF_Click()
    Dim sockByte(7) As Byte
    sockByte(0) = &H5A
    sockByte(1) = &H1
    sockByte(2) = &H2
    sockByte(3) = &H2
    sockByte(4) = &H0
    sockByte(5) = &H0
    sockByte(6) = &HB9
    sockByte(7) = &HB8
    
    WinsockClient.SendData sockByte
End Sub

Private Sub Command_RED_Click()
    Dim sockByte(7) As Byte
    sockByte(0) = &H5A
    sockByte(1) = &H1
    sockByte(2) = &H2
    sockByte(3) = &H2
    sockByte(4) = &H0
    sockByte(5) = &H2
    sockByte(6) = &H38
    sockByte(7) = &H79
    
    WinsockClient.SendData sockByte
End Sub

Private Sub Command_Red_Green_Click()
    Dim sockByte(7) As Byte
    sockByte(0) = &H5A
    sockByte(1) = &H1
    sockByte(2) = &H2
    sockByte(3) = &H2
    sockByte(4) = &H0
    sockByte(5) = &H3
    sockByte(6) = &HF9
    sockByte(7) = &HB9
    
    WinsockClient.SendData sockByte
End Sub

'客户机程序使用的控件如下:
'（1）Command1：退出按钮；
'（2）Command2：连接按钮；
'（3）Winsockclient：客户Winsock；
'（4）Text1：主机名文本框；
'（5）Textsend：发送数据文本框；
'（6）Textget：接收数据文本框；
'客户机程序的源代码如下:
Private Sub Command1_Click()
    End
End Sub
Private Sub Command2_Click()
    WinsockClient.Close
    WinsockClient.RemotePort = CDbl(Text2.Text)
    WinsockClient.RemoteHost = Text1.Text
    WinsockClient.Connect
End Sub

Private Sub Command3_Click()
    Call Winsockclient_Close
End Sub

Private Sub Form_Load()
'    Textsend.Visible = False
'    textget.Visible = False
'    Winsockclient.RemotePort = 1001
'    Winsockclient.RemoteHost = "192.168.1.131"
    Frame2.Enabled = False
    Frame3.Enabled = False
End Sub
'Private Sub Text1_Change()
'    Winsockclient.RemoteHost = Text1.Text
'End Sub

Private Sub Winsockclient_Close()
    WinsockClient.Close
    Label2.Caption = "状态:已断开"
    Frame2.Enabled = False
    Frame3.Enabled = False
'    End
End Sub
Private Sub winsockclient_Connect()
    Label2.Caption = "状态:已连接"
    Textsend.Visible = True
    textget.Visible = True
    Frame2.Enabled = True
    Frame3.Enabled = True
'    Command2.Visible = False
End Sub
Private Sub winsockclient_DataArrival(ByVal bytesTotal As Long)
    Dim RecData() As Byte
    Dim tmpstr As String
    Dim i As Integer
    Call WinsockClient.GetData(RecData)
    
    For i = 0 To UBound(RecData)
        tmpstr = tmpstr & " " & Right("00" & Hex(RecData(i)), 2)
    Next
    
    textget.Text = tmpstr & vbCrLf & textget.Text
    
End Sub

