VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSComm32.Ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer minuta 
      Interval        =   60000
      Left            =   3960
      Top             =   2520
   End
   Begin VB.Timer pinger 
      Interval        =   2000
      Left            =   3960
      Top             =   1920
   End
   Begin VB.Timer powerCOM 
      Enabled         =   0   'False
      Interval        =   360
      Left            =   3960
      Top             =   1200
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3960
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   480
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ip = "192.168.1.1"
Const interval1 = 3
Const interval0 = 1
Dim ping_time, pinger_time
Private Type WSAdata
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 255) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type
Private Type Hostent
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
Private Type IP_OPTION_INFORMATION
    TTL As Byte
    Tos As Byte
    flags As Byte
    OptionsSize As Long
    OptionsData As String * 128
End Type
Private Type IP_ECHO_REPLY
    ADDRESS(0 To 3) As Byte
    Status As Long
    RoundTripTime As Long
    DataSize As Integer
    Reserved As Integer
    data As Long
    Options As IP_OPTION_INFORMATION
End Type
Private Declare Function GetHostByName Lib "wsock32.dll" Alias "gethostbyname" (ByVal HostName As String) As Long
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired&, lpWSAdata As WSAdata) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal handle As Long) As Boolean
Private Declare Function IcmpSendEcho Lib "ICMP" (ByVal IcmpHandle As Long, ByVal DestAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptns As IP_OPTION_INFORMATION, ReplyBuffer As IP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Boolean
Dim PingCheck

Private Sub Form_Load()
powerCOM.Interval = 100
ping_time = 1
If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
MSComm1.Settings = "38400,n,8,1"
MSComm1.CommPort = 1
MSComm1.PortOpen = Not MSComm1.PortOpen
MSComm1.PortOpen = False
End Sub

Private Sub minuta_Timer()
pinger.Enabled = False
pinger_time = pinger_time + 1
If pinger_time >= ping_time Then
pinger.Enabled = True
pinger_time = 0
End If
End Sub

Private Sub pinger_Timer()
Ping (ip)
If PingCheck = 0 Then
powerCOM.Enabled = True
ping_time = interval1
Else: ping_time = interval0
End If
Label1 = Time
pinger.Enabled = False
End Sub

Private Sub powerCOM_Timer()
If MSComm1.PortOpen Then
MSComm1.PortOpen = False
powerCOM.Enabled = False
powerCOM.Interval = 100
Else
MSComm1.PortOpen = True
powerCOM.Interval = 36000
End If
End Sub
'''''''''''''''
Private Sub Ping(ip)
Dim cnt As Boolean
    Dim hFile As Long
    Dim lpWSAdata As WSAdata
    Dim hHostent As Hostent, AddrList As Long
    Dim ADDRESS As Long, rIP As String
    Dim OptInfo As IP_OPTION_INFORMATION
    Dim EchoReply As IP_ECHO_REPLY
    Call WSAStartup(&H101, lpWSAdata)
cnt = False
   PingCheck = 0
    If GetHostByName(ip + String(64 - Len(ip), 0)) <> SOCKET_ERROR Then
        CopyMemory hHostent.h_name, ByVal GetHostByName(ip + String(64 - Len(ip), 0)), Len(hHostent)
        CopyMemory AddrList, ByVal hHostent.h_addr_list, 4
        CopyMemory ADDRESS, ByVal AddrList, 4
    End If
    hFile = IcmpCreateFile()
    OptInfo.TTL = 255
    If IcmpSendEcho(hFile, ADDRESS, String(32, "A"), 32, OptInfo, EchoReply, Len(EchoReply) + 8, 1) Then
        cnt = True
    End If
    Call IcmpCloseHandle(hFile)
'Next
    Call WSACleanup
If cnt = True Then
PingCheck = 1
'MsgBox IP & " - пинг проходит."
Else
PingCheck = 0
'MsgBox IP & " - узел не отвечает"
End If
End Sub
