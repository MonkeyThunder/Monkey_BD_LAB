VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   BorderStyle     =   0  '����
   ClientHeight    =   5655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "���� ���"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   6855
   StartUpPosition =   1  '������ ���
   Begin VB.PictureBox P1 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  '����
      Height          =   5130
      Left            =   45
      ScaleHeight     =   5130
      ScaleWidth      =   6765
      TabIndex        =   2
      Top             =   495
      Width           =   6765
      Begin VB.TextBox txtID 
         Appearance      =   0  '���
         Height          =   345
         Left            =   1560
         TabIndex        =   7
         Top             =   90
         Width           =   5055
      End
      Begin VB.ComboBox Ser 
         Height          =   345
         Left            =   120
         TabIndex        =   5
         Text            =   "��ü"
         Top             =   90
         Width           =   1335
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '���
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   320
         Left            =   3465
         ScaleHeight     =   285
         ScaleWidth      =   3120
         TabIndex        =   3
         Top             =   4755
         Width           =   3150
         Begin SHDocVwCtl.WebBrowser WB 
            Height          =   15000
            Left            =   -120
            TabIndex        =   4
            Top             =   -120
            Width           =   15000
            ExtentX         =   26458
            ExtentY         =   26458
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
      End
      Begin MSComctlLib.ListView Lv 
         Height          =   4095
         Left            =   120
         TabIndex        =   6
         Top             =   570
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   7223
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ĳ����"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "������"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "����"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "����"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "url"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  '����
         Caption         =   "�뷡 ���� : ��������Ǻ�α�"
         Height          =   255
         Left            =   135
         TabIndex        =   8
         Top             =   4800
         Width           =   3255
      End
   End
   Begin VB.PictureBox P2 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  '����
      Height          =   5100
      Left            =   45
      ScaleHeight     =   5100
      ScaleWidth      =   6735
      TabIndex        =   12
      Top             =   495
      Width           =   6735
      Begin VB.ComboBox cBoard 
         Height          =   345
         Left            =   120
         TabIndex        =   13
         Text            =   "���� ���Ͽ�"
         Top             =   120
         Width           =   1455
      End
      Begin MSComctlLib.ListView LV2 
         Height          =   4335
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   7646
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   5997
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "�۾���"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "�����"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "��ȸ��"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "url"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.PictureBox P 
      Height          =   375
      Left            =   4680
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox Picture12 
      Appearance      =   0  '���
      BackColor       =   &H00FFA244&
      BorderStyle     =   0  '����
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   5940
      Picture         =   "Form1.frx":0B3A
      ScaleHeight     =   300
      ScaleWidth      =   360
      TabIndex        =   15
      Top             =   -15
      Width           =   355
   End
   Begin VB.PictureBox Label1 
      BackColor       =   &H00FFA244&
      BorderStyle     =   0  '����
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   3855
      TabIndex        =   9
      Top             =   0
      Width           =   3855
      Begin VB.Label lb2 
         BackColor       =   &H8000000E&
         BackStyle       =   0  '����
         Caption         =   "�κ�"
         BeginProperty Font 
            Name            =   "���� ���"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lb1 
         BackColor       =   &H8000000E&
         BackStyle       =   0  '����
         Caption         =   "�����"
         BeginProperty Font 
            Name            =   "���� ���"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   615
      End
   End
   Begin Project1.Skin1 Skin11 
      Height          =   5655
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   9975
   End
   Begin VB.Label Label12 
      Appearance      =   0  '���
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   9600
      TabIndex        =   0
      Top             =   6960
      Width           =   6735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim W As New WinHttpRequest

Private Declare Function SetWindowPos Lib "user32" _
(ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Type NOTIFYICONDATA
  cbSize As Long
  hWnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeoutOrVersion As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
End Type
  
Private Const NIIF_WARNING = 2
Private Const NIIF_ERROR = 3
Private Const NIIF_INFO = 1
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim SysTrayT As NOTIFYICONDATA




Private Sub cBoard_Click()
Select Case cBoard.Text
Case "���� ���Ͽ�"
�������Ͽ� LV2
Case "������"
���� LV2, 3588
Case "������"
���� LV2, 3589
Case "�Ҽ���"
���� LV2, 3591
Case "���̾�Ʈ"
���� LV2, 3590
Case "�ݼ���"
���� LV2, 4167
Case Else
MsgBox "�߸��� �����Դϴ�.", vbInformation, ""
End Select
LV2.SetFocus
End Sub

Private Sub Form_Initialize()
    Me.Width = 6855
    Me.Height = 5655
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim lngReturnValue As Long
 If Button = 1 Then
  Call ReleaseCapture
  lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
 End If
End Sub
Sub GetID()
    Dim RE$, Temp$, Total$, Url$, Nick$, Nick2$, Class$, Server$, Guild$
    Dim i&
    Dim Con$
    Dim sServer As Integer
    
    Select Case Ser
        Case "���ö�"
        sServer = 21
        Case "���̵�"
        sServer = 11
        Case "Į���"
        sServer = 31
        Case "�޵��"
        sServer = 41
        Case "�÷θ�"
        sServer = 51
        Case "�߷��þ�"
        sServer = 61
        Case "���丮��"
        sServer = 71
        Case "�������"
        sServer = 81
        Case Else
        sServer = 0
    End Select
    
    Con = txtID
    Con = URLEncodeUTF8(txtID)
    Con = Base64Encode(Con)
    With W
        .Open "GET", "http://black.game.daum.net/black/world/friend/search.daum?searchServer=" & sServer & "&searchEncName=" & Con & "&findName=" & Change(txtID) & "&page=1"
        .Send
        RE = .ResponseText
        Total = Split(Split(RE, "txt_item"">")(3), "<")(0)
        If Total > 15 Then: Total = 15
        Lv.ListItems.Clear
        For i = 1 To Total
            DoEvents
            Temp = Split(Split(RE, "<div class=""append_info"">")(i), "</li>")(0)
            Url = "http://black.game.daum.net/" & Split(Split(Temp, "</span><em><a href=""")(1), """")(0)
            Nick = Split(Split(Split(Temp, "data-userName=""")(1), "</a>")(0), ">")(1)
            Nick2 = Split(Split(Nick, "(")(1), ")")(0)
            Nick = Split(Nick, "(")(0)
            Class = Split(Split(Temp, "class=""txt_emph"">")(1), "<")(0)
            'Server = Split(Split(Temp, "class=""txt_emph"">")(2), "<")(0)
            Guild = Split(Split(Temp, "���Ա�� : </span><span class=""txt_g"">")(1), "</a>")(0)
            
            If InStr(Guild, "href") Then
                Guild = Split(Guild, ">")(1)
            Else
                Guild = Split(Guild, "<")(0)
            End If
            Guild = Replace$(Guild, " ", vbNullString)
            
            Lv.ListItems.Add , , Nick
            Lv.ListItems.Item(Lv.ListItems.Count).SubItems(1) = Nick2
            Lv.ListItems.Item(Lv.ListItems.Count).SubItems(2) = Class
            Lv.ListItems.Item(Lv.ListItems.Count).SubItems(3) = Guild
            Lv.ListItems.Item(Lv.ListItems.Count).SubItems(4) = Url
        Next i
    End With
End Sub
Sub �ʱ⼼��()
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, &H1 Or &H2
    Ser.AddItem "��ü"
    Ser.AddItem "���ö�"
    Ser.AddItem "���̵�"
    Ser.AddItem "Į���"
    Ser.AddItem "�޵��"
    Ser.AddItem "�÷θ�"
    Ser.AddItem "�߷��þ�"
    Ser.AddItem "���丮��"
    Ser.AddItem "�������"
    
    cBoard.AddItem "���� ���Ͽ�"
    cBoard.AddItem "������"
    cBoard.AddItem "������"
    cBoard.AddItem "�Ҽ���"
    cBoard.AddItem "���̾�Ʈ"
    cBoard.AddItem "�ݼ���"
    
    WB.Navigate2 "http://blog.naver.com/blogdj"
End Sub
Private Sub Form_Load()
    �ʱ⼼��
End Sub
Private Sub Lb2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lb2.ForeColor = &H8000000D
End Sub
Private Sub Lb2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lb2.ForeColor = &H80000012
    If lb2.Caption = "�˻�" Then
        lb2.Caption = "�κ�"
        P1.Visible = True
        P2.Visible = False
    Else
        lb2.Caption = "�˻�"
        P1.Visible = False
        P2.Visible = True
    End If
End Sub
Private Sub Lb1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lb1.ForeColor = &H8000000D
End Sub
Private Sub Lb1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lb1.ForeColor = &H80000012
    If lb1.Caption = "�����" Then
        lb1.Caption = "��ġ��"
        Me.Width = 795
        Me.Height = 465
    Else
        lb1.Caption = "�����"
        Me.Width = 6855
        Me.Height = 5655
    End If
End Sub
Private Sub Lv_DblClick()
On Error Resume Next
Dim Level$
W.Open "GET", Lv.ListItems.Item(Lv.SelectedItem.Index).SubItems(4)
W.Send
MsgBox Split(Split(W.ResponseText, "<span class=""txt_level"">")(1), "<")(0), vbInformation, ""
End Sub



Private Sub LV2_DblClick()
Shell Environ("windir") & "\explorer.exe """ & LV2.ListItems.Item(LV2.SelectedItem.Index).SubItems(4) & """", vbNormalFocus
End Sub

Private Sub P_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Static rec As Boolean, msg As Long
    msg = X / Screen.TwipsPerPixelX
    If rec = False Then
        rec = True
        Select Case msg
            Case &H202: Form1.Show '���ʸ��콺 Ŭ���ϸ� �߻��ϴ� �̺�Ʈ
        End Select
        rec = False
    End If
End Sub
Private Sub Picture12_Click()
Me.Hide
    With SysTrayT
        .cbSize = Len(SysTrayT)
        .hWnd = P.hWnd
        .uID = 1
        .uFlags = &H2 Or &H1 Or &H10 Or &H4
        .hIcon = Me.Icon
        .uCallbackMessage = &H200
        
        .szInfo = "Ʈ���̸�尡 �����Ǿ����ϴ�." & Chr(0)  '//ǳ�� �޼��� �����Դϴ�
        .uTimeoutOrVersion = 100000 '//ǳ���� �����ִ� �ð� [1000 = 1��]
        .dwInfoFlags = 1 ' Tip! : ǳ�� ������ : 1 = ����, 2 = ����, 3 = ����
    End With
    
        Shell_NotifyIcon &H0, SysTrayT
End Sub

Private Sub Ser_Change()

End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    GetID
    txtID = vbNullString
End If
End Sub
Private Sub WB_NewWindow2(ppDisp As Object, Cancel As Boolean)
    Dim frm As Form2
    Set frm = New Form2
    Set ppDisp = frm.WB.object
    frm.Show
End Sub
