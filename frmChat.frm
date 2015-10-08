VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmChat 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Chat em rede local"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   225
      Top             =   5310
   End
   Begin VB.ComboBox cmbPessoas 
      Height          =   315
      ItemData        =   "frmChat.frx":058A
      Left            =   1260
      List            =   "frmChat.frx":0591
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4950
      Width           =   7845
   End
   Begin VB.TextBox txtMsg 
      Height          =   930
      Left            =   1260
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   5400
      Width           =   6450
   End
   Begin VB.CommandButton btSend 
      Caption         =   "Enviar"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   330
      Left            =   7875
      TabIndex        =   2
      Top             =   5400
      Width           =   1230
   End
   Begin VB.TextBox txtChat 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4650
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   90
      Width           =   9060
   End
   Begin MSWinsockLib.Winsock wskListen 
      Left            =   5760
      Top             =   5490
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      LocalPort       =   32865
   End
   Begin MSWinsockLib.Winsock wskSender 
      Left            =   6300
      Top             =   5490
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      LocalPort       =   32864
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Mensagem:"
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   5445
      Width           =   2040
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enviar para:"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   4995
      Width           =   2040
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984
Option Explicit
Public WithEvents sysTray As frmSysTray
Attribute sysTray.VB_VarHelpID = -1
Public permiteSaida As Boolean
Public listaPessoas As New cMap
Enum eTipo
    CMDTIPO_MENSAGEM
    CMDTIPO_COMANDO
    CMDTIPO_WHO
    CMDTIPO_HELLO
    CMDTIPO_BYE
End Enum
Private Type tMensagem
    emissor As String
    destinatario As String
    dados As String
    tipo As eTipo
End Type
Const DELIMITER = vbTab
Function MinhaAssinatura() As String
    MinhaAssinatura = Environ("USERNAME") & " em " & Environ("USERDOMAIN")
End Function
Sub Enviar(ByVal t As eTipo, Optional ByVal dados As String, Optional ByVal ip As String = "255.255.255.255", Optional destino As String)
    Call wskSender.Connect(ip, wskListen.LocalPort)
    Call wskSender.SendData(t & DELIMITER & MinhaAssinatura & DELIMITER & dados & DELIMITER & destino)
    Call wskSender.Close
End Sub
Sub AlguemFalaPara(ByVal emissor As String, ByVal destinatario, ByVal mensagem As String)
    Call ExibirNoChat("_____________")
    Call ExibirNoChat(Format(Now, "dd/MM/yyyy hh:nn:ss"))
    Call ExibirNoChat("De: " & emissor)
    Call ExibirNoChat("Para: " & IIf(destinatario = "", "todos", IIf(destinatario = MinhaAssinatura, "você", destinatario)) & vbCrLf)
    Call ExibirNoChat(mensagem)
            
End Sub
Private Sub btSend_Click()
    If cmbPessoas.ListIndex <> 0 And cmbPessoas.Text <> (MinhaAssinatura) Then
        'Call ExibirNoChat("" & MinhaAssinatura & IIf(cmbPessoas.ListIndex <> 0, " fala reservadamente para " & cmbPessoas.Text, "") & " " & vbCrLf & vbTab & txtMsg.Text)
        Call AlguemFalaPara(MinhaAssinatura, IIf(cmbPessoas.ListIndex <> 0, cmbPessoas.Text, ""), txtMsg.Text)
    End If
    Enviar CMDTIPO_MENSAGEM, txtMsg.Text, listaPessoas.Item(cmbPessoas.Text), IIf(cmbPessoas.ListIndex <> 0, cmbPessoas.Text, "")
    txtMsg.Text = ""
End Sub
Sub ExibirNoChat(ByVal msg As String)
    txtChat.Text = txtChat.Text & msg & vbCrLf
    txtChat.SelStart = Len(txtChat.Text)
End Sub
Sub AdicionarPessoa(ByVal nome As String, ByVal ip As String)
    On Error GoTo fim
    listaPessoas.Add nome, ip
    cmbPessoas.AddItem nome
    If nome <> MinhaAssinatura Then
        Call ExibirNoChat(vbCrLf & "** " & nome & " está on-line.")
    End If
fim:
    
End Sub
Private Sub Form_Load()
    On Error Resume Next
    Dim c As Control
    For Each c In Controls
'FIXIT: 'FontName' is not a property of the generic 'Control' object in Visual Basic .NET. To access 'FontName' declare 'c' using its actual type instead of 'Control'     FixIT90210ae-R1460-RCFE85
        c.FontName = "Segoe UI"
       ' c.FontSize = 11
    Next
    
    On Error GoTo erro
    wskListen.Bind
    Visible = False
    Set sysTray = New frmSysTray
    cmbPessoas.ListIndex = 0
    listaPessoas.Add "todos", "255.255.255.255"
    With sysTray
        .AddMenuItem "Mostrar chat", "open", True
        .AddMenuItem "-"
        .AddMenuItem "Sair", "exit"
        .ToolTip = Me.Caption
        .IconHandle = Me.Icon.Handle
        Timer1_Timer
        
        
    End With
    Exit Sub
erro:
    End
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        If Not permiteSaida Then
            Visible = False
            Cancel = -1
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload sysTray
End Sub

Private Sub sysTray_BalloonClicked()
    Visible = True
    txtChat.SelStart = Len(txtChat.Text)
End Sub

Private Sub sysTray_MenuClick(ByVal lIndex As Long, ByVal sKey As String)
    Select Case sKey
        Case "open"
            Call frmChat.Show
            txtChat.SelStart = Len(txtChat.Text)
        Case "exit"
            permiteSaida = True
            Unload Me
    End Select
End Sub

Private Sub sysTray_SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
    frmChat.Visible = True
    txtChat.SelStart = Len(txtChat.Text)
End Sub

Private Sub sysTray_SysTrayMouseUp(ByVal eButton As MouseButtonConstants)
    Call sysTray.ShowMenu
End Sub

Private Sub Timer1_Timer()
    Call Enviar(CMDTIPO_HELLO, wskListen.LocalIP)
End Sub

Private Sub txtMsg_Change()
    btSend.Enabled = (txtMsg.Text <> "")
End Sub

Private Sub wskListen_DataArrival(ByVal bytesTotal As Long)
    Dim sData As String
    Dim m As tMensagem
    Dim spl() As String
    Call wskListen.GetData(sData)
    spl = Split(sData, DELIMITER)
    m.emissor = spl(1)
    m.dados = spl(2)
    On Error Resume Next
    m.destinatario = spl(3)
    On Error GoTo pulo
    Select Case spl(0)
        Case CMDTIPO_MENSAGEM
            On Error GoTo 0
            'Call ExibirNoChat("_____________" & vbCrLf)
            'Call ExibirNoChat(" " & m.emissor & IIf(m.destinatario <> "", " fala reservadamente para " & IIf(m.destinatario = MinhaAssinatura, "você", m.destinatario), "fala para todos") & ": " & m.dados)
            Call AlguemFalaPara(m.emissor, m.destinatario, m.dados)
            If Not Visible Then
                Call sysTray.ShowBalloonTip( _
                                            "Para: " & IIf(m.destinatario = "", "todos", IIf(m.destinatario = MinhaAssinatura, "você", m.destinatario)) & vbCrLf & vbCrLf & _
                                            m.dados, "Nova mensagem de " & m.emissor, NIIF_NOSOUND Or NIIF_INFO, 32760)
            End If
            
        Case CMDTIPO_HELLO
            Call AdicionarPessoa(m.emissor, m.dados)
            If m.emissor <> MinhaAssinatura Then
                Call Enviar(CMDTIPO_WHO, wskListen.LocalIP)
            End If
            
        Case CMDTIPO_WHO
            Call AdicionarPessoa(m.emissor, m.dados)
            
    End Select
pulo:
End Sub

