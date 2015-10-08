VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPessoa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
 Option Explicit
Private mvarNome As String 'local copy
Private mvarIP As String 'local copy
Private mvarMaquina As String 'local copy
Private mvarUID As String 'local copy
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent DadosRecebidos[(arg1, arg2, ... , argn)]
Public Event DadosRecebidos(dados As Object)

'FIXIT: Declare 'dados' with an early-bound data type                                      FixIT90210ae-R1672-R1B8ZE
Public Sub Enviar(dados As Object)
    
End Sub

Public Property Let UID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UID = 5
    mvarUID = vData
End Property


Public Property Get UID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UID
    UID = mvarUID
End Property



Public Property Let Maquina(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Maquina = 5
    mvarMaquina = vData
End Property


Public Property Get Maquina() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Maquina
    Maquina = mvarMaquina
End Property



Public Property Let IP(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IP = 5
    mvarIP = vData
End Property


Public Property Get IP() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IP
    IP = mvarIP
End Property



Public Property Let Nome(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Nome = 5
    mvarNome = vData
End Property


Public Property Get Nome() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Nome
    Nome = mvarNome
End Property



