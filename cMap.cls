VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984
' Copyright (C) 2011 by David Wright (davidwright@digitalwindfire.com)
' All Rights Reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification or permission, are permitted.
'
' Additional information available at http://www.digitalwindfire.com.
'
'
'        Name: Map
' Description: Map class for Visual Basic 6. Stores primitive types
'              as variants in both the key and value. Allows sorting
'              by either key or value, retrieval of just the keys or
'              values as an array and supports default property
'              syntax (e.g. mymap("mykey") = "myvalue").
'
 
'FIXIT: Declare 'm_Keys' with an early-bound data type                                     FixIT90210ae-R1672-R1B8ZE
Private m_Keys() As Variant
'FIXIT: Declare 'm_Values' with an early-bound data type                                   FixIT90210ae-R1672-R1B8ZE
Private m_Values() As Variant
Private m_Count As Long
 
Public Property Get Count() As Long
    Count = m_Count
End Property
 
'FIXIT: Declare 'Item' and 'Key' with an early-bound data type                             FixIT90210ae-R1672-R1B8ZE
Public Property Get Item(ByVal Key As Variant) As Variant
    Item = m_Values(GetIndexOf(m_Keys, Key))
End Property
 
'FIXIT: Declare 'Key' and 'Value' with an early-bound data type                            FixIT90210ae-R1672-R1B8ZE
Public Property Let Item(ByVal Key As Variant, ByVal Value As Variant)
    If GetIndexOf(m_Keys, Key) = -1 Then
        Add Key, Value
    Else
        m_Values(GetIndexOf(m_Keys, Key)) = Value
    End If
End Property
 
'FIXIT: Declare 'Key' and 'Value' with an early-bound data type                            FixIT90210ae-R1672-R1B8ZE
Public Sub Add(ByVal Key As Variant, ByVal Value As Variant)
    If GetIndexOf(m_Keys, Key) = -1 Then
        Push m_Keys, Key
        Push m_Values, Value
        m_Count = m_Count + 1
    Else
        Err.Raise 457
    End If
End Sub
 
'FIXIT: Declare 'Key' with an early-bound data type                                        FixIT90210ae-R1672-R1B8ZE
Public Sub Delete(ByVal Key As Variant)
    Dim i As Integer
    If KeyExists(Key) Then
'FIXIT: Declare 'new_Keys' with an early-bound data type                                   FixIT90210ae-R1672-R1B8ZE
        Dim new_Keys() As Variant
'FIXIT: Declare 'new_Values' with an early-bound data type                                 FixIT90210ae-R1672-R1B8ZE
        Dim new_Values() As Variant
 
        For i = 0 To UBound(m_Keys)
            If Not (m_Keys(i) = Key) Then
                Push new_Keys, m_Keys(i)
                Push new_Values, m_Values(i)
            End If
        Next
        
        m_Keys = new_Keys
        m_Values = new_Values
    End If
End Sub
 
'FIXIT: Declare 'Key' with an early-bound data type                                        FixIT90210ae-R1672-R1B8ZE
Public Function KeyExists(ByVal Key As Variant) As Boolean
    If GetIndexOf(m_Keys, Key) > -1 Then
        KeyExists = True
    End If
End Function
 
'FIXIT: Declare 'Value' with an early-bound data type                                      FixIT90210ae-R1672-R1B8ZE
Public Function ValueExists(ByVal Value As Variant) As Boolean
    If GetIndexOf(m_Values, Value) > -1 Then
        ValueExists = True
    End If
End Function
 
'FIXIT: Declare 'Keys' with an early-bound data type                                       FixIT90210ae-R1672-R1B8ZE
Public Function Keys() As Variant
    Keys = m_Keys
End Function
 
'FIXIT: Declare 'Values' with an early-bound data type                                     FixIT90210ae-R1672-R1B8ZE
Public Function Values() As Variant
    Values = m_Values
End Function
 
Public Sub SortKeys()
    Sort m_Keys, m_Values
End Sub
 
Public Sub SortValues()
    Sort m_Values, m_Keys
End Sub
 
'FIXIT: Declare 'TheArray' and 'TheElement' with an early-bound data type                  FixIT90210ae-R1672-R1B8ZE
Private Function GetIndexOf(TheArray, ByVal TheElement As Variant) As Long
    GetIndexOf = -1
    
    If IsArrayEmpty(TheArray) Then Exit Function
    
    Dim i As Long
    
    For i = 0 To UBound(TheArray)
        If TheArray(i) = TheElement Then
            GetIndexOf = i
            Exit Function
        End If
    Next
End Function
 
'FIXIT: Declare 'TheArray' with an early-bound data type                                   FixIT90210ae-R1672-R1B8ZE
Private Function IsArrayEmpty(TheArray) As Boolean
    Dim upperBound As Long
    On Error Resume Next
    upperBound = UBound(TheArray)
     
    If Err.Number = 9 Then
        IsArrayEmpty = True
        Err.Clear
    End If
End Function
 
'FIXIT: Declare 'TheArray' and 'TheElement' with an early-bound data type                  FixIT90210ae-R1672-R1B8ZE
Private Sub Push(TheArray, ByVal TheElement)
    If IsArrayEmpty(TheArray) Then
        ReDim TheArray(0)
    Else
        ReDim Preserve TheArray(UBound(TheArray) + 1)
    End If
    
    TheArray(UBound(TheArray)) = TheElement
End Sub
 
' Shell sort routine does an in-place sorting of two related arrays to maintain mapping
' Original version of this shell sort implementation written by Francesco Balena
' and available online at http://www.devx.com/vb2themax/Tip/18
'FIXIT: Declare 'first' and 'second' with an early-bound data type                         FixIT90210ae-R1672-R1B8ZE
Private Sub Sort(first, second)
'FIXIT: Declare 'Value' with an early-bound data type                                      FixIT90210ae-R1672-R1B8ZE
    Dim Value
'FIXIT: Declare 'index' with an early-bound data type                                      FixIT90210ae-R1672-R1B8ZE
    Dim index
'FIXIT: Declare 'index2' with an early-bound data type                                     FixIT90210ae-R1672-R1B8ZE
    Dim index2
'FIXIT: Declare 'firstEl' with an early-bound data type                                    FixIT90210ae-R1672-R1B8ZE
    Dim firstEl
'FIXIT: Declare 'distance' with an early-bound data type                                   FixIT90210ae-R1672-R1B8ZE
    Dim distance
'FIXIT: Declare 'numEls' with an early-bound data type                                     FixIT90210ae-R1672-R1B8ZE
    Dim numEls
 
    lastEl = UBound(first)
    firstEl = LBound(first)
    numEls = lastEl - firstEl + 1
 
    ' find the best value for distance
    Do
        distance = distance * 3 + 1
    Loop Until distance > numEls
 
    Do
        distance = distance \ 3
        For index = distance + firstEl To lastEl
            ' first
            Value = first(index)
            index2 = index
            
            ' second
            value2 = second(index)
            index3 = index
            
            ' first
            Do While (first(index2 - distance) > Value) Xor descending
                ' first
                first(index2) = first(index2 - distance)
                index2 = index2 - distance
                
                ' second
                second(index3) = second(index3 - distance)
                index3 = index3 - distance
                
                If index2 - distance < firstEl Then Exit Do
            Loop
            
            ' first
            first(index2) = Value
            
            ' second
            second(index3) = value2
        Next
    Loop Until distance = 1
End Sub
