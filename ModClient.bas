Attribute VB_Name = "ModClient"
Option Explicit

Const API_URL = "http://localhost:8989/"
Const QUOTE As String = """"

Function GetNumOfDimensions(Arr()) As Integer

    Dim Count As Integer
    Count = 0
    
    On Error Resume Next
    Do While True
        Dim Test As Integer
        Test = UBound(Arr, Count + 1)
        
        If Err.Number = 0 Then
            Count = Count + 1
        Else
            Exit Do
        End If
    Loop
    On Error GoTo 0
    
    GetNumOfDimensions = Count
End Function

Sub Append(ByRef Str1 As String, Str2)
    Str1 = Str1 & Str2
End Sub

Function Serialize(Data()) As String
    
    ' Validate data?
    Dim NumOfDimensions As Integer
    NumOfDimensions = GetNumOfDimensions(Data)
    
    Dim Serialized As String
    'Serialized = "{"
    Call Append(Serialized, "[")
    
    Dim Row As Long
    For Row = LBound(Data, 1) To UBound(Data, 1)
        Call Append(Serialized, "[")
        Dim Col As Long
        For Col = LBound(Data, 2) To UBound(Data, 2)
            Dim Value
            Value = Data(Row, Col)
            
            If TypeName(Value) = "String" Or TypeName(Value) = "Date" Then
                Value = QUOTE & Value & QUOTE
            ElseIf TypeName(Value) = "Boolean" Then
                Value = LCase(CStr(Value))
            End If
            
            Call Append(Serialized, Value)
            
            If Col < UBound(Data, 2) Then
                Call Append(Serialized, ",")
            End If
        Next
        Call Append(Serialized, "]")
        
        If Row < UBound(Data, 1) Then
            Call Append(Serialized, ",")
        End If
    Next
    
    Call Append(Serialized, "]")
    'Call Append(Serialized, "}")
    
    Serialize = Serialized

End Function

Function Deserialize(Data As String) As Variant()

    ' Remove
    Data = Mid(Data, 2)
    Data = Left(Data, Len(Data) - 1)
    
    Dim Rows() As String
    Rows = Split(Data, "],")
    
    Dim Temp() As String
    Temp = Split(Rows(0), ",")
    
    Dim RowArr() 'As String
    ReDim RowArr(1 To UBound(Rows) + 1, 1 To UBound(Temp) + 1)
    
    Dim Row As Long
    For Row = LBound(Rows) To UBound(Rows)
        ' Remove
        Dim RowData As String
        RowData = Mid(Rows(Row), 2)
        
        If Mid(RowData, Len(RowData)) = "]" Then
            RowData = Left(RowData, Len(RowData) - 1)
        End If
        
        Dim Cols() As String
        Cols = Split(RowData, ",")
        
        Dim Col As Long
        For Col = LBound(Cols) To UBound(Cols)
            Dim Value As String
            Value = Cols(Col)
            
            If IsNumeric(Value) Then
                RowArr(Row + 1, Col + 1) = CDbl(Value)
            ElseIf Left(Value, 1) = QUOTE Then
                Value = Split(Cols(Col), """")(1)
                If IsDate(Value) Then
                    RowArr(Row + 1, Col + 1) = CDate(Value)
                Else
                    RowArr(Row + 1, Col + 1) = Value
                End If
            Else
                RowArr(Row + 1, Col + 1) = CBool(Value)
            End If
        Next
    Next
    
    Deserialize = RowArr ' Return
End Function

Function SendData() As String

    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Worksheets("Data")
    
    Dim Data()
    Data = Ws.Range("A2:H4").Value
    
    Dim SerializedData
    SerializedData = Serialize(Data)
    
    Dim Request As WinHttpRequest
    Set Request = New WinHttpRequest

    Request.Open "POST", API_URL & "mixed-data", True
    Request.Send SerializedData

    Request.WaitForResponse 50

    Dim Response As String
    Response = Request.ResponseText
    
    Dim NewData() 'As String
    NewData = Deserialize(Response)

    SendData = Response ' Return

End Function

Function Send() As String
    Dim Request As WinHttpRequest
    Set Request = New WinHttpRequest
    
    Request.Open "GET", API_URL & "user/35/profile"
    Request.Send
    
    Dim Response
    Response = Request.ResponseText
    
    Send = Response ' Return
End Function
