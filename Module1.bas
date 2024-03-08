Attribute VB_Name = "Module1"
Private Function ReDimPreserve(MyArray As Variant, nNewFirstUBound As Long, nNewLastUBound As Long) As Variant
    ' credit: https://wellsr.com/vba/2016/excel/dynamic-array-with-redim-preserve-vba/
    Dim i, j As Long
    Dim nOldFirstUBound, nOldLastUBound, nOldFirstLBound, nOldLastLBound As Long
    Dim TempArray() As Variant 'Change this to "String" or any other data type if want it to work for arrays other than Variants. MsgBox UCase(TypeName(MyArray))
    '---------------------------------------------------------------
    'COMMENT THIS BLOCK OUT IF YOU CHANGE THE DATA TYPE OF TempArray
        If InStr(1, UCase(TypeName(MyArray)), "VARIANT") = 0 Then
            MsgBox "This function only works if your array is a Variant Data Type." & vbNewLine & _
                   "You have two choice:" & vbNewLine & _
                   " 1) Change your array to a Variant and try again." & vbNewLine & _
                   " 2) Change the DataType of TempArray to match your array and comment the top block out of the function ReDimPreserve" _
                    , vbCritical, "Invalid Array Data Type"
            End
        End If
    '---------------------------------------------------------------
    ReDimPreserve = False
    'check if its in array first
    If Not IsArray(MyArray) Then MsgBox "You didn't pass the function an array.", vbCritical, "No Array Detected": End
    
    'get old lBound/uBound
    nOldFirstUBound = UBound(MyArray, 1): nOldLastUBound = UBound(MyArray, 2)
    nOldFirstLBound = LBound(MyArray, 1): nOldLastLBound = LBound(MyArray, 2)
    'create new array
    ReDim TempArray(nOldFirstLBound To nNewFirstUBound, nOldLastLBound To nNewLastUBound)
    'loop through first
    For i = LBound(MyArray, 1) To nNewFirstUBound
        For j = LBound(MyArray, 2) To nNewLastUBound
            'if its in range, then append to new array the same way
            If nOldFirstUBound >= i And nOldLastUBound >= j Then
                TempArray(i, j) = MyArray(i, j)
            End If
        Next
    Next
    'return the array redimmed
    If IsArray(TempArray) Then ReDimPreserve = TempArray
End Function

Function duckdb(query As String)
    
    Dim rs As ADODB.Recordset
    Dim FieldCount As Long
    Dim resp() As Variant
    
    Set rs = make_query(query)
    
    FieldCount = rs.Fields.Count
    ReDim resp(2, FieldCount - 1)
    
    j = 0
    For Each Field In rs.Fields
        resp(0, j) = Field.Name
        j = j + 1
    Next Field
        i = 1
    
    ' Iterate through each row in the recrodset and return a dynamic array
    Do While Not rs.EOF
      If i > 1 Then
        resp = ReDimPreserve(resp, i + 1, FieldCount - 1)
      End If
    
      g = 0
      For Each Field In rs.Fields
         If IsNull(Field.Value) Then
              resp(i, g) = ""
          Else
              resp(i, g) = Field.Value
          End If
          g = g + 1
      Next Field
      i = i + 1
      rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    duckdb = resp
    
End Function

Function make_query(query) As ADODB.Recordset
    ' The default DSN is an in meomory duckdb odbc connection called 'quack'
    ' Additional extensions boiler plate can be added here:
    
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    Set conn = New ADODB.Connection
    conn.Open "DSN=excel-duckdb;"
    
    ' Extension management
    q0 = "install azure;"
    q1 = "load azure;"
    'q2 = "set azure_storage_connection_string='';"
    q3 = query & ";"
    
    Set rs = New ADODB.Recordset
    rs.Open q0, conn
    rs.Open q1, conn
    'rs.Open q2, conn
    rs.Open q3, conn
    
    Set make_query = rs
    
    Exit Function
    
End Function


