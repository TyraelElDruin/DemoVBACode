    Public Sub BuildQueryTab(TheQuery As String, Optional RollUp = False, Optional QueryName = "") ', Optional FormatCurrency = True)
    TheQuery = Replace(TheQuery, "{PROPERTY_DATE}", Sheets(T_M).Range("PROPERTY_DATE").Text)
    TheQuery = Replace(TheQuery, "{STATUS_SWITCH}", Sheets(T_MWBM).Range("STATUS_SWITCH").Value)
    TheQuery = Replace(TheQuery, "{STATUS_SWITCH_SHORT}", Sheets(T_MWBM).Range("STATUS_SWITCH_SHORT").Value)
    AppScreenUpdating TurnOff, "BuildQueryTab"
    Sheets(T_Q).Cells.Clear
    
    Dim rs As ADODB.Recordset
    Dim cn As ADODB.Connection
    Dim strQuery As String
    Dim iCols As Integer
    
    Set cn = New ADODB.Connection
    
    With cn
      .Provider = "Microsoft.ACE.OLEDB.12.0"
      .ConnectionString = "Data Source=" & Application.ActiveWorkbook.FullName & _
      ";Extended Properties=""Excel 8.0;HDR=Yes;""" 'Might want to experiment with IMEX parameter - for data types
    .Open
    End With
    'Microsoft.ACE.OLEDB.12.0 for database engine built in Windows 7 64
    
    'strQuery = "SELECT * FROM [ETL - ResTenants-txtradew$A5:CJ225] AS tmp LEFT JOIN [ETL - ResLeaseCharges-txtradew$A5:N2481] AS tmpy ON tmp.Tenant_Code = tmpy.Tenant_Code"
    strQuery = TheQuery
    
    On Error Resume Next
    Set rs = cn.Execute(strQuery)
    If Err.Number <> 0 Then
        DebugAndMessage "Invalid Query" & vbCr & "Error Number: " & Err.Number & vbCr & "Error Description: " & Err.Description, Error
        Err.Clear
        On Error GoTo 0
        AppScreenUpdating TurnOn, "BuildQueryTab"
        Exit Sub
    End If
    On Error GoTo 0
    
    If Not rs.EOF Then
        Sheets(T_Q).Activate 'This fixes excel bug and formats the dates correctly
        Sheets(T_Q).Range("A1").CopyFromRecordset rs 'useful method
        Sheets(T_Q).Activate 'This fixes excel bug and formats the dates correctly
        Sheets(T_Q).Range("A1").EntireRow.Insert xlShiftDown
        
        'Display Column Headers
        For iCols = 0 To rs.Fields.Count - 1
            Sheets(T_Q).Cells(1, iCols + 1).Value = rs.Fields(iCols).Name
        Next
        
        With Intersect(Sheets(T_Q).Range("A1").EntireRow, Sheets(T_Q).UsedRange)
            .Interior.Color = COLOR_GRAY
            .Font.Bold = True
        End With
        
        Sheets(T_Q).UsedRange.Columns.AutoFit
    Else
        If QueryName = "" Then
            DebugAndMessage "Empty Query Returned [" & TheQuery & "]", Code
        Else
            DebugAndMessage "Empty Query Returned [" & QueryName & "]", Code
        End If
    End If
    
    If RollUp Then
        Dim RTypes As New Dictionary
        RTypes.Add adBigInt, 1
        RTypes.Add adCurrency, 2
        RTypes.Add adDecimal, 3
        RTypes.Add adDouble, 4
        RTypes.Add adInteger, 5
        RTypes.Add adNumeric, 6
        RTypes.Add adSmallInt, 7
        With Sheets(T_Q)
            Dim LastRow, LastCol, x
            LastRow = GetLastRow("A", T_Q)
            LastCol = GetLastCol(1, T_Q)
            For x = 1 To LastCol
                If RTypes.Exists(rs.Fields(x - 1).Type) Then
                    .Cells(LastRow + 1, x).Formula = "=SUM(" & .Cells(2, x).Address & ":" & .Cells(LastRow, x).Address & ")"
                    'If FormatCurrency Then .Cells(LastRow + 1, x).NumberFormat = "$#,##0.00"
                Else
                    If x = 1 Then .Cells(LastRow + 1, x).Value = "Grand Totals:"
                End If
            Next x
            With .Range(.Cells(LastRow + 1, 1), .Cells(LastRow + 1, LastCol))
                .Copy
                .PasteSpecial xlPasteValues
                .Interior.Color = COLOR_GRAY
                .Font.Bold = True
            End With
        End With
        Sheets(T_Q).UsedRange.Columns.AutoFit
    End If
    rs.Close
    Set cn = Nothing
    Set rs = Nothing
    Set RTypes = Nothing
    AppScreenUpdating TurnOn, "BuildQueryTab"
    End Sub
