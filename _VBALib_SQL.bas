Attribute VB_Name = "_VBALib_SQL"
Option Compare Database





'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function SQLsqueezer(sqlString As String, serverAddress As String, Optional userID As String = vbNullString, Optional password As String = vbNullString, Optional database As String = vbNullString, Optional shouldReturnResultsAtRange As Range = Nothing, Optional shouldReturnListArray As Boolean = True, Optional shouldReturnHeadersAtTopOfRange As Boolean = True) As Variant
    'Executes the given SQL string on the specified database using the specified credentials and can return the results in a variaty of configurations.
    'Will return results to an array if specified using a parameter option or will return results to a range on sheet, or any combination
    'Ex: SQLsqueezer sqlString, "server-address", "my_user_id_string", "my_user_password", "database_to_connect_to", shouldReturnResultsAtRange:=ActiveSheet.Range("A1"), shouldReturnListArray:=False, shouldReturnHeadersAtTopOfRange:=True
    Dim adoConnection As Object: Set adoConnection = CreateObject("ADODB.Connection")
    Dim adoRcdSource As Object: Set adoRcdSource = CreateObject("ADODB.Recordset")
    On Error GoTo Errs: 'For Excel DB
    adoConnection.Open "Provider=SQLOLEDB.1;Server=" & serverAddress & ";Database=" & database & ";User Id=" & userID & ";Password=" & password & ";"
    If UCase(Left(sqlString, 6)) = "SELECT" Then
        adoRcdSource.Open sqlString, adoConnection, 3
        If shouldReturnListArray = True Then
            If (adoRcdSource.BOF Or adoRcdSource.EOF) = False Then SQLsqueezer = adoRcdSource.GetRows
        End If
        If Not shouldReturnResultsAtRange Is Nothing Then
            With shouldReturnResultsAtRange
                If shouldReturnHeadersAtTopOfRange Then
                    Dim header As Variant, currentColumn As Long: currentColumn = 1
                    For Each header In adoRcdSource.Fields
                        .Cells(1, currentColumn).Value = header.Name
                        .Cells(1, currentColumn).Font.Bold = True
                        currentColumn = currentColumn + 1
                    Next header
                    .Cells(2).CopyFromRecordset adoRcdSource
                    freezePanesOnSheet shouldReturnResultsAtRange.Parent, .Range("A2")
                    .Parent.Columns.AutoFit
                Else
                    .Cells(1).CopyFromRecordset adoRcdSource
                End If
            End With
        End If
    Else
        adoConnection.Execute sqlString
    End If
    GoTo NormalExit
Errs:
    MsgBox Err.Description, vbCritical, "Error!"
    Err.Clear: On Error GoTo 0: On Error GoTo -1
NormalExit:
    On Error Resume Next
    adoRcdSource.Close
    On Error GoTo 0
    Set adoConnection = Nothing
    Set adoRcdSource = Nothing
End Function





'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Sub RunSQL_CmdWithoutWarning(SQL_cmd As String)
    'Run SQL command without warning msg
    DoCmd.SetWarnings False
    Application.SetOption "Confirm Action Queries", False
    DoCmd.RunSQL SQL_cmd
    Application.SetOption "Confirm Action Queries", True
    DoCmd.SetWarnings True
End Sub
Public Function CreateSqlSeg_VectorColAgg(col_pattern As String, str_agg As String, Idx_start As Integer, Idx_end As Integer, Optional wildcard As String = "#") As String
    'To create an expression that consists of a set of vector columns aggregated in a specified pattern
    On Error GoTo Err_CreateSqlSeg_VectorColAgg
    Dim SQL_Seg As String, col_idx As Integer
    For col_idx = Idx_start To Idx_end
        SQL_Seg = SQL_Seg & Replace(col_pattern, wildcard, col_idx) & " " & str_agg & " "
    Next col_idx
    SQL_Seg = Left(SQL_Seg, Len(SQL_Seg) - 2)
Exit_CreateSqlSeg_VectorColAgg:
    CreateSqlSeg_VectorColAgg = SQL_Seg
    Exit Function
Err_CreateSqlSeg_VectorColAgg:
    ShowMsgBox (Err.Description)
    Resume Exit_CreateSqlSeg_VectorColAgg
End Function
Public Sub ModifyTbl_ReSelect(Tbl_name, str_select)
    'Re-Select table columns
    Dim Tbl_T_name As String
    Tbl_T_name = Tbl_name & "_temp"
    DelTable (Tbl_T_name)
    SQL_cmd = "SELECT " & str_select & " " & vbCrLf & _
                "INTO [" & Tbl_T_name & "]" & vbCrLf & _
                "FROM [" & Tbl_name & "]" & vbCrLf & ";"
    RunSQL_CmdWithoutWarning (SQL_cmd)
    DelTable (Tbl_name)
    DoCmd.Rename Tbl_name, acTable, Tbl_T_name
End Sub
Public Function UpdateTblColBatchly(Tbl_src_name As String, Str_Col_Update As String, SQL_Format_Set As String, SQL_Format_Where As String) As String
    'Update multiple columns of a table under the same condition
    On Error GoTo Err_UpdateTblColBatchly
    Dim FailedReason As String
    If TableExist(Tbl_src_name) = False Then
        FailedReason = Tbl_src_name & "does not exist!"
        GoTo Exit_UpdateTblColBatchly
    End If
    Str_Col_Update = Trim(Str_Col_Update)
    
    
    If Str_Col_Update = "*" Then
        With CurrentDb
            Dim RS_Tbl_src As Recordset
            Set RS_Tbl_src = .OpenRecordset(Tbl_src_name)
            
            Dim fld_idx As Integer
            
            With RS_Tbl_src
                For fld_idx = 0 To .Fields.count - 1
                    Call UpdateTblCol(Tbl_src_name, .Fields(fld_idx).Name, SQL_Format_Set, SQL_Format_Where)
                Next fld_idx
                
                .Close
            End With 'RS_Tbl_src
            
            .Close
        End With 'CurrentDb
        
    Else
        Dim Col_Update As Variant, Col_Idx As Integer, ColName As String
        Col_Update = SplitStrIntoArray(Str_Col_Update, ",")
        For Col_Idx = 0 To UBound(Col_Update)
            ColName = Col_Update(Col_Idx)
            Call UpdateTblCol(Tbl_src_name, ColName, SQL_Format_Set, SQL_Format_Where)
        Next Col_Idx
    End If
Exit_UpdateTblColBatchly:
    UpdateTblColBatchly = FailedReason
    Exit Function
Err_UpdateTblColBatchly:
    FailedReason = Err.Description
    Resume Exit_UpdateTblColBatchly
End Function
Public Function UpdateTblCol(Tbl_src_name As String, ColName As String, SQL_Format_Set, SQL_Format_Where As String) As String
    'Update a column of a table under a specified condition
    On Error GoTo Err_UpdateTblCol
    Dim FailedReason As String, SQL_Seq_Set As String
    Dim SQL_cmd As String, SQL_Seq_Where As String
    If TableExist(Tbl_src_name) = False Then
        FailedReason = Tbl_src_name & "does not exist!"
        GoTo Exit_UpdateTblCol
    End If
    SQL_Seg_Set = "SET [" & ColName & "] = " & Replace(SQL_Format_Set, "*", "[" & ColName & "]") & " "
    If SQL_Format_Where <> vbNullString Then SQL_Seg_Where = "WHERE " & Replace(SQL_Format_Where, "*", "[" & ColName & "]")
    SQL_cmd = "UPDATE " & Tbl_src_name & " " & vbCrLf & SQL_Seg_Set & " " & vbCrLf & SQL_Seg_Where & " " & vbCrLf & ";"
    RunSQL_CmdWithoutWarning (SQL_cmd)
Exit_UpdateTblCol:
    UpdateTblCol = FailedReason
    Exit Function
Err_UpdateTblCol:
    FailedReason = Err.Description
    Resume Exit_UpdateTblCol
End Function
Public Function CreateTbl_ColAndExpr(Tbl_src_name As String, Str_Col_Id As String, Str_Col_Order As String, SQL_Seg_ColAndExpr As String, SQL_Seg_Where As String, Tbl_output_name As String) As String
    'Create Table with dedicated Column and Expressions from a source table
    On Error GoTo Err_CreateTbl_ColAndExpr: Dim FailedReason As String
    If TableExist(Tbl_src_name) = False Then FailedReason = Tbl_src_name & "does not exist!": GoTo Exit_CreateTbl_ColAndExpr
    Dim SQL_cmd As String, SQL_Seq_Select As String, SQL_Seq_OrderBy As String
    Dim Col_Order As Variant, Col_Id As Variant, Col_Idx As Integer
    SQL_Seg_Select = "SELECT ": SQL_Seg_OrderBy = vbNullString
    Col_Order = SplitStrIntoArray(Str_Col_Order, ",")
    Col_Id = SplitStrIntoArray(Str_Col_Id, ",")
    DelTable (Tbl_output_name)
    For Col_Idx = 0 To UBound(Col_Id)
        SQL_Seg_Select = SQL_Seg_Select & "[" & Col_Id(Col_Idx) & "], "
    Next Col_Idx
    SQL_Seq_Select = IIf(SQL_Seq_ColAndExpr <> vbNullString, SQL_Seq_Select & SQL_Seq_ColAndExpr, Left(SQL_Seq_Select, Len(SQL_Seq_ColAndExpr)))    
    SQL_Seg_Select = SQL_Seg_Select & SQL_Seg_ColAndExpr
    If UBound(Col_Order) >= 0 Then
        SQL_Seg_OrderBy = "ORDER BY "
        For Col_Idx = 0 To UBound(Col_Order)
            SQL_Seg_OrderBy = SQL_Seg_OrderBy & "[" & Col_Order(Col_Idx) & "], "
        Next Col_Idx
        SQL_Seg_OrderBy = Left(SQL_Seg_OrderBy, Len(SQL_Seg_OrderBy) - 2)
    End If
    If SQL_Seg_Where <> vbNullString Then SQL_Seg_Where = "WHERE " & SQL_Seg_Where
    SQL_cmd = SQL_Seg_Select & " " & vbCrLf & "INTO [" & Tbl_output_name & "] " & vbCrLf & "FROM [" & _
              Tbl_src_name & "] " & vbCrLf & SQL_Seg_Where & " " & vbCrLf & SQL_Seg_OrderBy & " " & vbCrLf & ";"
    'MsgBox SQL_cmd
    RunSQL_CmdWithoutWarning (SQL_cmd)
Exit_CreateTbl_ColAndExpr:
    CreateTbl_ColAndExpr = FailedReason
    Exit Function
Err_CreateTbl_ColAndExpr:
    FailedReason = Err.Description
    Resume Exit_CreateTbl_ColAndExpr
End Function
Public Function CreateTbl_Group(Tbl_input_name As String, Tbl_output_name As String, Str_Col_Group As String, Optional Str_GroupFunc_all As String = "", Optional GF_all_dbTypes As Variant = "", Optional Str_Col_UnSelected As String = "", Optional ByVal GroupFunc_Col_Pairs As Variant = "", Optional SQL_Seg_Where As String = "", Optional Str_Col_Order As String = "") As String
    'Create Table of group function, there is a default Group function for all columns, columns can be specified to different group fucntion
    On Error GoTo Err_CreateTbl_Group: Dim FailedReason As String
    If TableValid(Tbl_input_name) = False Then FailedReason = Tbl_input_name & " is not valid!": GoTo Exit_CreateTbl_Group
    If Len(Str_Col_Group) = 0 Then FailedReason = "No Any Group Columns": GoTo Exit_CreateTbl_Group   
    If Str_GroupFunc_all <> vbNullString Then
        If UBound(GF_all_dbTypes) < 0 Then FailedReason = "No db Type is assigned for the general group function": GoTo Exit_CreateTbl_Group
    End If
    If VarType(GroupFunc_Col_Pairs) <> vbArray + vbVariant Then
        If Str_GroupFunc_all = vbNullString Then
            FailedReason = "No Any Group Functions for all or specified columns"
            GoTo Exit_CreateTbl_Group
        Else
            GroupFunc_Col_Pairs = Array()
        End If
    End If
    Dim Col_Group As Variant, Col_UnSelected As Variant, Col_Order As Variant
    Dim GF_C_P_idx As Integer, col_idx As Integer
    For GF_C_P_idx = 0 To UBound(GroupFunc_Col_Pairs)
        GroupFunc_Col_Pairs(GF_C_P_idx)(1) = SplitStrIntoArray(GroupFunc_Col_Pairs(GF_C_P_idx)(1) & "", ",")
    Next GF_C_P_idx
    Str_GroupFunc_all = Trim(Str_GroupFunc_all)
    Col_Group = SplitStrIntoArray(Str_Col_Group, ",")
    Col_UnSelected = SplitStrIntoArray(Str_Col_UnSelected, ",")
    Col_Order = SplitStrIntoArray(Str_Col_Order, ",")
    DelTable (Tbl_output_name)
    With CurrentDb
        Dim RS_Tbl_input As Recordset: Set RS_Tbl_input = .OpenRecordset(Tbl_input_name)
        With RS_Tbl_input
            Dim SQL_Seg_Select As String, SQL_Seg_GroupBy As String, SQL_Seg_OrderBy As String
            SQL_Seg_Select = "SELECT ": SQL_Seg_GroupBy = "GROUP BY ": SQL_Seg_OrderBy = vbNullString
            Dim fld_idx As Integer, fld_name As String, Col_GroupBy As Variant
            Dim IsColForGroupBy As Boolean, NumOfCol_Group_found As Integer
            Dim GroupFunc_Col_Pair As Variant, GroupFunc As String
            NumOfCol_Group_found = 0
            For fld_idx = 0 To .Fields.count - 1
                fld_name = .Fields(fld_idx).Name
                IsColForGroupBy = False
                If NumOfCol_Group_found <= UBound(Col_Group) Then
                    If FindStrInArray(Col_Group, fld_name) > -1 Then
                        SQL_Seg_GroupBy = SQL_Seg_GroupBy & "[" & fld_name & "], "
                        IsColForGroupBy = True
                        NumOfCol_Group_found = NumOfCol_Group_found + 1
                    End If
                End If                           
                If IsColForGroupBy = True Then
                    SQL_Seg_Select = SQL_Seg_Select & "[" & fld_name & "], "
                ElseIf FindStrInArray(Col_UnSelected, fld_name) < 0 Then
                    GroupFunc = vbNullString
                    For Each GroupFunc_Col_Pair In GroupFunc_Col_Pairs
                        If FindStrInArray(GroupFunc_Col_Pair(1), fld_name) > -1 Then GroupFunc = GroupFunc_Col_Pair(0)
                    Next GroupFunc_Col_Pair
                    If GroupFunc = vbNullString And Str_GroupFunc_all <> vbNullString Then
                        For Each GF_all_dbType In GF_all_dbTypes
                            If .Fields(fld_idx).Type = GF_all_dbType Then GroupFunc = Str_GroupFunc_all
                        Next GF_all_dbType
                    End If
                    If GroupFunc <> vbNullString Then SQL_Seg_Select = SQL_Seg_Select & GroupFunc & "([" & Tbl_input_name & "].[" & fld_name & "]) AS [" & fld_name & "], "
                End If
Next_CreateTbl_Group_1:
            Next fld_idx
            SQL_Seg_Select = Left(SQL_Seg_Select, Len(SQL_Seg_Select) - 2)
            SQL_Seg_GroupBy = Left(SQL_Seg_GroupBy, Len(SQL_Seg_GroupBy) - 2)
            .Close
        End With 'RS_Tbl_input
        If UBound(Col_Order) >= 0 Then
            SQL_Seg_OrderBy = "ORDER BY "
            For col_idx = 0 To UBound(Col_Order)
                SQL_Seg_OrderBy = SQL_Seg_OrderBy & "[" & Col_Order(col_idx) & "], "
            Next col_idx
            SQL_Seg_OrderBy = Left(SQL_Seg_OrderBy, Len(SQL_Seg_OrderBy) - 2)
        End If
        If SQL_Seg_Where <> vbNullString Then SQL_Seg_Where = "WHERE " & SQL_Seg_Where
        Dim SQL_cmd As String: SQL_cmd = SQL_Seg_Select & " " & vbCrLf & "INTO [" & Tbl_output_name & "] " & vbCrLf & "FROM [" & Tbl_input_name & "] " & vbCrLf & SQL_Seg_Where & " " & vbCrLf & SQL_Seg_GroupBy & " " & vbCrLf & SQL_Seg_OrderBy & " " & vbCrLf & ";"
        'MsgBox SQL_cmd
        RunSQL_CmdWithoutWarning (SQL_cmd)
        .Close
    End With 'CurrentDb
Exit_CreateTbl_Group:
    CreateTbl_Group = FailedReason
    Exit Function
Err_CreateTbl_Group:
    FailedReason = Err.Description
    Resume Exit_CreateTbl_Group
End Function
Public Function CreateTbls_Group(Tbl_MT_name As String) As String
    'Create a set of grouped table, the grouping config is set in a specified table
    On Error GoTo Err_CreateTbls_Group : Dim FailedReason As String
    If TableExist(Tbl_MT_name) = False Then FailedReason = Tbl_MT_name & " does not exist!": GoTo Exit_CreateTbls_Group
    With CurrentDb
        Dim RS_Tbl_MT As Recordset: Set RS_Tbl_MT = .OpenRecordset(Tbl_MT_name)
        With RS_Tbl_MT
            Dim FailedReason_1, Tbl_src_name, Tbl_Group_name As String
            Dim Str_Col_Group, Str_Col_UnSelected, Str_GroupFunc_all As String
            Dim GF_all_dbTypes, GroupFunc_Col_Pairs As Variant
            Dim SQL_Seg_Where, Str_Col_Order As String
            .MoveFirst
            Do Until .EOF
                If .Fields("Enable").Value = False Then GoTo Loop_CreateTbls_Group_1
                Tbl_src_name = .Fields("Tbl_src").Value
                If TableExist(Tbl_src_name) = False Then GoTo Loop_CreateTbls_Group_1
                Tbl_Group_name = .Fields("Tbl_Group").Value
                If Len(Tbl_Group_name) = 0 Then  GoTo Loop_CreateTbls_Group_1
                If IsNull(.Fields("Col_Group").Value) = True Then GoTo Loop_CreateTbls_Group_1
                Str_GroupFunc_all = IIf(IsNull(.Fields("GroupFunc_all").Value) = True, vbNullString, .Fields("GroupFunc_all").Value)
                GF_all_dbTypes = Array(dbInteger, dbLong, dbSingle, dbDouble, dbDecimal)
                Str_Col_Sum = IIf(IsNull(.Fields("Col_Sum").Value) = True, vbNullString, .Fields("Col_Sum").Value)
                Str_Col_Avg = IIf(IsNull(.Fields("Col_Avg").Value) = True, vbNullString, .Fields("Col_Avg").Value)
                Str_Col_Max = IIf(IsNull(.Fields("Col_Max").Value) = True, vbNullString, .Fields("Col_Max").Value)
                GroupFunc_Col_Pairs = Array(Array("SUM", Str_Col_Sum), Array("AVG", Str_Col_Avg), Array("MAX", Str_Col_Max))
                Str_Col_Order = IIf(IsNull(.Fields("Col_Order").Value) = True, vbNullString, .Fields("Col_Order").Value)
                SQL_Seq_Where = IIf(IsNull(.Fields("Cond").Value) = True, vbNullString, .Fields("Cond").Value)
                FailedReason_1 = CreateTbl_Group(Tbl_src_name, Tbl_Group_name, .Fields("Col_Group").Value, Str_GroupFunc_all:=Str_GroupFunc_all, GF_all_dbTypes:=GF_all_dbTypes, GroupFunc_Col_Pairs:=GroupFunc_Col_Pairs, Str_Col_Order:=Str_Col_Order)
                If FailedReason_1 <> vbNullString Then FailedReason = FailedReason & Tbl_Group_name & ": " & FailedReason_1 & vbCrLf
Loop_CreateTbls_Group_1:
                .MoveNext
            Loop
            .Close
        End With 'RS_Tbl_MT
        .Close
    End With 'CurrentDb
Exit_CreateTbls_Group:
    CreateTbls_Group = FailedReason
    Exit Function
Err_CreateTbls_Group:
    FailedReason = Err.Description
    Resume Exit_CreateTbls_Group
End Function
Public Function CreateTbl_JoinTwoTbl(Tbl_src_1_name As String, Tbl_src_2_name As String, JoinCond As String, ColSet_Join_1 As Variant, ColSet_Join_2 As Variant, Tbl_des_name As String, Optional ColSet_src_1 As Variant = Null, Optional ColSet_src_2 As Variant = Null, Optional ColSet_Order As Variant = Null) As String
    'Create table which are joined from two tables
    On Error GoTo Err_CreateTbl_JoinTwoTbl: Dim FailedReason As String
    If TableExist(Tbl_src_1_name) = False Then FailedReason = Tbl_src_1_name & "does not exist!": GoTo Exit_CreateTbl_JoinTwoTbl
    If TableExist(Tbl_src_2_name) = False Then FailedReason = Tbl_src_2_name & "does not exist!": GoTo Exit_CreateTbl_JoinTwoTbl
    If IsNull(ColSet_Join_1) = True Then GoTo Exit_CreateTbl_JoinTwoTbl
    If IsNull(ColSet_Join_2) = True Then GoTo Exit_CreateTbl_JoinTwoTbl
    DelTable (Tbl_des_name): Dim Col_Idx As Integer
    With CurrentDb
        If IsNull(ColSet_src_1) = True Then
            Dim RS_Tbl_src As Recordset: Set RS_Tbl_src = .OpenRecordset(Tbl_src_1_name)
            Dim fld_idx As Integer, fld_name As String: ColSet_src_1 = Array()
            With RS_Tbl_src
                For fld_idx = 0 To .Fields.count - 1
                    fld_name = .Fields(fld_idx).name
                    Call AppendArray(ColSet_src_1, Array("[" & fld_name & "]"))
                Next fld_idx
                .Close
            End With 'RS_Tbl_src
        End If
        If IsNull(ColSet_src_2) = True Then
            Set RS_Tbl_src = .OpenRecordset(Tbl_src_2_name)
            With RS_Tbl_src
                Dim NumOfColSet_Join_found As Integer
                NumOfColSet_Join_found = 0
                ColSet_src_2 = Array()
                For fld_idx = 0 To .Fields.count - 1
                    fld_name = .Fields(fld_idx).name
                    If NumOfColSet_Join_found <= UBound(ColSet_Join_2) And FindStrInArray(ColSet_Join_2, fld_name) > -1 Then
                        NumOfColSet_Join_found = NumOfColSet_Join_found + 1
                    Else
                        Call AppendArray(ColSet_src_2, Array("[" & fld_name & "]"))
                    End If
                Next fld_idx
                .Close
            End With 'RS_Tbl_src
        End If
    End With 'CurrentDb
    Dim SQL_Seg_Select, SQL_Seq_JoinOn, SQL_Seq_OrderBy, SQL_cmd As String
    SQL_Seg_Select = "SELECT " & "[" & Tbl_src_1_name & "]." & Join(ColSet_src_1, ", [" & Tbl_src_1_name & "].") & ", " & "[" & Tbl_src_2_name & "]." & Join(ColSet_src_2, ", [" & Tbl_src_2_name & "].")
    SQL_Seg_JoinOn = "("
    For Col_Idx = LBound(ColSet_Join_1) To UBound(ColSet_Join_1)
        SQL_Seg_JoinOn = SQL_Seg_JoinOn & "[" & Tbl_src_1_name & "].[" & ColSet_Join_1(Col_Idx) & "] = [" & Tbl_src_2_name & "].[" & ColSet_Join_2(Col_Idx) & "] AND "
    Next Col_Idx
    SQL_Seg_JoinOn = Left(SQL_Seg_JoinOn, Len(SQL_Seg_JoinOn) - 4) & ")"
    SQL_Seg_OrderBy = vbNullString
    If IsNull(ColSet_Order) = False Then
        SQL_Seg_OrderBy = "ORDER BY "
        For Col_Idx = LBound(ColSet_Order) To UBound(ColSet_Order)
            SQL_Seg_OrderBy = SQL_Seg_OrderBy & "[" & Tbl_src_1_name & "].[" & ColSet_Order(Col_Idx) & "], "
        Next Col_Idx
        SQL_Seg_OrderBy = Left(SQL_Seg_OrderBy, Len(SQL_Seg_OrderBy) - 2)
    End If
    SQL_cmd = SQL_Seg_Select & " " & vbCrLf & "INTO [" & Tbl_des_name & "] " & vbCrLf & "FROM [" & Tbl_src_1_name & "] " & JoinCond & " JOIN [" & Tbl_src_2_name & "] " & vbCrLf & "ON " & SQL_Seg_JoinOn & vbCrLf & SQL_Seg_OrderBy & " " & vbCrLf & ";"
    RunSQL_CmdWithoutWarning (SQL_cmd)
Exit_CreateTbl_JoinTwoTbl:
    CreateTbl_JoinTwoTbl = FailedReason
    Exit Function
Err_CreateTbl_JoinTwoTbl:
    FailedReason = Err.Description
    Resume Exit_CreateTbl_JoinTwoTbl
End Function
Public Function CreateTbl_ConcatTbls(Tbl_src_Set As Variant, Tbl_des_name As String, Optional Type_Set As Variant = "") As String
    'Create table which is cancatenated from multiple tables of the same structure
    On Error GoTo Err_CreateTbl_ConcatTbls: Dim FailedReason As String
    If UBound(Tbl_src_Set) < 0 Then FailedReason = "No table in the table set": GoTo Exit_CreateTbl_ConcatTbls
    Dim Tbl_src_name As Variant
    For Each Tbl_src_name In Tbl_src_Set
        If TableExist(Tbl_src_name & "") = False Then
            FailedReason = Tbl_src_name & " does not exist!"
            GoTo Exit_CreateTbl_ConcatTbls
        End If
    Next        'Initialize Tbl_des
    Dim SQL_cmd, SQL_Seq_Type As String, tbl_idx As Integer
    DelTable (Tbl_des_name): Tbl_src_name = Tbl_src_Set(0)
    SQL_cmd = "SELECT " & Chr(34) & "null" & Chr(34) & " AS [Type], " & Tbl_src_name & ".* " & vbCrLf & "INTO " & Tbl_des_name & " " & vbCrLf & "FROM " & Tbl_src_name & " " & vbCrLf & "WHERE 1 = 0 " & vbCrLf & ";"
    RunSQL_CmdWithoutWarning (SQL_cmd)        'Start Append
    For tbl_idx = 0 To UBound(Tbl_src_Set)
        Tbl_src_name = Tbl_src_Set(tbl_idx)
        SQL_Seq_Type = IIf(VarType(Type_Set) > vbArray And Type_Set(tbl_idx) = vbNullString, vbNullString, Chr(34) & Type_Set(tbl_idx) & Chr(34) & " AS [Type], ")
        SQL_cmd = "INSERT INTO " & Tbl_des_name & " " & vbCrLf & "SELECT " & SQL_Seq_Type & "[" & Tbl_src_name & "].* " & vbCrLf & "FROM [" & Tbl_src_name & "] " & vbCrLf & ";"
        RunSQL_CmdWithoutWarning (SQL_cmd)
    Next
    If UBound(Type_Set) < 0 Then
        SQL_cmd = "ALTER TABLE [" & Tbl_des_name & "] " & vbCrLf & "DROP COLUMN [Type]" & vbCrLf & ";"
        RunSQL_CmdWithoutWarning (SQL_cmd)
    End If
Exit_CreateTbl_ConcatTbls:
    CreateTbl_ConcatTbls = FailedReason
    Exit Function
Err_CreateTbl_ConcatTbls:
    FailedReason = Err.Description
    Resume Exit_CreateTbl_ConcatTbls
End Function
Public Function ExecuteSQLiteCmdSet(SQLiteDb_path As String, CmdSet As String) As String
    'Execute SQLite Command Set
    On Error GoTo Err_ExecuteSQLiteCmdSet: Dim FailedReason As String
    If FileExists(SQLiteDb_path) = False Then FailedReason = SQLiteDb_path: GoTo Exit_ExecuteSQLiteCmdSet
    'Create a SQLite Command file, and then parse it into the Python SQLite Command Parser for execution
    Dim SQLiteCmdFile_path, SQLiteCmdLog_path, SQLiteCmdLog_line As String
    Dim iFileNum_SQLiteCmd, iFileNum_SQLiteCmdLog As Integer
    SQLiteCmdFile_path = [CurrentProject].[Path] & "\" & "SQLiteCmd.txt"
    iFileNum_SQLiteCmd = FreeFile()
    If FileExists(SQLiteCmdFile_path) = True Then Kill SQLiteCmdFile_path
    Open SQLiteCmdFile_path For Output As iFileNum_SQLiteCmd
    Print #iFileNum_SQLiteCmd, CmdSet
    Close #iFileNum_SQLiteCmd
    SQLiteCmdLog_path = [CurrentProject].[Path] & "\SQLiteCmd.log"
    If FileExists(SQLiteCmdLog_path) = True Then Kill SQLiteCmdLog_path
    ShellCmd = "python " & [CurrentProject].[Path] & "\SQLiteCmdParser.py " & SQLiteDb_path & " " & SQLiteCmdFile_path & " " & SQLiteCmdLog_path
    Call ShellAndWait(ShellCmd, vbHide)
    If FileExists(SQLiteCmdLog_path) = False Then FailedReason = "SQLiteCmdParser": GoTo Exit_ExecuteSQLiteCmdSet
    iFileNum_SQLiteCmdLog = FreeFile()
    Open SQLiteCmdLog_path For Input As iFileNum_SQLiteCmdLog
    If Not EOF(iFileNum_SQLiteCmdLog) Then Line Input #iFileNum_SQLiteCmdLog, SQLiteCmdLog_line
    If SQLiteCmdLog_line <> "done" Then FailedReason = SQLiteCmdLog_path: GoTo Exit_ExecuteSQLiteCmdSet
    Close iFileNum_SQLiteCmdLog
    Kill SQLiteCmdFile_path
    Kill SQLiteCmdLog_path
Exit_ExecuteSQLiteCmdSet:
    ExecuteSQLiteCmdSet = FailedReason
    Exit Function
Err_ExecuteSQLiteCmdSet:
    Call ShowMsgBox(Err.Description)
    Resume Exit_ExecuteSQLiteCmdSet
End Function
Public Function AppendTblToSQLite(Tbl_src_name As String, Tbl_des_name As String) As String
    'Append Table into a SQLite database
    On Error GoTo Err_AppendTblToSQLite
    Dim FailedReason As String, TempDb_path As String, ShellCmd As String
    If TableExist(Tbl_src_name) = False Then FailedReason = Tbl_src_name: GoTo Exit_AppendTblToSQLite
    If TableExist(Tbl_des_name) = False Then FailedReason = Tbl_des_name: GoTo Exit_AppendTblToSQLite
    TempDb_path = [CurrentProject].[Path] & "\TempDb.mdb"   'Create Db
    If FileExists(TempDb_path) = True Then Kill TempDb_path
    Call CreateDatabase(TempDb_path, dbLangGeneral)
    Dim SQL_cmd As String, SQLiteDb_path As String   'Copy Table into the TempDb
    SQL_cmd = "SELECT * " & vbCrLf & "INTO [" & Tbl_des_name & "]" & vbCrLf & "IN '" & TempDb_path & "'" & vbCrLf & "FROM [" & Tbl_src_name & "] " & vbCrLf & ";"
    RunSQL_CmdWithoutWarning (SQL_cmd)               'Convert TempDb into SQLite
    SQLiteDb_path = [CurrentProject].[Path] & "\TempDb.sqlite"
    If FileExists(SQLiteDb_path) = True Then Kill SQLiteDb_path
    ShellCmd = "java -jar " & [CurrentProject].[Path] & "\mdb-sqlite.jar " & TempDb_path & " " & SQLiteDb_path
    Call ShellAndWait(ShellCmd, vbHide)
    SQL_cmd = "ATTACH """ & SQLiteDb_path & """ AS TempDb;" & vbCrLf & "INSERT INTO [" & Tbl_des_name & "] SELECT * FROM TempDb.[" & Tbl_des_name & "];"
    FailedReason = ExecuteSQLiteCmdSet(GetLinkTblConnInfo(Tbl_des_name, "DATABASE"), SQL_cmd)
    If FailedReason <> vbNullString Then GoTo Exit_AppendTblToSQLite
    Kill SQLiteDb_path
    Kill TempDb_path
Exit_AppendTblToSQLite:
    AppendTblToSQLite = FailedReason
    Exit Function
Err_AppendTblToSQLite:
    Call ShowMsgBox(Err.Description)
    Resume Exit_AppendTblToSQLite
End Function


