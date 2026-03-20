VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QueryBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ============================================================================
' QueryBuilder Class - Fluent SQL Query Builder for VBA
' ============================================================================
'
' PURPOSE:
'   A fluent SQL query builder for VBA (Excel, Access, VB6) that generates
'   parameterized SQL queries with support for SELECT, JOIN, WHERE, GROUP BY,
'   ORDER BY, LIMIT, and OFFSET clauses.
'
' FEATURES:
'   • Fluent API for method chaining
'   • Parameterized queries for SQL injection safety
'   • Support for BETWEEN clauses with dual bindings
'   • Support for IN (...) and grouped WHERE clauses
'   • Bracket-quoted identifiers in generated SQL (e.g., [users].[id])
'   • SQL generation and execution via ADO
'   • SQL-part getter methods (GetWhereClause, GetGroupByClause, etc.)
'   • Compatible with MySQL, SQL Server, and other ODBC databases
'
' DEPENDENCIES:
'   • Microsoft ActiveX Data Objects 2.8 Library (or later)
'   • Scripting Runtime (for Dictionary support in Insert/Update)
'
' BASIC USAGE EXAMPLE:
'   Dim q As QueryBuilder
'   Set q = New QueryBuilder
'   Dim sql As String
'   sql = q.Table("users") _
'          .SelectColumns("*") _
'          .WhereClause("age", ">", 18) _
'          .OrderByClause("name", "ASC") _
'          .Take(10) _
'          .ToSql()
'   ' Returns: SELECT * FROM [users] WHERE [age] > 18 ORDER BY [name] ASC LIMIT 10
'
' BETWEEN EXAMPLE:
'   sql = q.Table("products") _
'          .SelectColumns("*") _
'          .WhereBetween("price", 10, 100) _
'          .ToSql()
'   ' Returns: SELECT * FROM [products] WHERE [price] BETWEEN 10 AND 100
'
' DATABASE QUERY EXAMPLE:
'   Dim conn As Object
'   Set conn = CreateObject("ADODB.Connection")
'   conn.Open "Driver={MySQL};Server=localhost;Database=mydb;..."
'   Set q = New QueryBuilder
'   Dim rs As Object
'   Set rs = q.Table("users") _
'              .SelectColumns("id, name") _
'              .WhereClause("active", "=", 1) _
'              .GetRows(conn)
'   If Not rs.EOF Then
'       ' Process recordset
'   End If
'   rs.Close
'   conn.Close
'
' ============================================================================

Option Explicit

' ADO Command Constants
Private Const adCmdText As Long = 1              ' Text command type
Private Const adParamInput As Long = 1           ' Input parameter
Private Const adVariant As Long = 12             ' Variant data type
Private Const adExecuteNoRecords As Long = &H80  ' No records to return

' Private state variables - Query builder state
Private mTable As String                    ' Table name for query
Private mSelect As String                   ' SELECT clause columns (default "*")
Private mWhere As Collection                ' WHERE conditions array
Private mBindings As Collection             ' Parameter bindings for prepared statements
Private mJoins As Collection                ' JOIN clauses
Private mGroupBy As Collection              ' GROUP BY columns
Private mOrderBy As String                  ' ORDER BY clause
Private mLimit As Variant                   ' LIMIT value (Null if not set)
Private mOffset As Variant                  ' OFFSET value (Null if not set)
Private mUseSqlServerPagination As Boolean  ' True: TOP / OFFSET FETCH, False: LIMIT / OFFSET
Private mPDO As Object                      ' ADO Connection object (late-bound)
Private mOpenWhereGroups As Long            ' Tracks open where-group parentheses


' ============================================================================
' CLASS INITIALIZATION AND CONNECTION
' ============================================================================

' Class_Initialize
' Initializes all collections and state to defaults when a new instance is created
' Called automatically by VBA when: Set q = New QueryBuilder
'
Private Sub Class_Initialize()
    mSelect = "*"
    Set mWhere = New Collection
    Set mBindings = New Collection
    Set mJoins = New Collection
    Set mGroupBy = New Collection
    mOrderBy = vbNullString
    mLimit = Null
    mOffset = Null
    mUseSqlServerPagination = True
    mOpenWhereGroups = 0
End Sub

' SetConnection
' Stores the ADO connection object for use with GetRows(), CountRows(), etc.
'
' Parameters:
'   connection - ADO.Connection object (already opened and connected)
'
' Returns: (None - sets internal reference)
'
' Example:
'   Dim conn As Object
'   Set conn = CreateObject("ADODB.Connection")
'   conn.Open "Provider=MSDASQL;Driver={MySQL ODBC Driver};Server=localhost;Database=mydb;User=root;Password=pass;"
'   q.SetConnection conn
'
Public Sub SetConnection(ByVal connection As Object)
    Set mPDO = connection
End Sub


' ============================================================================
' FOUNDATION METHODS - TABLE AND SELECT
' ============================================================================

' Table
' Sets the table name for the query
'
' Parameters:
'   tblName (String) - The name of the table to query
'
' Returns: QueryBuilder (Self for method chaining)
'
' Example:
'   q.Table("users") -> SELECT * FROM users
'
Public Function Table(ByVal tblName As String) As QueryBuilder
    mTable = tblName
    Set Table = Me
End Function

' SelectColumns
' Sets the SELECT clause columns
'
' Parameters:
'   columns (Variant) - Either a string like "id, name, email" or an array of column names
'
' Returns: QueryBuilder (Self for method chaining)
'
' Example:
'   q.SelectColumns("id, name") or q.SelectColumns(Array("id", "name"))
'
Public Function SelectColumns(ByVal columns As Variant) As QueryBuilder
    If IsArray(columns) Then
        mSelect = Join(columns, ", ")
    Else
        mSelect = CStr(columns)
    End If
    Set SelectColumns = Me
End Function


' ============================================================================
' JOIN AND GROUP BY METHODS
' ============================================================================

' JoinClause
' Adds a JOIN clause to the query
'
' Parameters:
'   tblName (String) - Table name to join with
'   firstColumn (String) - Left-side column (e.g., "users.id")
'   op (String) - Comparison operator (e.g., "=", ">", "<")
'   secondColumn (String) - Right-side column (e.g., "orders.user_id")
'   joinType (String, Optional) - Type of JOIN: "INNER" (default), "LEFT", "RIGHT"
'
' Returns: QueryBuilder (Self for method chaining)
'
' Example:
'   q.JoinClause("cars", "users.id", "=", "cars.user_id", "left")
'
Public Function JoinClause(ByVal tblName As String, ByVal firstColumn As String, ByVal op As String, ByVal secondColumn As String, Optional ByVal joinType As String = "INNER") As QueryBuilder
    mJoins.Add UCase$(joinType) & " JOIN " & QuoteIdentifierPath(tblName) & " ON " & QuoteIdentifierPath(firstColumn) & " " & op & " " & QuoteIdentifierPath(secondColumn)
    Set JoinClause = Me
End Function

' GroupByColumns
' Sets the GROUP BY clause columns
'
' Parameters:
'   columns (Variant) - Either a string like "category, status" or an array of column names
'
' Returns: QueryBuilder (Self for method chaining)
'
' Example:
'   q.GroupByColumns("category") or q.GroupByColumns(Array("category", "status"))
'
Public Function GroupByColumns(ByVal columns As Variant) As QueryBuilder
    Dim i As Long
    Set mGroupBy = New Collection

    If IsArray(columns) Then
        For i = LBound(columns) To UBound(columns)
            mGroupBy.Add CStr(columns(i))
        Next i
    Else
        mGroupBy.Add CStr(columns)
    End If

    Set GroupByColumns = Me
End Function


' ============================================================================
' LIMIT, OFFSET, AND ORDER BY METHODS
' ============================================================================

' Take
' Sets the LIMIT clause for the query
'
' Parameters:
'   limitValue (Long) - Maximum number of rows to return
'
' Returns: QueryBuilder (Self for method chaining)
'
' Example:
'   q.Take(10) -> LIMIT 10
'
Public Function Take(ByVal limitValue As Long) As QueryBuilder
    mLimit = limitValue
    Set Take = Me
End Function

' SkipRows
' Sets the OFFSET clause for the query (for pagination)
'
' Parameters:
'   offsetValue (Long) - Number of rows to skip
'
' Returns: QueryBuilder (Self for method chaining)
'
' Example:
'   q.Take(10).SkipRows(20) -> LIMIT 10 OFFSET 20 (page 3 with 10 items per page)
'
Public Function SkipRows(ByVal offsetValue As Long) As QueryBuilder
    mOffset = offsetValue
    Set SkipRows = Me
End Function

' UseSqlServerPagination
' Enables SQL Server-style pagination output.
'   - limit only: SELECT TOP (n) ...
'   - offset only: ... ORDER BY ... OFFSET n ROWS
'   - offset + limit: ... ORDER BY ... OFFSET n ROWS FETCH NEXT m ROWS ONLY
'
Public Function UseSqlServerPagination(Optional ByVal enabled As Boolean = True) As QueryBuilder
    mUseSqlServerPagination = enabled
    Set UseSqlServerPagination = Me
End Function

' UseStandardPagination
' Restores default LIMIT / OFFSET output.
'
Public Function UseStandardPagination() As QueryBuilder
    mUseSqlServerPagination = False
    Set UseStandardPagination = Me
End Function

' OrderByClause
' Sets the ORDER BY clause
'
' Parameters:
'   col (String) - Column name to sort by
'   direction (String, Optional) - "ASC" (default) or "DESC" for sort direction
'
' Returns: QueryBuilder (Self for method chaining)
'
' Raises: InvalidArgumentException if direction is not "ASC" or "DESC"
'
' Example:
'   q.OrderByClause("name", "ASC") -> ORDER BY name ASC
'
Public Function OrderByClause(ByVal col As String, Optional ByVal direction As String = "ASC") As QueryBuilder
    direction = UCase$(direction)
    If direction <> "ASC" And direction <> "DESC" Then
        Err.Raise vbObjectError + 1001, "QueryBuilder.OrderByClause", "Invalid sorting direction. Only 'ASC' or 'DESC' are allowed."
    End If

    mOrderBy = QuoteIdentifierPath(col) & " " & direction
    Set OrderByClause = Me
End Function


' ============================================================================
' WHERE CLAUSE METHODS
' ============================================================================

' WhereClause
' Adds a WHERE condition to the query
'
' Parameters:
'   col (String) - Column name (e.g., "age", "status")
'   op (String) - Comparison operator (e.g., "=", ">", "<", ">=", "<=", "!=", "<>", "LIKE")
'   val (Variant) - Value to compare against (will be parameterized)
'   boolOp (String, Optional) - Boolean operator ("AND", "OR") - empty for first condition
'
' Returns: QueryBuilder (Self for method chaining)
'
' Example:
'   q.WhereClause("age", ">", 18).AndWhere("status", "=", "active")
'   ' Generates: WHERE age > ? AND status = ?
'
Public Function WhereClause(ByVal col As String, ByVal op As String, ByVal val As Variant, Optional ByVal boolOp As String = vbNullString) As QueryBuilder
    mWhere.Add Array(col, op, val, NormalizeBoolean(boolOp), False, Empty, Empty, "COMPARE")
    mBindings.Add val
    Set WhereClause = Me
End Function

' AndWhere
' Adds a WHERE condition with AND operator
'
' Parameters:
'   col (String) - Column name
'   op (String) - Comparison operator
'   val (Variant) - Value to compare against
'
' Returns: QueryBuilder (Self for method chaining)
'
' Example:
'   q.WhereClause("status", "=", "active").AndWhere("age", ">", 18)
'
Public Function AndWhere(ByVal col As String, ByVal op As String, ByVal val As Variant) As QueryBuilder
    Set AndWhere = WhereClause(col, op, val, "AND")
End Function

' OrWhere
' Adds a WHERE condition with OR operator
'
' Parameters:
'   col (String) - Column name
'   op (String) - Comparison operator
'   val (Variant) - Value to compare against
'
' Returns: QueryBuilder (Self for method chaining)
'
' Example:
'   q.WhereClause("status", "=", "active").OrWhere("status", "=", "pending")
'
Public Function OrWhere(ByVal col As String, ByVal op As String, ByVal val As Variant) As QueryBuilder
    Set OrWhere = WhereClause(col, op, val, "OR")
End Function

' WhereBetween
' Adds a BETWEEN condition to the query
'
' Parameters:
'   col (String) - Column name
'   fromVal (Variant) - Start of range (inclusive)
'   toVal (Variant) - End of range (inclusive)
'   boolOp (String, Optional) - Boolean operator ("AND", "OR") - empty for first condition
'
' Returns: QueryBuilder (Self for method chaining)
'
' Notes:
'   BETWEEN adds TWO parameter bindings to support the range
'
' Example:
'   q.WhereBetween("price", 10, 100)
'   ' Generates: WHERE price BETWEEN ? AND ?
'
Public Function WhereBetween(ByVal col As String, ByVal fromVal As Variant, ByVal toVal As Variant, Optional ByVal boolOp As String = vbNullString) As QueryBuilder
    mWhere.Add Array(col, "BETWEEN", Empty, NormalizeBoolean(boolOp), True, fromVal, toVal, "BETWEEN")
    mBindings.Add fromVal
    mBindings.Add toVal
    Set WhereBetween = Me
End Function

' AndWhereBetween
' Adds a BETWEEN condition with AND operator
'
' Parameters:
'   col (String) - Column name
'   fromVal (Variant) - Start of range
'   toVal (Variant) - End of range
'
' Returns: QueryBuilder (Self for method chaining)
'
' Example:
'   q.WhereClause("status", "=", "active").AndWhereBetween("price", 10, 100)
'
Public Function AndWhereBetween(ByVal col As String, ByVal fromVal As Variant, ByVal toVal As Variant) As QueryBuilder
    Set AndWhereBetween = WhereBetween(col, fromVal, toVal, "AND")
End Function

' OrWhereBetween
' Adds a BETWEEN condition with OR operator
'
' Parameters:
'   col (String) - Column name
'   fromVal (Variant) - Start of range
'   toVal (Variant) - End of range
'
' Returns: QueryBuilder (Self for method chaining)
'
' Example:
'   q.WhereClause("category", "=", "premium").OrWhereBetween("price", 100, 500)
'
Public Function OrWhereBetween(ByVal col As String, ByVal fromVal As Variant, ByVal toVal As Variant) As QueryBuilder
    Set OrWhereBetween = WhereBetween(col, fromVal, toVal, "OR")
End Function

' WhereIn
' Adds an IN (...) condition to the query
'
' Parameters:
'   col (String) - Column name
'   values (Variant) - Array, Collection, or single value
'   boolOp (String, Optional) - Boolean operator ("AND", "OR") - empty for first condition
'
' Returns: QueryBuilder (Self for method chaining)
'
' Example:
'   q.WhereIn("status", Array("active", "pending"))
'   ' Generates: WHERE status IN (?, ?)
'
Public Function WhereIn(ByVal col As String, ByVal values As Variant, Optional ByVal boolOp As String = vbNullString) As QueryBuilder
    Dim normalizedValues As Variant
    Dim i As Long

    normalizedValues = NormalizeInValues(values)
    mWhere.Add Array(col, "IN", normalizedValues, NormalizeBoolean(boolOp), False, Empty, Empty, "IN")

    For i = LBound(normalizedValues) To UBound(normalizedValues)
        mBindings.Add normalizedValues(i)
    Next i

    Set WhereIn = Me
End Function

Public Function AndWhereIn(ByVal col As String, ByVal values As Variant) As QueryBuilder
    Set AndWhereIn = WhereIn(col, values, "AND")
End Function

Public Function OrWhereIn(ByVal col As String, ByVal values As Variant) As QueryBuilder
    Set OrWhereIn = WhereIn(col, values, "OR")
End Function

' WhereGroup
' Starts a grouped where block, allowing conditions to be enclosed in parentheses
'
' Example:
'   q.WhereClause("tenant_id", "=", 10) _
'    .AndWhereGroup() _
'    .WhereClause("status", "=", "active") _
'    .OrWhere("status", "=", "pending") _
'    .EndWhereGroup() _
'    .AndWhere("is_deleted", "=", 0)
'
Public Function WhereGroup(Optional ByVal boolOp As String = vbNullString) As QueryBuilder
    mWhere.Add Array(vbNullString, vbNullString, Empty, NormalizeBoolean(boolOp), False, Empty, Empty, "GROUP_START")
    mOpenWhereGroups = mOpenWhereGroups + 1
    Set WhereGroup = Me
End Function

Public Function AndWhereGroup() As QueryBuilder
    Set AndWhereGroup = WhereGroup("AND")
End Function

Public Function OrWhereGroup() As QueryBuilder
    Set OrWhereGroup = WhereGroup("OR")
End Function

Public Function EndWhereGroup() As QueryBuilder
    If mOpenWhereGroups <= 0 Then
        Err.Raise vbObjectError + 1004, "QueryBuilder.EndWhereGroup", "No open where group to close."
    End If

    mWhere.Add Array(vbNullString, vbNullString, Empty, vbNullString, False, Empty, Empty, "GROUP_END")
    mOpenWhereGroups = mOpenWhereGroups - 1
    Set EndWhereGroup = Me
End Function


' ============================================================================
' SQL GENERATION AND EXECUTION METHODS
' ============================================================================

' BuildSelectSql
' Builds the complete SQL query string with parameterized placeholders
'
' Returns: String - SQL query with ? placeholders for each parameter
'
' Notes:
'   This is an INTERNAL method used by ToSql() and GetRows()
'   Does not interpolate values - parameters are sent separately to the database
'
' Example output:
'   SELECT * FROM [users] WHERE [age] > ? AND [status] = ? LIMIT 10
'
Public Function BuildSelectSql() As String
    Dim sql As String
    sql = GetSelectClause() & " " & GetFromClause()

    If LenB(GetJoinClause()) > 0 Then sql = sql & " " & GetJoinClause()
    If LenB(GetWhereClause()) > 0 Then sql = sql & " " & GetWhereClause()
    If LenB(GetGroupByClause()) > 0 Then sql = sql & " " & GetGroupByClause()
    If LenB(GetOrderByClause()) > 0 Then sql = sql & " " & GetOrderByClause()
    If LenB(GetLimitClause()) > 0 Then sql = sql & " " & GetLimitClause()
    If LenB(GetOffsetClause()) > 0 Then sql = sql & " " & GetOffsetClause()

    BuildSelectSql = sql
End Function

' ============================================================================
' SQL CLAUSE GETTERS
' ============================================================================
'
' Each getter returns its fully-formed clause fragment (e.g., "WHERE [age] > ?")
' or an empty string when no value has been set.
'
' All getters accept an optional includeBindings parameter. When True, any ?
' placeholders in the clause are replaced with their interpolated literal values.
' This is intended for inspection/debugging only - not for database execution.
'
' Only GetWhereClause produces ? placeholders; the parameter is accepted on all
' getters for a consistent API but is a no-op on those that never contain bindings.
'
' Parameters (all getters):
'   includeBindings (Boolean, Optional) - When True, interpolates bound values
'                                         into ? placeholders. Default: False.
'
' Example without bindings (parameterized - safe for execution):
'   q.GetWhereClause()
'   ' Returns: WHERE [status] = ? AND [age] > ?
'
' Example with bindings (interpolated - for inspection only):
'   q.GetWhereClause(includeBindings:=True)
'   ' Returns: WHERE [status] = 'active' AND [age] > 18

' GetSelectClause
' Returns the SELECT clause (e.g., "SELECT [id], [name]")
'
Public Function GetSelectClause(Optional ByVal includeBindings As Boolean = False) As String
    Dim clause As String
    clause = "SELECT "

    If mUseSqlServerPagination Then
        If Not IsNull(mLimit) And IsNull(mOffset) Then
            clause = clause & "TOP (" & CStr(mLimit) & ") "
        End If
    End If

    clause = clause & QuoteColumnList(mSelect)
    GetSelectClause = clause
End Function

' GetFromClause
' Returns the FROM clause (e.g., "FROM [users]")
'
Public Function GetFromClause(Optional ByVal includeBindings As Boolean = False) As String
    GetFromClause = "FROM " & QuoteIdentifierPath(mTable)
End Function

' GetJoinClause
' Returns all JOIN clauses as a single string, or empty string if none set
' (e.g., "INNER JOIN [orders] ON [users].[id] = [orders].[user_id]")
'
Public Function GetJoinClause(Optional ByVal includeBindings As Boolean = False) As String
    If mJoins.Count > 0 Then
        GetJoinClause = JoinCollectionItems(mJoins, " ")
    Else
        GetJoinClause = vbNullString
    End If
End Function

' GetWhereClause
' Returns the WHERE clause (e.g., "WHERE [status] = ? AND [age] > ?")
' When includeBindings is True, ? placeholders are replaced with literal values.
'
Public Function GetWhereClause(Optional ByVal includeBindings As Boolean = False) As String
    Dim clause As String

    If mWhere.Count = 0 Then
        GetWhereClause = vbNullString
        Exit Function
    End If

    clause = "WHERE " & BuildWhereClause()

    If includeBindings Then
        GetWhereClause = InterpolateBindings(clause)
    Else
        GetWhereClause = clause
    End If
End Function

' GetWhereFilter
' Returns the WHERE clause without the "WHERE" (e.g., "[status] = 3 AND [age] > 17"), or empty string if not set
' Can be applied directly to a Form.Filter
'
Public Function GetWhereFilter() As String
    GetWhereFilter = Trim(Replace(GetWhereClause(True), "WHERE", "", 1, 1, vbTextCompare))
End Function

' GetGroupByClause
' Returns the GROUP BY clause (e.g., "GROUP BY [category]"), or empty string if not set
'
Public Function GetGroupByClause(Optional ByVal includeBindings As Boolean = False) As String
    If mGroupBy.Count > 0 Then
        GetGroupByClause = "GROUP BY " & QuoteCollectionIdentifiers(mGroupBy)
    Else
        GetGroupByClause = vbNullString
    End If
End Function

' GetOrderByClause
' Returns the ORDER BY clause (e.g., "ORDER BY [name] ASC"), or empty string if not set
'
Public Function GetOrderByClause(Optional ByVal includeBindings As Boolean = False) As String
    If LenB(mOrderBy) > 0 Then
        GetOrderByClause = "ORDER BY " & mOrderBy
    Else
        GetOrderByClause = vbNullString
    End If
End Function

' GetLimitClause
' Returns the LIMIT clause (e.g., "LIMIT 10"), or empty string if not set
'
Public Function GetLimitClause(Optional ByVal includeBindings As Boolean = False) As String
    If mUseSqlServerPagination Then
        GetLimitClause = vbNullString
        Exit Function
    End If

    If Not IsNull(mLimit) Then
        GetLimitClause = "LIMIT " & CStr(mLimit)
    Else
        GetLimitClause = vbNullString
    End If
End Function

' GetOffsetClause
' Returns the OFFSET clause (e.g., "OFFSET 20"), or empty string if not set
'
Public Function GetOffsetClause(Optional ByVal includeBindings As Boolean = False) As String
    If mUseSqlServerPagination Then
        If IsNull(mOffset) Then
            GetOffsetClause = vbNullString
            Exit Function
        End If

        EnsureOrderByForSqlServerOffset
        GetOffsetClause = "OFFSET " & CStr(mOffset) & " ROWS"

        If Not IsNull(mLimit) Then
            GetOffsetClause = GetOffsetClause & " FETCH NEXT " & CStr(mLimit) & " ROWS ONLY"
        End If
        Exit Function
    End If

    If Not IsNull(mOffset) Then
        GetOffsetClause = "OFFSET " & CStr(mOffset)
    Else
        GetOffsetClause = vbNullString
    End If
End Function

' ToSql
' Returns the complete SQL string with values interpolated (for testing/debugging)
'
' Returns: String - Complete SQL query with all parameters interpolated into the string
'
' WARNING:
'   This is intended for inspection/testing only. Do NOT use this for database execution
'   as it does not properly parameterize queries (SQL injection risk).
'   Use GetRows(), CountRows(), etc. instead for safe database queries.
'
' Example:
'   Debug.Print q.Table("users").WhereClause("age", ">", 18).ToSql()
'   ' Output: SELECT * FROM [users] WHERE [age] > 18
'
Public Function ToSql() As String
    ToSql = InterpolateBindings(BuildSelectSql())
End Function

' GetRows
' Executes the SELECT query and returns an ADO Recordset with results
'
' Parameters:
'   conn (Object) - Open ADO.Connection object
'
' Returns: ADO.Recordset - Open recordset with query results
'
' Notes:
'   Caller is responsible for closing the recordset when done
'   Uses parameterized queries for SQL injection safety
'
' Example:
'   Dim rs As Object
'   Set rs = q.Table("users").SelectColumns("id, name").GetRows(conn)
'   Do Until rs.EOF
'       Debug.Print rs.Fields("name").Value
'       rs.MoveNext
'   Loop
'   rs.Close
'
Public Function GetRows(ByVal conn As Object) As Object
    Dim cmd As Object
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = BuildSelectSql()
    AddBindingsToCommand cmd

    Dim rs As Object
    Set rs = cmd.Execute
    Set GetRows = rs
End Function

' CountRows
' Executes a COUNT(*) query based on current WHERE and JOIN clauses
'
' Parameters:
'   conn (Object) - Open ADO.Connection object
'
' Returns: Long - Count of rows matching the WHERE conditions
'
' Example:
'   Dim totalRows As Long
'   totalRows = q.Table("users").WhereClause("status", "=", "active").CountRows(conn)
'   ' Uses: SELECT COUNT(*) AS [total] FROM [users] WHERE [status] = ?
'
Public Function CountRows(ByVal conn As Object) As Long
    Dim sql As String
    Dim totalValue As Variant
    sql = "SELECT COUNT(*) AS [total] " & GetFromClause()

    If LenB(GetJoinClause()) > 0 Then sql = sql & " " & GetJoinClause()
    If LenB(GetWhereClause()) > 0 Then sql = sql & " " & GetWhereClause()
    If LenB(GetGroupByClause()) > 0 Then sql = sql & " " & GetGroupByClause()

    Dim cmd As Object, rs As Object
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = sql
    AddBindingsToCommand cmd

    Set rs = cmd.Execute
    If Not rs.EOF Then
        totalValue = rs.Fields("total").Value
        If IsNull(totalValue) Or IsEmpty(totalValue) Then
            CountRows = 0
        Else
            CountRows = CLng(totalValue)
        End If
    End If
End Function

' DeleteRows
' Executes a DELETE query based on current WHERE clause
'
' Parameters:
'   conn (Object) - Open ADO.Connection object
'
' Returns: Long - Number of rows deleted
'
' Example:
'   Dim deletedCount As Long
'   deletedCount = q.Table("users").WhereClause("status", "=", "inactive").DeleteRows(conn)
'   MsgBox "Deleted " & deletedCount & " rows"
'   ' Uses: DELETE FROM [users] WHERE [status] = ?
'
Public Function DeleteRows(ByVal conn As Object) As Long
    Dim sql As String
    sql = "DELETE " & GetFromClause()
    If LenB(GetWhereClause()) > 0 Then sql = sql & " " & GetWhereClause()

    Dim cmd As Object, recordsAffected As Long
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = sql
    AddBindingsToCommand cmd

    cmd.Execute recordsAffected, , adExecuteNoRecords
    DeleteRows = recordsAffected
End Function

' InsertRow
' Executes an INSERT query with the provided column values
'
' Parameters:
'   conn (Object) - Open ADO.Connection object
'   data (Object) - Scripting.Dictionary with column names as keys, values as values
'
' Returns: Long - Number of rows inserted (usually 1)
'
' Example:
'   Dim dict As Object
'   Set dict = CreateObject("Scripting.Dictionary")
'   dict.Add "name", "John Doe"
'   dict.Add "email", "john@example.com"
'   dict.Add "age", 30
'   Dim result As Long
'   result = q.Table("users").InsertRow(conn, dict)
'
Public Function InsertRow(ByVal conn As Object, ByVal data As Object) As Long
    ' data is expected to be Scripting.Dictionary (late-bound)
    Dim keys As Variant, i As Long, placeholders As String, columns As String
    keys = data.Keys

    For i = LBound(keys) To UBound(keys)
        columns = columns & IIf(i > LBound(keys), ", ", "") & QuoteIdentifierPath(CStr(keys(i)))
        placeholders = placeholders & IIf(i > LBound(keys), ", ", "") & "?"
    Next i

    Dim cmd As Object, recordsAffected As Long
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "INSERT INTO " & QuoteIdentifierPath(mTable) & " (" & columns & ") VALUES (" & placeholders & ")"

    For i = LBound(keys) To UBound(keys)
        cmd.Parameters.Append cmd.CreateParameter(, adVariant, adParamInput, , data(keys(i)))
    Next i

    cmd.Execute recordsAffected, , adExecuteNoRecords
    InsertRow = recordsAffected
End Function

' UpdateRows
' Executes an UPDATE query with the provided column values and WHERE conditions
'
' Parameters:
'   conn (Object) - Open ADO.Connection object
'   data (Object) - Scripting.Dictionary with column names as keys, new values as values
'
' Returns: Long - Number of rows updated
'
' Notes:
'   WHERE clause conditions must be set via WhereClause/AndWhere/OrWhere before calling UpdateRows
'
' Example:
'   Dim dict As Object
'   Set dict = CreateObject("Scripting.Dictionary")
'   dict.Add "status", "inactive"
'   dict.Add "last_updated", Now()
'   Dim result As Long
'   result = q.Table("users").WhereClause("id", "=", 42).UpdateRows(conn, dict)
'
Public Function UpdateRows(ByVal conn As Object, ByVal data As Object) As Long
    ' data is expected to be Scripting.Dictionary (late-bound)
    Dim keys As Variant, i As Long, setClause As String
    keys = data.Keys

    For i = LBound(keys) To UBound(keys)
        setClause = setClause & IIf(i > LBound(keys), ", ", "") & QuoteIdentifierPath(CStr(keys(i))) & " = ?"
    Next i

    Dim sql As String
    sql = "UPDATE " & QuoteIdentifierPath(mTable) & " SET " & setClause
    If LenB(GetWhereClause()) > 0 Then sql = sql & " " & GetWhereClause()

    Dim cmd As Object, recordsAffected As Long
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = sql

    For i = LBound(keys) To UBound(keys)
        cmd.Parameters.Append cmd.CreateParameter(, adVariant, adParamInput, , data(keys(i)))
    Next i

    For i = 1 To mBindings.Count
        cmd.Parameters.Append cmd.CreateParameter(, adVariant, adParamInput, , mBindings(i))
    Next i

    cmd.Execute recordsAffected, , adExecuteNoRecords
    UpdateRows = recordsAffected
End Function


' ============================================================================
' PRIVATE HELPER METHODS
' ============================================================================

' BuildWhereClause (PRIVATE)
' Constructs the WHERE clause portion of SQL from stored condition array
'
' Returns: String - WHERE clause fragment with ? placeholders (e.g., "age > ? AND status = ?")
'
' Notes:
'   • Handles regular, BETWEEN, IN, and grouped WHERE conditions
'   • BETWEEN conditions generate "column BETWEEN ? AND ?" with two placeholders
'   • IN conditions generate "column IN (?, ?, ...)" based on input size
'   • Properly handles AND/OR boolean operators
'   • Returns empty string if no WHERE conditions exist
'
Private Function BuildWhereClause() As String
    Dim parts As Collection: Set parts = New Collection
    Dim item As Variant
    Dim col As String, op As String, boolOp As String
    Dim clauseType As String
    Dim values As Variant

    For Each item In mWhere
        col = CStr(item(0))
        op = CStr(item(1))
        boolOp = CStr(item(3))
        clauseType = GetWhereClauseType(item)

        Select Case clauseType
            Case "GROUP_START"
                If ShouldIgnoreBooleanPrefix(parts) Then
                    parts.Add "("
                ElseIf LenB(boolOp) > 0 Then
                    parts.Add boolOp & " ("
                Else
                    parts.Add "AND ("
                End If

            Case "GROUP_END"
                parts.Add ")"

            Case "BETWEEN"
                AppendWherePart parts, QuoteIdentifierPath(col) & " " & op & " ? AND ?", boolOp

            Case "IN"
                values = item(2)
                AppendWherePart parts, QuoteIdentifierPath(col) & " IN (" & BuildPlaceholders(values) & ")", boolOp

            Case Else
                AppendWherePart parts, QuoteIdentifierPath(col) & " " & op & " ?", boolOp
        End Select
    Next item

    If mOpenWhereGroups <> 0 Then
        Err.Raise vbObjectError + 1005, "QueryBuilder", "Unclosed where-group detected. Call EndWhereGroup() for each WhereGroup()."
    End If

    BuildWhereClause = JoinCollectionItems(parts, " ")
End Function

Private Sub AppendWherePart(ByRef parts As Collection, ByVal clauseSql As String, ByVal boolOp As String)
    If ShouldIgnoreBooleanPrefix(parts) Then
        parts.Add clauseSql
    ElseIf LenB(boolOp) > 0 Then
        parts.Add boolOp & " " & clauseSql
    Else
        parts.Add clauseSql
    End If
End Sub

Private Function ShouldIgnoreBooleanPrefix(ByVal parts As Collection) As Boolean
    If parts.Count = 0 Then
        ShouldIgnoreBooleanPrefix = True
        Exit Function
    End If

    ShouldIgnoreBooleanPrefix = (Right$(Trim$(CStr(parts(parts.Count))), 1) = "(")
End Function

Private Function GetWhereClauseType(ByVal item As Variant) As String
    If UBound(item) >= 7 Then
        GetWhereClauseType = CStr(item(7))
    ElseIf CBool(item(4)) Then
        GetWhereClauseType = "BETWEEN"
    Else
        GetWhereClauseType = "COMPARE"
    End If
End Function

Private Function NormalizeBoolean(ByVal boolOp As String) As String
    Dim normalized As String
    normalized = UCase$(Trim$(boolOp))

    If LenB(normalized) = 0 Then
        NormalizeBoolean = vbNullString
        Exit Function
    End If

    If normalized <> "AND" And normalized <> "OR" Then
        Err.Raise vbObjectError + 1003, "QueryBuilder", "Invalid boolean operator. Only 'AND' or 'OR' are allowed."
    End If

    NormalizeBoolean = normalized
End Function

Private Function NormalizeInValues(ByVal values As Variant) As Variant
    Dim arr() As Variant
    Dim i As Long

    If IsArray(values) Then
        If (UBound(values) - LBound(values) + 1) <= 0 Then
            Err.Raise vbObjectError + 1002, "QueryBuilder.WhereIn", "WhereIn requires at least one value."
        End If
        NormalizeInValues = values
        Exit Function
    End If

    If IsObject(values) Then
        If TypeName(values) = "Collection" Then
            If values.Count = 0 Then
                Err.Raise vbObjectError + 1002, "QueryBuilder.WhereIn", "WhereIn requires at least one value."
            End If

            ReDim arr(0 To values.Count - 1)
            For i = 1 To values.Count
                arr(i - 1) = values(i)
            Next i
            NormalizeInValues = arr
            Exit Function
        End If
    End If

    ReDim arr(0 To 0)
    arr(0) = values
    NormalizeInValues = arr
End Function

Private Function BuildPlaceholders(ByVal values As Variant) As String
    Dim i As Long
    Dim token As String

    For i = LBound(values) To UBound(values)
        If i > LBound(values) Then token = token & ", "
        token = token & "?"
    Next i

    BuildPlaceholders = token
End Function

Private Sub EnsureOrderByForSqlServerOffset()
    If LenB(mOrderBy) = 0 Then
        Err.Raise vbObjectError + 1006, "QueryBuilder", "SQL Server OFFSET/FETCH requires ORDER BY. Call OrderByClause(...) before SkipRows(...)."
    End If
End Sub

Private Function QuoteColumnList(ByVal columnList As String) As String
    Dim parts() As String
    Dim i As Long
    Dim output As String

    parts = Split(columnList, ",")
    For i = LBound(parts) To UBound(parts)
        If i > LBound(parts) Then output = output & ", "
        output = output & QuoteIdentifierPath(Trim$(parts(i)))
    Next i

    QuoteColumnList = output
End Function

Private Function QuoteCollectionIdentifiers(ByVal items As Collection) As String
    Dim i As Long
    Dim output As String

    For i = 1 To items.Count
        If i > 1 Then output = output & ", "
        output = output & QuoteIdentifierPath(CStr(items(i)))
    Next i

    QuoteCollectionIdentifiers = output
End Function

Private Function QuoteIdentifierPath(ByVal identifier As String) As String
    Dim rawValue As String
    Dim pieces() As String
    Dim i As Long
    Dim quoted As String
    Dim asPos As Long
    Dim lhs As String
    Dim rhs As String
    Dim firstToken As String
    Dim restTokens As String
    Dim spacePos As Long

    rawValue = Trim$(identifier)

    If LenB(rawValue) = 0 Then
        QuoteIdentifierPath = rawValue
        Exit Function
    End If

    If rawValue = "*" Then
        QuoteIdentifierPath = rawValue
        Exit Function
    End If

    ' Keep expressions/functions intact; quote simple identifiers and dotted paths.
    If InStr(rawValue, "(") > 0 Or InStr(rawValue, ")") > 0 Or InStr(rawValue, "?") > 0 Then
        QuoteIdentifierPath = rawValue
        Exit Function
    End If

    asPos = InStr(1, UCase$(rawValue), " AS ", vbTextCompare)
    If asPos > 0 Then
        lhs = Left$(rawValue, asPos - 1)
        rhs = Mid$(rawValue, asPos + 4)
        QuoteIdentifierPath = QuoteIdentifierPath(lhs) & " AS " & QuoteIdentifierPath(rhs)
        Exit Function
    End If

    spacePos = InStr(1, rawValue, " ", vbBinaryCompare)
    If spacePos > 0 Then
        firstToken = Left$(rawValue, spacePos - 1)
        restTokens = Mid$(rawValue, spacePos + 1)
        QuoteIdentifierPath = QuoteIdentifierPath(firstToken) & " " & QuoteIdentifierPath(restTokens)
        Exit Function
    End If

    pieces = Split(rawValue, ".")
    For i = LBound(pieces) To UBound(pieces)
        If i > LBound(pieces) Then quoted = quoted & "."
        quoted = quoted & QuoteIdentifierSegment(Trim$(pieces(i)))
    Next i

    QuoteIdentifierPath = quoted
End Function

Private Function QuoteIdentifierSegment(ByVal segment As String) As String
    Dim cleaned As String
    cleaned = Trim$(segment)

    If LenB(cleaned) = 0 Then
        QuoteIdentifierSegment = cleaned
        Exit Function
    End If

    If cleaned = "*" Then
        QuoteIdentifierSegment = cleaned
        Exit Function
    End If

    If Left$(cleaned, 1) = "[" And Right$(cleaned, 1) = "]" Then
        QuoteIdentifierSegment = cleaned
        Exit Function
    End If

    QuoteIdentifierSegment = "[" & Replace(cleaned, "]", "]]" ) & "]"
End Function

' AddBindingsToCommand (PRIVATE)
' Adds all parameter bindings to an ADO Command object as input parameters
'
' Parameters:
'   cmd (Object) - ADO.Command object to receive parameters
'
' Notes:
'   • Called automatically before executing any database query
'   • Parameters are added in the same order as placeholders in SQL
'   • Ensures proper SQL injection prevention
'
Private Sub AddBindingsToCommand(ByVal cmd As Object)
    Dim i As Long
    For i = 1 To mBindings.Count
        cmd.Parameters.Append cmd.CreateParameter(, adVariant, adParamInput, , mBindings(i))
    Next i
End Sub

' JoinCollectionItems (PRIVATE)
' Joins all items in a Collection into a delimited string
'
' Parameters:
'   items (Collection) - Collection of items to join
'   delimiter (String) - String to place between items
'
' Returns: String - All items joined by delimiter
'
' Example:
'   JoinCollectionItems(coll, ", ") with coll = {id, name, email}
'   Returns: "id, name, email"
'
Private Function JoinCollectionItems(ByVal items As Collection, ByVal delimiter As String) As String
    Dim i As Long, s As String
    For i = 1 To items.Count
        If i > 1 Then s = s & delimiter
        s = s & CStr(items(i))
    Next i
    JoinCollectionItems = s
End Function

' InterpolateBindings (PRIVATE)
' Replaces all ? placeholders in a SQL string with their SqlLiteral values
' from mBindings, in positional order.
'
' Parameters:
'   sql (String) - SQL string containing ? placeholders
'
' Returns: String - SQL with literals substituted for each ?
'
' Notes:
'   • Shared by ToSql() and GetWhereClause(includeBindings:=True)
'   • For inspection/debugging only - do NOT use for database execution
'
Private Function InterpolateBindings(ByVal sql As String) As String
    Dim i As Long
    For i = 1 To mBindings.Count
        sql = ReplaceFirst(sql, "?", SqlLiteral(mBindings(i)))
    Next i
    InterpolateBindings = sql
End Function

' ReplaceFirst (PRIVATE)
' Replaces the first occurrence of a search string with replacement text
'
' Parameters:
'   text (String) - String to search within
'   findText (String) - Text to find (e.g., "?")
'   replaceText (String) - Text to replace with (e.g., a literal value)
'
' Returns: String - Modified string with first occurrence replaced
'
' Notes:
'   • Used by ToSql() to interpolate parameter values for inspection
'   • Only replaces FIRST occurrence, preserving subsequent placeholders
'
Private Function ReplaceFirst(ByVal text As String, ByVal findText As String, ByVal replaceText As String) As String
    Dim p As Long
    p = InStr(1, text, findText, vbBinaryCompare)
    If p = 0 Then
        ReplaceFirst = text
    Else
        ReplaceFirst = Left$(text, p - 1) & replaceText & Mid$(text, p + Len(findText))
    End If
End Function

' SqlLiteral (PRIVATE)
' Converts a value to a SQL literal string for inspection/debugging
'
' Parameters:
'   v (Variant) - Value to convert
'
' Returns: String - SQL literal representation:
'   • NULL for Null values
'   • Numeric string for numeric values (no quotes, via VBA built-in IsNumeric)
'   • Quoted and escaped string for text values
'
' Example:
'   SqlLiteral(42) returns "42"
'   SqlLiteral("John's") returns "'John''s'" (escaped for SQL)
'   SqlLiteral(Null) returns "NULL"
'
Private Function SqlLiteral(ByVal v As Variant) As String
    If IsNull(v) Then
        SqlLiteral = "NULL"
    ElseIf IsNumeric(v) Then
        SqlLiteral = CStr(v)
    Else
        SqlLiteral = "'" & Replace(CStr(v), "'", "''") & "'"
    End If
End Function

