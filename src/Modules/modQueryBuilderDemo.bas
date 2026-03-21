Attribute VB_Name = "modQueryBuilderDemo"
Option Explicit

' Class: QueryBuilderDemo
' Demo module showing QueryBuilder usage
'
' Requires: Microsoft ActiveX Data Objects 2.8 Library (or later)
'
' <View Source: https://github.com/scottsweep/VBA-FluentQueryBuilder/blob/main/src/Modules/modQueryBuilderDemo.bas>
'
' Function: DemoQueryBuilder
' Demonstrates various QueryBuilder features by constructing different SQL queries and printing them to the Immediate Window
Public Sub DemoQueryBuilder()
    ' Example 1: Basic SELECT with ToSql (no DB connection needed)
    Dim q As QueryBuilder
    Set q = New QueryBuilder

    Dim sql As String
    sql = q.Table("users") _
           .SelectColumns("users.*") _
           .JoinClause("cars", "users.id", "=", "cars.user_id", "left") _
           .WhereClause("cars.color", "=", "blue") _
           .AndWhere("users.active", "=", 1) _
           .GroupByColumns("users.id") _
           .OrderByClause("users.id", "asc") _
           .Take(10) _
           .SkipRows(20) _
           .ToSql()

    Debug.Print "Example 1 - Complex Query:"
    Debug.Print sql
    ' Output (quoted): SELECT [users].* FROM [users] LEFT JOIN [cars] ON [users].[id] = [cars].[user_id]
    '                  WHERE [cars].[color] = 'blue' AND [users].[active] = 1
    '                  GROUP BY [users].[id] ORDER BY [users].[id] ASC OFFSET 20 ROWS FETCH NEXT 10 ROWS ONLY
    Debug.Print vbCrLf

    ' Example 2: BETWEEN clause
    Set q = New QueryBuilder
    sql = q.Table("products") _
           .SelectColumns("*") _
           .WhereBetween("price", 10, 100) _
           .ToSql()

    Debug.Print "Example 2 - BETWEEN:"
    Debug.Print sql
    ' Output (quoted): SELECT * FROM [products] WHERE [price] BETWEEN 10 AND 100
    Debug.Print vbCrLf

    ' Example 3: AndWhereBetween
    Set q = New QueryBuilder
    sql = q.Table("products") _
           .SelectColumns("*") _
           .WhereClause("status", "=", "active") _
           .AndWhereBetween("price", 10, 100) _
           .AndWhere("category", "=", "electronics") _
           .ToSql()

    Debug.Print "Example 3 - AndWhereBetween:"
    Debug.Print sql
    ' Output (quoted): SELECT * FROM [products] WHERE [status] = 'active' AND [price] BETWEEN 10 AND 100 AND [category] = 'electronics'
    Debug.Print vbCrLf

    ' Example 4: OrWhereBetween
    Set q = New QueryBuilder
    sql = q.Table("orders") _
           .SelectColumns("id, customer, amount") _
           .WhereClause("status", "=", "pending") _
           .OrWhereBetween("amount", 100, 500) _
           .ToSql()

    Debug.Print "Example 4 - OrWhereBetween:"
    Debug.Print sql
    ' Output (quoted): SELECT [id], [customer], [amount] FROM [orders] WHERE [status] = 'pending' OR [amount] BETWEEN 100 AND 500
    Debug.Print vbCrLf

    ' Example 5: Multiple JOINs and ORDER BY
    Set q = New QueryBuilder
    sql = q.Table("users") _
           .SelectColumns("users.id, users.name, orders.total") _
           .JoinClause("orders", "users.id", "=", "orders.user_id") _
           .JoinClause("products", "orders.product_id", "=", "products.id", "left") _
           .WhereClause("orders.status", "=", "completed") _
           .OrderByClause("orders.total", "desc") _
           .Take(5) _
           .ToSql()

    Debug.Print "Example 5 - Multiple Joins:"
    Debug.Print sql
    ' Output (quoted): SELECT TOP (5) [users].[id], [users].[name], [orders].[total] FROM [users]
    '                  INNER JOIN [orders] ON [users].[id] = [orders].[user_id]
    '                  LEFT JOIN [products] ON [orders].[product_id] = [products].[id]
    '                  WHERE [orders].[status] = 'completed' ORDER BY [orders].[total] DESC
    Debug.Print vbCrLf

    ' Example 6: WhereIn
    Set q = New QueryBuilder
    sql = q.Table("users") _
           .SelectColumns("id, name, role") _
           .WhereIn("role", Array("admin", "manager", "editor")) _
           .AndWhere("active", "=", 1) _
           .ToSql()

    Debug.Print "Example 6 - WhereIn:"
    Debug.Print sql
    ' Output (quoted): SELECT [id], [name], [role] FROM [users] WHERE [role] IN ('admin', 'manager', 'editor') AND [active] = 1
    Debug.Print vbCrLf

    ' Example 7: WhereGroup + chaining before and after group
    Set q = New QueryBuilder
    sql = q.Table("orders") _
           .SelectColumns("id, customer_id, status, total") _
           .WhereClause("tenant_id", "=", 42) _
           .AndWhereGroup() _
               .WhereClause("status", "=", "pending") _
               .OrWhere("status", "=", "processing") _
               .OrWhere("status", "=", "backorder") _
           .EndWhereGroup() _
           .AndWhereBetween("total", 100, 1000) _
           .ToSql()

    Debug.Print "Example 7 - WhereGroup:"
    Debug.Print sql
    ' Output (quoted): SELECT [id], [customer_id], [status], [total] FROM [orders]
    '                  WHERE [tenant_id] = 42 AND ( [status] = 'pending' OR [status] = 'processing' OR [status] = 'backorder' )
    '                  AND [total] BETWEEN 100 AND 1000
    Debug.Print vbCrLf

    ' Example 8: Retrieve SQL parts independently
    Set q = New QueryBuilder
    q.Table("users") _
     .SelectColumns("id, name, created_at") _
     .WhereClause("active", "=", 1) _
     .AndWhereIn("role", Array("admin", "manager")) _
     .GroupByColumns("role") _
     .OrderByClause("created_at", "desc") _
     .Take(25) _
     .SkipRows (50)

    Debug.Print "Example 8 - SQL Parts (parameterized):"
    ' Expected parts are fully-formed clauses (or empty string if unset), with quoted identifiers.
    Debug.Print q.GetSelectClause()
    ' Returns: SELECT [id], [name], [created_at]
    Debug.Print q.GetFromClause()
    ' Returns: FROM [users]
    Debug.Print q.GetWhereClause()
    ' Returns: WHERE [active] = ? AND [role] IN (?, ?)
    Debug.Print q.GetGroupByClause()
    ' Returns: GROUP BY [role]
    Debug.Print q.GetOrderByClause()
    ' Returns: ORDER BY [created_at] DESC
    Debug.Print q.GetLimitClause()
    ' Returns: (empty string in SQL Server mode)
    Debug.Print q.GetOffsetClause()
    ' Returns: OFFSET 50 ROWS FETCH NEXT 25 ROWS ONLY
    Debug.Print vbCrLf

    ' Example 8b: Same getters with includeBindings:=True
    Debug.Print "Example 8b - SQL Parts (with bindings interpolated):"
    Debug.Print q.GetWhereClause(includeBindings:=True)
    ' Returns: WHERE [active] = 1 AND [role] IN ('admin', 'manager')
    Debug.Print vbCrLf

    ' Example 9: SQL Server pagination demo (TOP and OFFSET/FETCH)
    Set q = New QueryBuilder
    sql = q.Table("users") _
           .UseSqlServerPagination() _
           .SelectColumns("id, name") _
           .OrderByClause("id", "asc") _
           .Take(10) _
           .ToSql()

    Debug.Print "Example 9a - SQL Server TOP:"
    Debug.Print sql
    ' Output: SELECT TOP (10) [id], [name] FROM [users] ORDER BY [id] ASC
    Debug.Print vbCrLf

    Set q = New QueryBuilder
    sql = q.Table("users") _
           .UseSqlServerPagination() _
           .SelectColumns("id, name") _
           .OrderByClause("id", "asc") _
           .SkipRows(20) _
           .Take(10) _
           .ToSql()

    Debug.Print "Example 9b - SQL Server OFFSET/FETCH:"
    Debug.Print sql
    ' Output: SELECT [id], [name] FROM [users] ORDER BY [id] ASC OFFSET 20 ROWS FETCH NEXT 10 ROWS ONLY
    Debug.Print vbCrLf

    ' Example 10: Switch back to standard LIMIT/OFFSET mode
    Set q = New QueryBuilder
    sql = q.Table("users") _
           .UseStandardPagination() _
           .SelectColumns("id, name") _
           .OrderByClause("id", "asc") _
           .Take(10) _
           .SkipRows(20) _
           .ToSql()

    Debug.Print "Example 10 - Standard LIMIT/OFFSET:"
    Debug.Print sql
    ' Output: SELECT [id], [name] FROM [users] ORDER BY [id] ASC LIMIT 10 OFFSET 20
    Debug.Print vbCrLf

    ' *** To use with actual database connection: ***
    ' Dim conn As Object
    ' Set conn = CreateObject("ADODB.Connection")
    ' conn.Open "Provider=MSDASQL;Driver={MySQL ODBC Driver};Server=localhost;Database=mydb;User=root;Password=pass;"
    '
    ' Set q = New QueryBuilder
    ' q.SetConnection conn
    ' Dim rs As Object
    ' Set rs = q.Table("users").SelectColumns("*").WhereClause("active", "=", 1).GetRows(conn)
    ' If Not rs.EOF Then
    '     Do Until rs.EOF
    '         Debug.Print rs.Fields("name").Value
    '         rs.MoveNext
    '     Loop
    ' End If
    ' rs.Close
    ' conn.Close

    MsgBox "QueryBuilder demo completed. Check Immediate Window (Ctrl+G) for output.", vbInformation
End Sub



