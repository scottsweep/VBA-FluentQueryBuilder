[![Lint VBA](https://github.com/scottsweep/VBA-FluentQueryBuilder/actions/workflows/lint_vba.yml/badge.svg?branch=main)](https://github.com/scottsweep/VBA-FluentQueryBuilder/actions/workflows/lint_vba.yml)
<a href="https://scottsweep.github.io/VBA-FluentQueryBuilder/#VisualBasicClass:QueryBuilder"><img src="https://img.shields.io/badge/code-documented-green.svg"/></a>

# QueryBuilder

Fluent SQL query builder for VBA (Excel, Access, VB6), with parameterized ADO execution.

QueryBuilder helps you build SQL via chainable methods instead of manual string concatenation, while still allowing full SQL inspection through clause getters.

## Features

- Fluent API (`Table`, `SelectColumns`, `WhereClause`, `JoinClause`, ...)
- Parameterized query execution via ADO (`GetRows`, `CountRows`, `InsertRow`, `UpdateRows`, `DeleteRows`)
- `WHERE` helpers: `AND`/`OR`, `BETWEEN`, `IN (...)`, grouped parentheses
- Bracket-quoted identifiers in generated SQL (e.g. `[users].[id]`)
- SQL clause getters (`GetWhereClause`, `GetGroupByClause`, `GetOrderByClause`, ...)
- Optional binding interpolation for inspection (`includeBindings:=True`)

## Requirements

- VBA host: Excel, Access, or VB6
- ADO library reference (Microsoft ActiveX Data Objects 2.8+)
- Optional: `Scripting.Dictionary` (for `InsertRow` / `UpdateRows`)

## Installation

1. Open the VBA editor (`Alt+F11`)
2. Right-click your project -> **Import File...**
3. Import:
   - `QueryBuilder.cls`
   - (optional) `modQueryBuilderDemo.bas` demo
4. In **Tools -> References**, enable:
   - **Microsoft ActiveX Data Objects 2.8 Library** (or newer)
   - Optional: **Microsoft Scripting Runtime**

## Quick Start

```vb
Dim q As QueryBuilder
Set q = New QueryBuilder

Dim sql As String
sql = q.Table("users") _
       .SelectColumns("id, name") _
       .WhereClause("active", "=", 1) _
       .OrderByClause("name", "asc") _
       .Take(10) _
       .SkipRows(20) _
       .ToSql()

Debug.Print sql
' SELECT [id], [name] FROM [users] WHERE [active] = 1 ORDER BY [name] ASC LIMIT 10 OFFSET 20
```

## SQL Server Pagination (TOP / OFFSET FETCH)

Enable SQL Server mode when targeting SQL Server pagination syntax. Enabled by default.

```vb
Dim q As QueryBuilder
Set q = New QueryBuilder

' TOP (limit-only)
Debug.Print q.Table("users") _
             .UseSqlServerPagination() _
             .SelectColumns("id, name") _
             .Take(10) _
             .ToSql()
' SELECT TOP (10) [id], [name] FROM [users]

' OFFSET ... FETCH NEXT (offset + limit)
Set q = New QueryBuilder
Debug.Print q.Table("users") _
             .UseSqlServerPagination() _
             .SelectColumns("id, name") _
             .OrderByClause("id", "asc") _
             .SkipRows(20) _
             .Take(10) _
             .ToSql()
' SELECT [id], [name] FROM [users] ORDER BY [id] ASC OFFSET 20 ROWS FETCH NEXT 10 ROWS ONLY
```

> In SQL Server mode, `SkipRows(...)` requires `OrderByClause(...)`.

## Standard Pagination (TOP / OFFSET FETCH)

Enable Standard mode when targeting Standard pagination syntax (i.e, Postgres, Mariadb, etc.).

```vb
Dim q As QueryBuilder
Set q = New QueryBuilder

' LIMIT / OFFSET
Set q = New QueryBuilder
Debug.Print q.Table("users") _
             .UseStandardPagination() _
             .SelectColumns("id, name") _
             .OrderByClause("id", "asc") _
             .SkipRows(20) _
             .Take(10) _
             .ToSql()
' SELECT [id], [name] FROM [users] ORDER BY [id] ASC LIMIT 10 OFFSET 20
```

## WHERE Examples

### Basic AND/OR

```vb
sql = q.Table("users") _
       .SelectColumns("*") _
       .WhereClause("status", "=", "active") _
       .AndWhere("age", ">=", 18) _
       .OrWhere("is_admin", "=", 1) _
       .ToSql()
' SELECT * FROM [users] WHERE [status] = 'active' AND [age] >= 18 OR [is_admin] = 1
```

### BETWEEN

```vb
sql = q.Table("products") _
       .SelectColumns("*") _
       .WhereBetween("price", 10, 100) _
       .ToSql()
' SELECT * FROM [products] WHERE [price] BETWEEN 10 AND 100
```

### IN (...)

```vb
sql = q.Table("users") _
       .SelectColumns("id, name, role") _
       .WhereIn("role", Array("admin", "manager", "editor")) _
       .AndWhere("active", "=", 1) _
       .ToSql()
' SELECT [id], [name], [role] FROM [users] WHERE [role] IN ('admin', 'manager', 'editor') AND [active] = 1       
```

### Grouped conditions

```vb
sql = q.Table("orders") _
       .SelectColumns("id, status, total") _
       .WhereClause("tenant_id", "=", 42) _
       .AndWhereGroup() _
           .WhereClause("status", "=", "pending") _
           .OrWhere("status", "=", "processing") _
           .OrWhere("status", "=", "backorder") _
       .EndWhereGroup() _
       .AndWhereBetween("total", 100, 1000) _
       .ToSql()
' SELECT [id], [status], [total] FROM [orders] WHERE [tenant_id] = 42 AND ( [status] = 'pending' OR [status] = 'processing' OR [status] = 'backorder' ) AND [total] BETWEEN 100 AND 1000       
```

## Clause Getters

Use getters when you need partial SQL output.

```vb
Set q = New QueryBuilder
q.Table("users") _
 .SelectColumns("id, name, created_at") _
 .WhereClause("active", "=", 1) _
 .AndWhereIn("role", Array("admin", "manager")) _
 .GroupByColumns("role") _
 .OrderByClause("created_at", "desc") _
 .Take(25) _
 .SkipRows(50)

Debug.Print q.GetSelectClause()
Debug.Print q.GetFromClause()
Debug.Print q.GetWhereClause()                       ' parameterized
Debug.Print q.GetWhereClause(includeBindings:=True)  ' interpolated
Debug.Print q.GetWhereFilter()                       ' no WHERE keyword (form filters)
Debug.Print q.GetGroupByClause()
Debug.Print q.GetOrderByClause()
Debug.Print q.GetLimitClause()
Debug.Print q.GetOffsetClause()
```

> `includeBindings:=True` is intended for debugging/inspection only.

## Execute SQL via ADO

```vb
Dim conn As Object
Set conn = CreateObject("ADODB.Connection")
conn.Open "Provider=MSDASQL;Driver={MySQL ODBC Driver};Server=localhost;Database=mydb;User=root;Password=pass;"

Dim q As QueryBuilder
Set q = New QueryBuilder

' SELECT
Dim rs As Object
Set rs = q.Table("users") _
          .SelectColumns("id, name") _
          .WhereClause("active", "=", 1) _
          .GetRows(conn)

Do Until rs.EOF
    Debug.Print rs.Fields("id").Value, rs.Fields("name").Value
    rs.MoveNext
Loop
rs.Close

' COUNT
Set q = New QueryBuilder
Debug.Print q.Table("users").WhereClause("active", "=", 1).CountRows(conn)

' INSERT
Dim rowData As Object
Set rowData = CreateObject("Scripting.Dictionary")
rowData.Add "name", "Jane Doe"
rowData.Add "email", "jane@example.com"
rowData.Add "active", 1

Set q = New QueryBuilder
Call q.Table("users").InsertRow(conn, rowData)

' UPDATE
Dim updateData As Object
Set updateData = CreateObject("Scripting.Dictionary")
updateData.Add "active", 0

Set q = New QueryBuilder
Call q.Table("users").WhereClause("id", "=", 7).UpdateRows(conn, updateData)

' DELETE
Set q = New QueryBuilder
Call q.Table("users").WhereClause("id", "=", 7).DeleteRows(conn)

conn.Close
```

## Compatibility

`QueryBuilder.cls` supports two pagination styles:

- **SQL Server mode (default):** `TOP` and `OFFSET ... FETCH NEXT`
- **Standard mode:** `LIMIT` / `OFFSET` (enable with `UseStandardPagination()`)

- **SQL Server**
  - Bracket quoting is native (`[table]`, `[column]`).
  - Use `UseSqlServerPagination()` to emit SQL Server-native paging syntax.
- **MySQL / MariaDB**
  - `LIMIT`/`OFFSET` works.
  - Bracket quoting is not native by default (backticks are typical), so adjust quoting strategy if your driver/server rejects brackets.
- **PostgreSQL**
  - `LIMIT`/`OFFSET` works.
  - Identifiers use double quotes, not brackets, so bracket quoting may need adaptation.
- **Microsoft Access (Jet/ACE)**
  - Bracket quoting is native.
  - `LIMIT`/`OFFSET` is not supported (use `TOP` and Access-specific paging patterns).

If your target engine differs, you can still use the fluent filtering/join logic and adapt final SQL syntax where needed.

## Notes

- Identifier quoting is enabled by design (`[table]`, `[table].[column]`).
- Builder instances are mutable. Use a new instance for a new query.
- Pagination syntax depends on mode + engine (`LIMIT/OFFSET` or `TOP/OFFSET FETCH`).
- Grouped conditions must be balanced (`WhereGroup` and `EndWhereGroup`).

## Included Files

- `QueryBuilder.cls` - core class
- `modQueryBuilderDemo.bas` - demo examples

