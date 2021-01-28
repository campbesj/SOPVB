Imports System
Imports System.Collections.Generic
Imports System.Collections
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient


Public Class clsGenericQuery


#Region "Declarations"
    Private myDataSetField As DataSet = New DataSet()

    Public Property MyDataSet As DataSet
        Get
            Return myDataSetField
        End Get
        Set(ByVal value As DataSet)
            myDataSetField = value
        End Set
    End Property

    Private myDataTableField As DataTable = New DataTable()

    Public Property MyDataTable As DataTable
        Get
            Return myDataTableField
        End Get
        Set(ByVal value As DataTable)
            myDataTableField = value
        End Set
    End Property

    Private myDataTablesField As List(Of DataTable) = New List(Of DataTable)()

    Public Property MyDataTables As List(Of DataTable)
        Get
            Return myDataTablesField
        End Get
        Set(ByVal value As List(Of DataTable))
            myDataTablesField = value
        End Set
    End Property

    Private myDataAdapterField As SqlDataAdapter

    Public Property MyDataAdapter As SqlDataAdapter
        Get
            Return myDataAdapterField
        End Get
        Set(ByVal value As SqlDataAdapter)
            myDataAdapterField = value
        End Set
    End Property

    Private myDataAdaptersField As List(Of SqlDataAdapter) = New List(Of SqlDataAdapter)()

    Public Property MyDataAdapters As List(Of SqlDataAdapter)
        Get
            Return myDataAdaptersField
        End Get
        Set(ByVal value As List(Of SqlDataAdapter))
            myDataAdaptersField = value
        End Set
    End Property

    Private myRowField As DataRow

    Public Property MyRow As DataRow
        Get
            Return myRowField
        End Get
        Set(ByVal value As DataRow)
            myRowField = value
        End Set
    End Property

    Private myRowsField As List(Of DataRow) = New List(Of DataRow)()

    Public Property MyRows As List(Of DataRow)
        Get
            Return myRowsField
        End Get
        Set(ByVal value As List(Of DataRow))
            myRowsField = value
        End Set
    End Property

    Public IsNew As Boolean = False
    Public IsNews As List(Of Boolean) = New List(Of Boolean)()
    Public MyCmd As SqlCommandBuilder
    Public MyCmds As List(Of SqlCommandBuilder) = New List(Of SqlCommandBuilder)()
    Private myConField As SqlConnection

    Public Property MyCon As SqlConnection
        Get
            Return myConField
        End Get
        Set(ByVal value As SqlConnection)
            myConField = value
        End Set
    End Property

    Public id As Integer = 0
#End Region

#Region "Constructors"

    ''' <summary>
    ''' Constructor for using parameters
    ''' </summary>
    ''' <paramname="Query">The query string</param>
    ''' <paramname="Params">A hashtable that contains the name and values of the parameters used by the query</param>
    ''' <paramname="ConnectionString">ConnectionString</param>
    Public Sub New(ByVal Query As String, ByVal Params As Hashtable, ByVal ConnectionString As String)
        MyDataSet.Tables.Clear()
        MyCon = New SqlConnection()
        MyCon.ConnectionString = ConnectionString
        MyCon.Open()
        MyDataAdapter = New SqlDataAdapter(Query, MyCon)
        MyCmd = New SqlCommandBuilder(MyDataAdapter)

        For Each o In Params.Keys
            Dim keyname As String = o.ToString()

            If Not keyname.StartsWith("@") Then
                keyname = "@" & keyname
            End If

            MyDataAdapter.SelectCommand.Parameters.AddWithValue(keyname, Params(keyname))
            'this.MyDataAdapter.UpdateCommand.Parameters.AddWithValue(keyname, QueryAndParams[keyname]);

        Next

        MyDataAdapter.Fill(MyDataSet, "analysis")
        MyDataTable = MyDataSet.Tables(0)
        MyCon.Close()

        If MyDataTable.Rows.Count = 0 Then
            IsNew = True
            MyRow = MyDataTable.NewRow()
            MyDataTable.Rows.Add(MyRow)
        Else
            IsNew = False
            MyRow = MyDataTable.Rows(0)
        End If

        If MyDataTable.Columns.Contains("touched") Then
            'this.MyRow["touched"] = clsGlobals.touched + "|" + DateTime.Now.ToString();
        End If
    End Sub 'constructor

    Public Sub New(ByVal Query As String, ByVal ConnectionString As String)
        If Query.Contains(";") Then
            Return
        End If

        MyDataSet.Tables.Clear()
        MyCon = New SqlConnection()
        MyCon.ConnectionString = ConnectionString
        MyCon.Open()
        MyDataTable.Clear()
        MyDataTable.Rows.Clear()
        MyDataTable.Columns.Clear()
        MyDataAdapter = New SqlDataAdapter(Query, MyCon)
        MyDataAdapter.SelectCommand.CommandTimeout = 3000
        MyDataAdapter.Fill(MyDataSet, "analysis")
        MyDataTable = MyDataSet.Tables(0)
        MyCmd = New SqlCommandBuilder(MyDataAdapter)
        MyCon.Close()

        If MyDataTable.Rows.Count = 0 Then
            IsNew = True
            MyRow = MyDataTable.NewRow()
            MyDataTable.Rows.Add(MyRow)
        Else
            IsNew = False
            MyRow = MyDataTable.Rows(0)
        End If

        If MyDataTable.Columns.Contains("touched") Then
            'this.MyRow["touched"] = clsGlobals.touched + "|" + DateTime.Now.ToString();
        End If
    End Sub 'constructor


    Public Sub New(ByVal Query As String, ByVal ConnectionString As String, ByVal WillInsert As Boolean)
        If Query.Contains(";") Then
            Return
        End If

        MyDataSet.Tables.Clear()
        MyCon = New SqlConnection()
        MyCon.ConnectionString = ConnectionString
        MyCon.Open()
        MyDataAdapter = New SqlDataAdapter(Query, MyCon)
        MyCmd = New SqlCommandBuilder(MyDataAdapter)
        AddHandler MyDataAdapter.RowUpdated, New SqlRowUpdatedEventHandler(AddressOf My_OnRowUpdate)
        MyDataAdapter.DeleteCommand = CType(MyCmd.GetDeleteCommand().Clone(), SqlCommand)
        MyDataAdapter.UpdateCommand = CType(MyCmd.GetUpdateCommand().Clone(), SqlCommand)

        ' now we modify the INSERT command, first we clone it and then modify
        Dim cmd As SqlCommand = CType(MyCmd.GetInsertCommand().Clone(), SqlCommand)

        ' adds the call to SCOPE_IDENTITY                                      
        cmd.CommandText += " SET @NEWID = SCOPE_IDENTITY()"
        ' the SET command writes to an output parameter "@ID"
        Dim parm As SqlParameter = New SqlParameter()
        parm.Direction = ParameterDirection.Output
        parm.Size = 4
        parm.SqlDbType = SqlDbType.Int
        parm.ParameterName = "@NEWID"
        parm.DbType = DbType.Int32

        ' adds parameter to command
        cmd.Parameters.Add(parm)

        ' adds our customized insert command to DataAdapter
        MyDataAdapter.InsertCommand = cmd

        ' CommandBuilder needs to be disposed otherwise 
        ' it still tries to generate its default INSERT command 
        MyCmd.Dispose()
        MyDataAdapter.Fill(MyDataSet, "analysis")
        MyDataTable = MyDataSet.Tables(0)
        MyCon.Close()

        If MyDataTable.Rows.Count = 0 Then
            IsNew = True
            MyRow = MyDataTable.NewRow()
            MyDataTable.Rows.Add(MyRow)
        Else
            IsNew = False
            MyRow = MyDataTable.Rows(0)
        End If

        If MyDataTable.Columns.Contains("touched") Then
            ' this.MyRow["touched"] = clsGlobals.touched + "|" + DateTime.Now.ToString();
        End If
    End Sub 'constructor

#End Region


#Region "Updaters"
    Public Sub update()
        MyCon.Open()
        ' this.MyDataAdapter.UpdateCommand = new SqlCommand();
        If MyDataAdapter.SelectCommand.CommandText.Contains("@") Then
            MyDataAdapter.SelectCommand.CommandText = MyDataAdapter.SelectCommand.CommandText.Replace("@", "")
            MyDataAdapter.SelectCommand.Parameters.Clear()
        End If

        MyDataAdapter.Update(MyDataTable)
        MyCon.Close()
    End Sub 'update       

    Public Sub transactionUpdate()
        MyCon.Open()

        If MyDataAdapters.Count > 0 Then
            Dim elements = MyDataAdapters.Count
            Dim siteInsert As List(Of SqlCommand) = New List(Of SqlCommand)()
            Dim siteUpdate As List(Of SqlCommand) = New List(Of SqlCommand)()

            For x = 0 To elements - 1
                ' this.MyDataAdapter.UpdateCommand = new SqlCommand();
                If MyDataAdapters(x).SelectCommand.CommandText.Contains("@") Then
                    MyDataAdapters(x).SelectCommand.CommandText = MyDataAdapter.SelectCommand.CommandText.Replace("@", "")
                    MyDataAdapters(x).SelectCommand.Parameters.Clear()
                End If

                siteInsert.Add(MyCmds(x).GetInsertCommand())
                siteUpdate.Add(MyCmds(x).GetUpdateCommand())
            Next

            Dim tran As SqlTransaction = MyCon.BeginTransaction()

            For x = 0 To elements - 1
                siteInsert(x).Transaction = tran
                siteUpdate(x).Transaction = tran
            Next

            Try

                For x = 0 To elements - 1
                    MyDataAdapters(x).Update(MyDataTables(x))
                Next

                tran.Commit()
            Catch ex As SqlException
                'Roll back the transaction.

                'Additional error handling if needed.
                tran.Rollback()
            Finally
                ' Close the connection.
                MyCon.Close()
            End Try '************************************************************************
        Else
            ' this.MyDataAdapter.UpdateCommand = new SqlCommand();
            If MyDataAdapter.SelectCommand.CommandText.Contains("@") Then
                MyDataAdapter.SelectCommand.CommandText = MyDataAdapter.SelectCommand.CommandText.Replace("@", "")
                MyDataAdapter.SelectCommand.Parameters.Clear()
            End If

            Dim siteInsert As SqlCommand = MyCmd.GetInsertCommand()
            Dim siteUpdate As SqlCommand = MyCmd.GetUpdateCommand()
            Dim tran As SqlTransaction = MyCon.BeginTransaction()
            siteInsert.Transaction = tran
            siteUpdate.Transaction = tran

            Try
                MyDataAdapter.Update(MyDataTable)
                tran.Commit()
            Catch ex As SqlException
                'Roll back the transaction.

                'Additional error handling if needed.
                tran.Rollback()
            Finally
                ' Close the connection.
                MyCon.Close()
            End Try
        End If
    End Sub 'update

    Public Sub insert()
        MyCon.Open()
        Dim counter = 0

        For Each c As DataColumn In MyDataTable.Columns

            If Not (Equals(c.ColumnName.ToString(), "id")) Then
                counter += 1
                ' this.MyDataAdapter.InsertCommand.Parameters.AddWithValue("@p" + counter.ToString(), this.MyRow[c.ColumnName]);
                MyDataAdapter.InsertCommand.Parameters("@p" & counter.ToString()).Value = MyRow(c.ColumnName)
            End If
        Next

        MyDataAdapter.Update(MyDataTable)
        MyCon.Close()
    End Sub 'upd
#End Region

    Public Function Col(ByVal s As String) As Integer
        Return MyDataTable.Columns(s).Ordinal
    End Function 'Col

    Private Sub My_OnRowUpdate(ByVal sender As Object, ByVal e As SqlRowUpdatedEventArgs)
        If e.StatementType = StatementType.Insert Then
            ' reads the identity value from the output parameter @ID
            Dim ai = e.Command.Parameters("@NEWID").Value
            id = CInt(ai)
            ' updates the identity column (autoincrement)                   

            Dim C = MyDataTable.Columns(0)

            If Equals(C.ColumnName, "id") Then
                C.ReadOnly = False
                e.Row(C) = ai
                C.ReadOnly = True
            End If

            e.Row.AcceptChanges()
        End If
    End Sub

    Public Shared Function RunCommand(ByVal query As String, ByVal ConnectionString As String) As Boolean
        Dim transaction As SqlTransaction = Nothing
        Dim conn As SqlConnection = New SqlConnection(ConnectionString)
        Dim cmd As SqlCommand = New SqlCommand()

        Try
            conn.Open()
            cmd.CommandText = query
            cmd.Connection = conn
            cmd.ExecuteNonQuery()
            conn.Close()


            'MySqlCommand psmt = new MySqlCommand();
            Return True
        Catch ex As Exception
            ' Utilities.Logit(ex.Message);
            Return False
        Finally
        End Try
    End Function

    Public Shared Function RunCommand(ByVal query As String, ByVal Params As Hashtable, ByVal ConnectionString As String) As Boolean
        Dim transaction As SqlTransaction = Nothing
        Dim conn As SqlConnection = New SqlConnection(ConnectionString)
        Dim cmd As SqlCommand = New SqlCommand()

        Try
            conn.Open()
            cmd.CommandText = query
            cmd.Connection = conn

            For Each o In Params.Keys
                Dim keyname As String = o.ToString()

                If Not keyname.StartsWith("@") Then
                    keyname = "@" & keyname
                End If

                cmd.Parameters.AddWithValue(keyname, Params(keyname))
                'this.MyDataAdapter.SelectCommand.Parameters.AddWithValue(keyname, Params[keyname]);
                'this.MyDataAdapter.UpdateCommand.Parameters.AddWithValue(keyname, QueryAndParams[keyname]);

            Next

            cmd.ExecuteNonQuery()
            conn.Close()


            'MySqlCommand psmt = new MySqlCommand();
            Return True
        Catch ex As Exception
            ' Utilities.Logit(ex.Message);
            Return False
        Finally
        End Try
    End Function

    Public Overrides Function ToString() As String
        Dim sb As StringBuilder = New StringBuilder()
        Dim counter = 1

        For Each r As DataRow In myDataTableField.Rows
            sb.Append(counter & " -------------------------------------" & Environment.NewLine)
            counter += 1

            For Each c As DataColumn In MyDataTable.Columns
                sb.Append(c.ColumnName & " = " & noNull(r(c.ColumnName)) & "" & Environment.NewLine)
            Next
        Next

        Return sb.ToString()
    End Function

    Public Shared Function GetUniqueID(ByVal ConnectionString As String) As Integer
        Dim gq As clsGenericQuery = New clsGenericQuery("select * from GetUnique where id=0", ConnectionString, True)
        gq.MyRow("Nothing") = "X"
        gq.insert()
        Return gq.id
    End Function

    Private Function TypeDict() As Dictionary(Of Type, DbType)
        Dim typeMap As Dictionary(Of Type, DbType) = New Dictionary(Of Type, DbType)()
        typeMap(GetType(Byte)) = DbType.Byte
        typeMap(GetType(SByte)) = DbType.SByte
        typeMap(GetType(Short)) = DbType.Int16
        typeMap(GetType(UShort)) = DbType.UInt16
        typeMap(GetType(Integer)) = DbType.Int32
        typeMap(GetType(UInteger)) = DbType.UInt32
        typeMap(GetType(Long)) = DbType.Int64
        typeMap(GetType(ULong)) = DbType.UInt64
        typeMap(GetType(Single)) = DbType.Single
        typeMap(GetType(Double)) = DbType.Double
        typeMap(GetType(Decimal)) = DbType.Decimal
        typeMap(GetType(Boolean)) = DbType.Boolean
        typeMap(GetType(String)) = DbType.String
        typeMap(GetType(Char)) = DbType.StringFixedLength
        typeMap(GetType(Guid)) = DbType.Guid
        typeMap(GetType(Date)) = DbType.DateTime
        typeMap(GetType(DateTimeOffset)) = DbType.DateTimeOffset
        typeMap(GetType(Byte())) = DbType.Binary
        typeMap(GetType(Byte?)) = DbType.Byte
        typeMap(GetType(SByte?)) = DbType.SByte
        typeMap(GetType(Short?)) = DbType.Int16
        typeMap(GetType(UShort?)) = DbType.UInt16
        typeMap(GetType(Integer?)) = DbType.Int32
        typeMap(GetType(UInteger?)) = DbType.UInt32
        typeMap(GetType(Long?)) = DbType.Int64
        typeMap(GetType(ULong?)) = DbType.UInt64
        typeMap(GetType(Single?)) = DbType.Single
        typeMap(GetType(Double?)) = DbType.Double
        typeMap(GetType(Decimal?)) = DbType.Decimal
        typeMap(GetType(Boolean?)) = DbType.Boolean
        typeMap(GetType(Char?)) = DbType.StringFixedLength
        typeMap(GetType(Guid?)) = DbType.Guid
        typeMap(GetType(Date?)) = DbType.DateTime
        typeMap(GetType(DateTimeOffset?)) = DbType.DateTimeOffset
        'typeMap[typeof(Binary)] = DbType.Binary;
        Return typeMap
    End Function

    Private Function noNull(ByVal o As Object) As String
        If TypeOf o Is DBNull Then
            Return "NULL"
        End If

        Return o.ToString()
    End Function
End Class 'class

