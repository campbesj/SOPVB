Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Data
Imports System.Reflection


Public Class SOP
        Public Property itemList As List(Of SOP_Items)

        Public Sub New()
        End Sub

    Public Sub loadItems(ByVal Institution As String, ByVal Optional loadBranches As Boolean = False)
        Me.itemList = New List(Of SOP_Items)
        Dim gq As clsGenericQuery = New clsGenericQuery("select * from SOP_Items where Institution=" & Institution, UtilitySOP.connstr())
        Dim x As String

        If Not gq.IsNew Then
            For Each r As DataRow In gq.MyDataTable.Rows
                Dim i As SOP_Items = New SOP_Items(r)

                If loadBranches Then
                    i.loadBranches()
                End If

                itemList.Add(i)
            Next
        End If
    End Sub

    Public Class SOP_Items
        Public Property branches As List(Of SOP_Branches)
        Public Property ID As Integer
        Public Property Institution As String
        Public Property Limited_By_Branch As Boolean
        Public Property Text As String
        Public Property Resp_Type As String
        Public Property Sort_No As Byte
        Public Property [Optional] As Boolean
        Public Property Completion_Period As String

        Public Sub New()
            branches = New List(Of SOP_Branches)()
        End Sub

        Public Sub New(ByVal ID As Integer)
            branches = New List(Of SOP_Branches)()
            Dim gq As clsGenericQuery = New clsGenericQuery("select * from SOP_Items where ID=" & ID, UtilitySOP.connstr())
            Call UtilitySOP.loadObject(gq, Me, [GetType]())
        End Sub

        Public Sub New(ByVal r As DataRow)
            branches = New List(Of SOP_Branches)()
            Call UtilitySOP.loadObjectFromRow(r, Me, [GetType]())
        End Sub

        Public Function store() As Boolean
            Try
                Dim gq As clsGenericQuery = New clsGenericQuery("select * from SOP_Items where ID='" & ID & "'", UtilitySOP.connstr())
                Call UtilitySOP.storeObject(gq, Me, [GetType]())
                gq.update()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

        Public Sub loadBranches(ByVal Optional loadResp As Boolean = False)
            Dim gq As clsGenericQuery = New clsGenericQuery("select * from SOP_Branches where SOP_Item=" & ID, UtilitySOP.connstr())

            If Not gq.IsNew Then
                For Each r As DataRow In gq.MyDataTable.Rows
                    branches.Add(New SOP_Branches(r, loadResp))
                Next
            End If
        End Sub

        Public Sub generateCompletionTracking()
            Dim dowDict As Dictionary(Of DayOfWeek, String) = New Dictionary(Of DayOfWeek, String)()
            dowDict.Add(DayOfWeek.Sunday, "EVSUN")
            dowDict.Add(DayOfWeek.Monday, "EVMON")
            dowDict.Add(DayOfWeek.Tuesday, "EVTUE")
            dowDict.Add(DayOfWeek.Wednesday, "EVWED")
            dowDict.Add(DayOfWeek.Thursday, "EVTHU")
            dowDict.Add(DayOfWeek.Friday, "EVFRI")
            dowDict.Add(DayOfWeek.Saturday, "EVSAT")
            Dim dow = Date.Now.DayOfWeek
            Dim dowString = dowDict(dow)
            Dim FMONTH = If(Date.Now.Day = 1, True, False)
            Dim endOfMonth As Date = New DateTime(Date.Now.Year, Date.Now.Month, 1).AddMonths(1).AddDays(-1)
            Dim LMONTH = If(Date.Now.Day = endOfMonth.Day, True, False)
            'DAILY, EVMON, EVTUE, EVWED, EVTHU, EVFRI, EVSAT, EVSUN, FMONTH, LMONTH


            For Each br In branches

                If br.Start_Date <= Date.Now AndAlso br.End_Date >= Date.Now Then
                    Dim cc As CalendarCalc = New CalendarCalc(br.Branch, Institution)

                    If Equals(Completion_Period, "DAILY") AndAlso cc.isopen OrElse Equals(Completion_Period, "FMONTH") AndAlso Date.Now.Day = cc.first OrElse Equals(Completion_Period, "LMONTH") AndAlso Date.Now.Day = cc.last OrElse Equals(Completion_Period, dowString) AndAlso cc.isopen Then
                        If Equals(br.Complete_Type, "ANYONECAN") Then
                            Dim ct As SOP_Completion_Tracking = New SOP_Completion_Tracking()
                            ct.Date = Date.Now
                            ct.Employee_ID = "ANYONECAN"
                            ct.SOP_Branch = br.ID
                            ct.SOP_Item = ID
                            ct.store()
                        Else

                            For Each resp In br.responsibilities
                                Dim ct As SOP_Completion_Tracking = New SOP_Completion_Tracking()
                                ct.Date = Date.Now
                                ct.Employee_ID = resp.Employee_ID
                                ct.SOP_Branch = br.ID
                                ct.SOP_Item = ID
                                ct.store()
                            Next
                        End If
                    End If
                End If
            Next
        End Sub
    End Class

    Public Class SOP_Branches
        Public Property responsibilities As List(Of SOP_Responsibility)
        Public Property ID As Integer
        Public Property SOP_Item As Integer
        Public Property Branch As String
        Public Property Start_Date As Date
        Public Property End_Date As Date
        Public Property Daily_Completed_By_Time As TimeSpan
        Public Property Complete_Type As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal ID As Integer)
            Dim gq As clsGenericQuery = New clsGenericQuery("select * from SOP_Branches where ID=" & ID, UtilitySOP.connstr())
            Call UtilitySOP.loadObject(gq, Me, [GetType]())
        End Sub

        Public Sub New(ByVal r As DataRow, ByVal Optional loadResp As Boolean = False)
            Call UtilitySOP.loadObjectFromRow(r, Me, [GetType]())

            If loadResp Then
                loadResponsibilities()
            End If
        End Sub

        Public Function store() As Boolean
            Try
                Dim gq As clsGenericQuery = New clsGenericQuery("select * from SOP_Branches where ID='" & ID & "'", UtilitySOP.connstr())
                Call UtilitySOP.storeObject(gq, Me, [GetType]())
                gq.update()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

        Public Sub loadResponsibilities()
            responsibilities = New List(Of SOP_Responsibility)()
            Dim gq As clsGenericQuery = New clsGenericQuery("select * from SOP_Responsibility where SOP_Item='" & SOP_Item & "' and SOP_Branch = '" & ID & "'", UtilitySOP.connstr())

            If Not gq.IsNew Then
                For Each r As DataRow In gq.MyDataTable.Rows
                    responsibilities.Add(New SOP_Responsibility(r))
                Next
            End If
        End Sub
    End Class

    Public Class SOP_Responsibility
        Public Property ID As Integer
        Public Property SOP_Item As Integer
        Public Property SOP_Branch As Integer
        Public Property Employee_ID As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal ID As Integer)
            Dim gq As clsGenericQuery = New clsGenericQuery("select * from SOP_Responsibility where ID=" & ID, UtilitySOP.connstr())
            Call UtilitySOP.loadObject(gq, Me, [GetType]())
        End Sub

        Public Sub New(ByVal r As DataRow)
            Call UtilitySOP.loadObjectFromRow(r, Me, [GetType]())
        End Sub

        Public Function store() As Boolean
            Try
                Dim gq As clsGenericQuery = New clsGenericQuery("select * from SOP_Responsibility where ID='" & ID & "'", UtilitySOP.connstr())
                Call UtilitySOP.storeObject(gq, Me, [GetType]())
                gq.update()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
    End Class

    Public Class SOP_Completion_Tracking
        Public Property ID As Integer
        Public Property SOP_Item As Integer
        Public Property SOP_Branch As Integer
        Public Property Employee_ID As String
        Public Property [Date] As Date
        Public Property Completion_Time As TimeSpan
        Public Property Response As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal ID As Integer)
            Dim gq As clsGenericQuery = New clsGenericQuery("select * from SOP_Completion_Tracking where ID=" & ID, UtilitySOP.connstr())
            Call UtilitySOP.loadObject(gq, Me, [GetType]())
        End Sub

        Public Sub New(ByVal r As DataRow)
            Call UtilitySOP.loadObjectFromRow(r, Me, [GetType]())
        End Sub

        Public Function store() As Boolean
            Try
                Dim gq As clsGenericQuery = New clsGenericQuery("select * from SOP_Completion_Tracking where ID='" & ID & "'", UtilitySOP.connstr())
                Call UtilitySOP.storeObject(gq, Me, [GetType]())
                gq.update()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
    End Class

    Public Class SOP_Notifications
        Public Property ID As Integer
        Public Property SOP_Item As Integer
        Public Property SOP_Branch As Integer
        Public Property Type As String
        Public Property Text As String
        Public Property Report As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal ID As Integer)
            Dim gq As clsGenericQuery = New clsGenericQuery("select * from SOP_Notificationswhere ID=" & ID, UtilitySOP.connstr())
            Call UtilitySOP.loadObject(gq, Me, [GetType]())
        End Sub

        Public Sub New(ByVal r As DataRow)
            Call UtilitySOP.loadObjectFromRow(r, Me, [GetType]())
        End Sub

        Public Function store() As Boolean
            Try
                Dim gq As clsGenericQuery = New clsGenericQuery("select * from SOP_Notifications where ID='" & ID & "'", UtilitySOP.connstr())
                Call UtilitySOP.storeObject(gq, Me, [GetType]())
                gq.update()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
    End Class

    Public Class SOP_Notify
        Public Property ID As Integer
        Public Property SOP_Item As Integer
        Public Property SOP_Branch As Integer
        Public Property Employee_ID As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal ID As Integer)
            Dim gq As clsGenericQuery = New clsGenericQuery("select * from SOP_Notify where ID=" & ID, UtilitySOP.connstr())
            Call UtilitySOP.loadObject(gq, Me, [GetType]())
        End Sub

        Public Sub New(ByVal r As DataRow)
            Call UtilitySOP.loadObjectFromRow(r, Me, [GetType]())
        End Sub

        Public Function store() As Boolean
            Try
                Dim gq As clsGenericQuery = New clsGenericQuery("select * from SOP_Notify where ID='" & ID & "'", UtilitySOP.connstr())
                Call UtilitySOP.storeObject(gq, Me, [GetType]())
                gq.update()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
    End Class

    Public Class CalendarCalc
        Public Property first As Integer
        Public Property last As Integer
        Public Property isopen As Boolean

        Public Sub New()
        End Sub

        Public Sub New(ByVal Branch As String, ByVal Institution As String)
            Dim query = "select * from Calendar where Branch='" & Branch & "' and Institution = '" & Institution
            query += "' and DATEPART(yyyy,[Date]) = '" & Date.Now.Year & "' and DATEPART(m,[Date]) = '" & Date.Now.Month & "' and Date_Open = '1'"
            Dim gq As clsGenericQuery = New clsGenericQuery(query, UtilitySOP.connstr())

            If Not gq.IsNew Then
                Dim foundRows As DataRow() = gq.MyDataTable.Select(" 1=1", "[Date] desc")
                last = Date.Parse(foundRows(0).Item("Date").ToString()).Day
                foundRows = gq.MyDataTable.Select(" 1=1", "[Date] ")
                first = Date.Parse(foundRows(0).Item("Date").ToString()).Day
                foundRows = gq.MyDataTable.Select("[Date] = '" & Date.Now.ToShortDateString() & "'")
                isopen = foundRows.Count() > 0
            End If
        End Sub
    End Class



    '-----------------------------------------------------------------------------------//

    Public Class UtilitySOP
        Public Shared Function connstr() As String
            Return "Data Source = 13.68.178.65,1443\ACPlus1; Initial Catalog = InstitutionSet1; Integrated Security = False; User ID = phpuser; Password = dodntint]"
        End Function

        Public Shared Sub storeObject(ByRef gq As clsGenericQuery, ByVal o As Object, ByVal t As Type)
            For Each c As DataColumn In gq.MyDataTable.Columns
                Dim propertyInfo = t.GetProperty(c.ColumnName)
                ' make sure object has the property we are after
                If propertyInfo IsNot Nothing Then
                    If Not c.AutoIncrement Then
                        Dim typeName = propertyInfo.PropertyType.Name

                        Select Case typeName
                            Case "String"
                                gq.MyRow.Item(c.ColumnName) = noNull(propertyInfo.GetValue(o))
                            Case "Decimal"
                                gq.MyRow.Item(c.ColumnName) = noNullDec(propertyInfo.GetValue(o))
                            Case "Double"
                                gq.MyRow.Item(c.ColumnName) = noNullDouble(propertyInfo.GetValue(o))
                            Case "DateTime"
                                gq.MyRow.Item(c.ColumnName) = noNullDate(propertyInfo.GetValue(o))
                            Case "Int32"
                                gq.MyRow.Item(c.ColumnName) = noNullInt(propertyInfo.GetValue(o))
                            Case "Int64"
                                gq.MyRow.Item(c.ColumnName) = noNulllong(propertyInfo.GetValue(o))
                            Case "Boolean"
                                gq.MyRow.Item(c.ColumnName) = propertyInfo.GetValue(o)
                            Case "Byte"
                                gq.MyRow.Item(c.ColumnName) = noNullByte(propertyInfo.GetValue(o))
                            Case "TimeSpan"
                                gq.MyRow.Item(c.ColumnName) = noNullTimespan(propertyInfo.GetValue(o))
                        End Select
                    End If
                End If
            Next
        End Sub

        Public Shared Sub loadObject(ByVal gq As clsGenericQuery, ByVal o As Object, ByVal t As Type)
            For Each c As DataColumn In gq.MyDataTable.Columns
                Dim propertyInfo = t.GetProperty(c.ColumnName)
                ' make sure object has the property we are after
                If propertyInfo IsNot Nothing Then
                    If Not gq.MyRow.Item(c.ColumnName) Is DBNull.Value Then
                        propertyInfo.SetValue(o, gq.MyRow.Item(c.ColumnName), Nothing)
                    End If
                End If
            Next
        End Sub

        Public Shared Sub loadObjectFromRow(ByVal r As DataRow, ByVal o As Object, ByVal t As Type)
                For Each c As DataColumn In r.Table.Columns
                    Dim propertyInfo = t.GetProperty(c.ColumnName)
                    ' make sure object has the property we are after
                    If propertyInfo IsNot Nothing Then
                    If Not r.Item(c.ColumnName) Is DBNull.Value Then
                        propertyInfo.SetValue(o, r.Item(c.ColumnName), Nothing)
                    End If
                End If
                Next
            End Sub

            Public Shared Function noNull(ByVal o As Object) As String
                If o Is Nothing Then
                    Return ""
                End If

                Return o.ToString()
            End Function

            Public Shared Function noNullBool(ByVal o As Object) As Boolean
                If o Is Nothing Then
                    Return False
                End If

                Return o
            End Function

            Public Shared Function noNullDec(ByVal o As Object) As Decimal
                If o Is Nothing Then
                    Return 0
                End If

                Return o
            End Function

            Public Shared Function noNullDouble(ByVal o As Object) As Double
                If o Is Nothing Then
                    Return 0
                End If

                Return o
            End Function

            Public Shared Function noNullInt(ByVal o As Object) As Integer
                If o Is Nothing Then
                    Return 0
                End If

                Return o
            End Function

            Public Shared Function noNullByte(ByVal o As Object) As Byte
                If o Is Nothing Then
                    Return 0
                End If

                Return o
            End Function

            Public Shared Function noNulllong(ByVal o As Object) As Long
                If o Is Nothing Then
                    Return 0
                End If

                Return o
            End Function

            Public Shared Function noNullTimespan(ByVal o As Object) As TimeSpan
                Dim dt As TimeSpan = o

                If o Is Nothing Then
                    Return New TimeSpan(0, 0, 0)
                End If

                Return o
            End Function

            Public Shared Function noNullDate(ByVal o As Object) As Date
                Dim dt As Date = o

                If dt.Year < 2000 Then
                    Return New DateTime(2001, 1, 1)
                End If

                Return o
            End Function
        End Class
    End Class

