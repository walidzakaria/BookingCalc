Imports System.Data.SqlClient

Public Class Utils
    Public Shared Property ConnectionString As String = ""

    Public Shared Function GetData(ByVal query As String) As DataTable
        Dim myConn = New SqlConnection(ConnectionString)
        Dim adt As New SqlDataAdapter
        Dim ds As New DataSet()
        Using cmd = New SqlCommand(query, myConn)
            If myConn.State = ConnectionState.Closed Then
                myConn.Open()
            End If

            adt.SelectCommand = cmd
            adt.Fill(ds)
            adt.Dispose()

            myConn.Close()
        End Using

        Return ds.Tables(0)
    End Function
End Class

