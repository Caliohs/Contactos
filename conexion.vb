
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.ComponentModel
Imports System.Text
Public Class conexion

    Private cn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data source=contactos.accdb")
    Public Function consulta(ByVal sql, ByVal tabla)
        Try
            Dim da As New OleDb.OleDbDataAdapter(sql, cn)
            Dim ds As New DataSet
            da.Fill(ds, tabla)
            Return ds
        Catch ex As Exception
            MsgBox("error" & ex.ToString)
            Dim ds As New DataSet
            Return ds
        End Try
    End Function
    Public Function insertar(ByVal sql)
        Try
            Dim d As New OleDbCommand(sql, cn)
            cn.Open()
            d.ExecuteNonQuery()
            cn.Close()
            Return True
        Catch ex As Exception
            MsgBox(ex.ToString)
            Return False
        End Try
    End Function

End Class
