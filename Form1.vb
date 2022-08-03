Imports System.Data
Imports System.Data.OleDb



Public Class frm_contactos

    Private con As New OleDb.OleDbConnection
    Private command As OleDb.OleDbCommand
    Private command1 As OleDb.OleDbCommand
    Private adapter As OleDbDataAdapter
    Private reader As OleDbDataReader
    Private reader1 As OleDbDataReader
    Private dataSt As DataSet
    Private sql As String
    Private sql1 As String
    Dim editar As Boolean = False
    Dim IDE As Integer
    Dim rowsEditar As Integer
    Dim cn As New conexion
    Dim idEditar As Integer
    Dim buscar As String = "Buscar... 🔎"
    Dim sp As String



    Private Sub frm_contactos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        consulta()
        Dim strCon As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath + "\contactos.accdb"
        ' ESTABLECE LA CONEXION

        With con
            If .State = ConnectionState.Closed Then
                .ConnectionString = strCon
                .Open()
            End If
        End With

        autocomplete() ' AUTOCOMPLETA PALABRAS DE BUSQUEDA
        cargarCombo() ' CARGA LOS COMBOX CON LOS DATOS DE ACCESS
        Call consulta() ' CARGA EL DATAGRID


        TextBox10.Text = buscar
        TextBox10.ForeColor = Color.Red
        TextBox17.Text = buscar
        TextBox17.ForeColor = Color.Red
        TextBox20.Text = buscar
        TextBox20.ForeColor = Color.Red

    End Sub
    Sub consulta()

        dataSt = cn.consulta("select * from contactos", "contactos")
        DataGridView1.DataSource = dataSt.Tables(0)

        dataSt = cn.consulta("select * from evertec", "contactos")
        DataGridView3.DataSource = dataSt.Tables(0)

        dataSt = cn.consulta("select * from puertos", "contactos")
        DataGridView2.DataSource = dataSt.Tables(0)

    End Sub
    Private Sub autocomplete()  'sugiere mediante la busqueda

        'AUTOCOMPLETA NOMBRE DE CLIENTE/pais para busqueda
        sql = "SELECT * FROM contactos"
        command = New OleDbCommand(sql, con)
        reader = command.ExecuteReader()
        Dim autoComp As New AutoCompleteStringCollection()

        While reader.Read()
            autoComp.Add(reader("cliente"))
            autoComp.Add(reader("pais"))
        End While
        reader.Close()
        TextBox10.AutoCompleteMode = AutoCompleteMode.Suggest
        TextBox10.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox10.AutoCompleteCustomSource = autoComp
        'termina con cliente/pais
        '***********************************************************************

        'AUTOCOMPLETA NOMBRE DE CLIENTE/pais para busqueda
        sql = "SELECT * FROM contactos"
        command = New OleDbCommand(sql, con)
        reader = command.ExecuteReader()
        Dim autoComp6 As New AutoCompleteStringCollection()

        While reader.Read()
            autoComp6.Add(reader("cliente"))
        End While
        reader.Close()
        TextBox14.AutoCompleteMode = AutoCompleteMode.Suggest
        TextBox14.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox14.AutoCompleteCustomSource = autoComp6
        'termina con cliente/pais
        '************************************************************************
        'AUTOCOMPLETA pais 
        sql = "SELECT pais FROM contactos"
        command = New OleDbCommand(sql, con)
        reader = command.ExecuteReader()
        Dim autoComp3 As New AutoCompleteStringCollection()

        While reader.Read()
            autoComp3.Add(reader("pais"))
        End While
        reader.Close()
        TextBox15.AutoCompleteMode = AutoCompleteMode.Suggest
        TextBox15.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox15.AutoCompleteCustomSource = autoComp3
        'termina con pais
        '***********************************************************************

        'autocompleta un nombre EN CONTACTOS EVERTEC
        sql = "SELECT Nombre FROM evertec"
        command = New OleDbCommand(sql, con)
        reader = command.ExecuteReader()
        Dim autoComp1 As New AutoCompleteStringCollection()

        While reader.Read()
            autoComp1.Add(reader("Nombre"))
        End While
        reader.Close()
        TextBox9.AutoCompleteMode = AutoCompleteMode.Suggest
        TextBox9.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox9.AutoCompleteCustomSource = autoComp1
        'termina con nombre

        '************************************************************************
        'autocompleta un nombre columna AREA
        sql = "SELECT Area FROM evertec"
        command = New OleDbCommand(sql, con)
        reader = command.ExecuteReader()
        Dim autoComp2 As New AutoCompleteStringCollection()

        While reader.Read()
            autoComp2.Add(reader("Area"))
        End While
        reader.Close()
        TextBox16.AutoCompleteMode = AutoCompleteMode.Suggest
        TextBox16.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox16.AutoCompleteCustomSource = autoComp2
        'termina con Area
        '************************************************************************
        'autocompleta busqueda contatos evertec
        sql = "SELECT * FROM evertec"
        command = New OleDbCommand(sql, con)
        reader = command.ExecuteReader()
        Dim autoComp8 As New AutoCompleteStringCollection()

        While reader.Read()
            autoComp8.Add(reader("Nombre"))
            autoComp8.Add(reader("Area"))
        End While
        reader.Close()
        TextBox17.AutoCompleteMode = AutoCompleteMode.Suggest
        TextBox17.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox17.AutoCompleteCustomSource = autoComp8
        'termina con busqueda
        '************************************************************************

        'autocompleta un puerto de transerver
        sql = "SELECT Puerto FROM puertos"
        command = New OleDbCommand(sql, con)
        reader = command.ExecuteReader()
        Dim autoComp4 As New AutoCompleteStringCollection()

        While reader.Read()
            autoComp4.Add(reader("Puerto"))
        End While
        reader.Close()
        TextBox6.AutoCompleteMode = AutoCompleteMode.Suggest
        TextBox6.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox6.AutoCompleteCustomSource = autoComp4
        'termina con puerto

        'AUTOCOMPLETA NOMBRE DE instnacia
        sql = "SELECT Instancia FROM puertos"
        command = New OleDbCommand(sql, con)
        reader = command.ExecuteReader()
        Dim autoComp7 As New AutoCompleteStringCollection()

        While reader.Read()
            autoComp7.Add(reader("Instancia"))
        End While
        reader.Close()
        TextBox18.AutoCompleteMode = AutoCompleteMode.Suggest
        TextBox18.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox18.AutoCompleteCustomSource = autoComp7
        'termina con instancia


        '************************************************************************
        'autocompleta un socket de transerver
        sql = "SELECT Socket FROM puertos"
        command = New OleDbCommand(sql, con)
        reader = command.ExecuteReader()
        Dim autoComp5 As New AutoCompleteStringCollection()

        While reader.Read()
            autoComp5.Add(reader("Socket"))
        End While
        reader.Close()
        TextBox7.AutoCompleteMode = AutoCompleteMode.Suggest
        TextBox7.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox7.AutoCompleteCustomSource = autoComp5
        'termina con socket
        '********************************************************************************
        'autocompleta busqueda de PUERTOS transerver
        sql = "SELECT * FROM puertos"
        command = New OleDbCommand(sql, con)
        reader = command.ExecuteReader()
        Dim autoComp9 As New AutoCompleteStringCollection()

        While reader.Read()
            autoComp9.Add(reader("Puerto"))
            autoComp9.Add(reader("Instancia"))
            autoComp9.Add(reader("Descripcion"))
            autoComp9.Add(reader("Socket"))
        End While
        reader.Close()
        TextBox20.AutoCompleteMode = AutoCompleteMode.Suggest
        TextBox20.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox20.AutoCompleteCustomSource = autoComp9
        'termina con busqueda

    End Sub



    Public Sub cargarCombo() ' carga los combos desde las tablas***********************

        'Dim tabla1 As New DataTable
        'Dim sql1 As String = "SELECT DISTINCT cliente FROM contactos"
        'Dim adapter1 As New OleDbDataAdapter(sql1, con)
        'adapter1.Fill(tabla1)
        'ComboBox1.DataSource = tabla1 'combo para clientes de contactos
        'ComboBox1.DisplayMember = "cliente"



    End Sub
    '***************************************************************************************

    '*********GUARDA CONTACTOS CLIENTES
    Private Sub btn_guardar_Click_1(sender As Object, e As EventArgs) Handles btn_guardar.Click
        If editar = False Then

            If TextBox14.Text = "" Or TextBox15.Text = "" Or TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
                MsgBox("No se permiten campos vacios")
            Else

                sql = "INSERT into Contactos (cliente,nombre,telefono,correo,descripcion, pais) values("
                sql = sql & "'" & TextBox14.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox15.Text & "')"
                cn.insertar(sql)
                TextBox14.Text = ""
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox4.Text = ""
                TextBox15.Text = ""
            End If

        Else
            If TextBox14.Text = "" Or TextBox15.Text = "" Or TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
                MsgBox("No se permiten campos vacios")
            Else
                cn.consulta("Update contactos set cliente='" & TextBox14.Text & "' where Id=" & idEditar, "contactos")
                cn.consulta("Update contactos set nombre='" & TextBox1.Text & "' where Id=" & idEditar, "contactos")
                cn.consulta("Update contactos set telefono='" & TextBox2.Text & "' where Id=" & idEditar, "contactos")
                cn.consulta("Update contactos set correo='" & TextBox3.Text & "' where Id=" & idEditar, "contactos")
                cn.consulta("Update contactos set descripcion='" & TextBox4.Text & "' where Id=" & idEditar, "contactos")
                cn.consulta("Update contactos set pais='" & TextBox15.Text & "' where Id=" & idEditar, "contactos")

                TextBox14.Text = ""
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox4.Text = ""
                TextBox15.Text = ""

                DataGridView1.ClearSelection()
                MsgBox("Se han editado los datos")
                editar = False
                idEditar = 0
            End If
        End If
    End Sub

    'EDITA CONTACTOS CLIENTES***********
    Private Sub btn_editar_Click_1(sender As Object, e As EventArgs) Handles btn_editar.Click


        If DataGridView1.SelectedRows.Count > 0 Then


            TextBox14.Text = DataGridView1.Rows(DataGridView1.SelectedRows(0).Index).Cells(1).Value.ToString
            TextBox1.Text = DataGridView1.Rows(DataGridView1.SelectedRows(0).Index).Cells(2).Value.ToString
            TextBox2.Text = DataGridView1.Rows(DataGridView1.SelectedRows(0).Index).Cells(3).Value.ToString
            TextBox3.Text = DataGridView1.Rows(DataGridView1.SelectedRows(0).Index).Cells(4).Value.ToString
            TextBox4.Text = DataGridView1.Rows(DataGridView1.SelectedRows(0).Index).Cells(5).Value.ToString
            TextBox15.Text = DataGridView1.Rows(DataGridView1.SelectedRows(0).Index).Cells(6).Value.ToString


            editar = True
            rowsEditar = DataGridView1.SelectedRows(0).Index
            idEditar = CInt(DataGridView1.Rows(DataGridView1.SelectedRows(0).Index).Cells(0).Value)

        Else
            MsgBox("no se han selecionado contactos")
        End If
    End Sub

    'ELIMINA CONTACTOS CLIENTES
    Private Sub btn_eliminar_Click_1(sender As Object, e As EventArgs) Handles btn_eliminar.Click
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim result = MessageBox.Show("¿Desea eliminar este registro?", "Alerta", MessageBoxButtons.YesNo)
            If result = Windows.Forms.DialogResult.Yes Then

                Dim id As Integer = 0
                id = CInt(DataGridView1.Rows(DataGridView1.SelectedRows(0).Index).Cells(0).Value)
                cn.consulta("delete from contactos where id=" & id, "contactos")
                Call consulta()


            End If
        Else
            MsgBox("No se ha seleccionado ninguna fila")
        End If
    End Sub

    'boton de editar contactos evtc**********************
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        If DataGridView3.SelectedRows.Count > 0 Then
            TextBox9.Text = DataGridView3.Rows(DataGridView3.SelectedRows(0).Index).Cells(1).Value.ToString
            TextBox16.Text = DataGridView3.Rows(DataGridView3.SelectedRows(0).Index).Cells(2).Value.ToString
            TextBox11.Text = DataGridView3.Rows(DataGridView3.SelectedRows(0).Index).Cells(3).Value.ToString
            TextBox12.Text = DataGridView3.Rows(DataGridView3.SelectedRows(0).Index).Cells(4).Value.ToString
            TextBox13.Text = DataGridView3.Rows(DataGridView3.SelectedRows(0).Index).Cells(5).Value.ToString
            editar = True
            rowsEditar = DataGridView3.SelectedRows(0).Index
            idEditar = CInt(DataGridView3.Rows(DataGridView3.SelectedRows(0).Index).Cells(0).Value)

        Else
            MsgBox("no se han selecionado contactos")
        End If
    End Sub

    'BOTON ELIMINAR CONTACTOS EVTC************************
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        If DataGridView3.SelectedRows.Count > 0 Then
            Dim result = MessageBox.Show("¿Desea eliminar este registro?", "Alerta", MessageBoxButtons.YesNo)
            If result = Windows.Forms.DialogResult.Yes Then

                Dim id As Integer = 0
                id = CInt(DataGridView3.Rows(DataGridView3.SelectedRows(0).Index).Cells(0).Value)
                cn.consulta("delete from evertec where id=" & id, "contactos")
                Call consulta()


            End If
        Else
            MsgBox("No se ha seleccionado ninguna fila")
        End If
    End Sub

    ' guarda valores de contactos evt******************
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        If editar = False Then

            If TextBox9.Text = "" Or TextBox16.Text = "" Or TextBox11.Text = "" Or TextBox12.Text = "" Or TextBox13.Text = "" Then
                MsgBox("No se permiten campos vacios")
            Else
                sql = "INSERT into evertec (Nombre,Area,Extension,Telefono,Correo) values("
                sql = sql & "'" & TextBox9.Text & "','" & TextBox16.Text & "','" & TextBox11.Text & "','" & TextBox12.Text & "','" & TextBox13.Text & "')"
                cn.insertar(sql)
                TextBox9.Text = ""
                TextBox16.Text = ""
                TextBox11.Text = ""
                TextBox12.Text = ""
                TextBox13.Text = ""
            End If

        Else

            If TextBox9.Text = "" Or TextBox16.Text = "" Or TextBox11.Text = "" Or TextBox12.Text = "" Or TextBox13.Text = "" Then
                MsgBox("No se permiten campos vacios")
            Else
                cn.consulta("Update evertec set Nombre='" & TextBox9.Text & "' where Id=" & idEditar, "contactos")
                cn.consulta("Update evertec set Area='" & TextBox16.Text & "' where Id=" & idEditar, "contactos")
                cn.consulta("Update evertec set Extension='" & TextBox11.Text & "' where Id=" & idEditar, "contactos")
                cn.consulta("Update evertec set Telefono='" & TextBox12.Text & "' where Id=" & idEditar, "contactos")
                cn.consulta("Update evertec set Correo='" & TextBox13.Text & "' where Id=" & idEditar, "contactos")

                TextBox9.Text = ""
                TextBox16.Text = ""
                TextBox11.Text = ""
                TextBox12.Text = ""
                TextBox13.Text = ""

                DataGridView3.ClearSelection()
                MsgBox("Se han editado los datos")
                editar = False
                idEditar = 0
            End If

        End If
    End Sub


    'editar para  puertos transerver************************************************
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If DataGridView2.SelectedRows.Count > 0 Then

            TextBox6.Text = DataGridView2.Rows(DataGridView2.SelectedRows(0).Index).Cells(1).Value.ToString
            TextBox18.Text = DataGridView2.Rows(DataGridView2.SelectedRows(0).Index).Cells(2).Value.ToString
            TextBox5.Text = DataGridView2.Rows(DataGridView2.SelectedRows(0).Index).Cells(3).Value.ToString
            TextBox7.Text = DataGridView2.Rows(DataGridView2.SelectedRows(0).Index).Cells(4).Value.ToString
            TextBox8.Text = DataGridView2.Rows(DataGridView2.SelectedRows(0).Index).Cells(5).Value.ToString
            TextBox19.Text = DataGridView2.Rows(DataGridView2.SelectedRows(0).Index).Cells(6).Value.ToString

            editar = True
            rowsEditar = DataGridView2.SelectedRows(0).Index
            idEditar = CInt(DataGridView2.Rows(DataGridView2.SelectedRows(0).Index).Cells(0).Value)

        Else
            MsgBox("No se ha selecionado ninguna fila")
        End If
    End Sub


    ' guarda valores de puertos ts**************************************************
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        If editar = False Then

            If TextBox6.Text = "" Or TextBox18.Text = "" Or TextBox5.Text = "" Or TextBox7.Text = "" Or TextBox8.Text = "" Or TextBox19.Text = "" Then
                MsgBox("No se permiten campos vacios")
            Else
                sql = "INSERT into puertos (Puerto,Instancia,Descripcion,Socket,IP,Estado) values("
                sql = sql & "'" & TextBox6.Text & "','" & TextBox18.Text & "','" & TextBox5.Text & "','" & TextBox7.Text & "','" & TextBox8.Text & "','" & TextBox19.Text & "')"
                cn.insertar(sql)
                TextBox6.Text = ""
                TextBox18.Text = ""
                TextBox5.Text = ""
                TextBox7.Text = ""
                TextBox8.Text = ""
                TextBox19.Text = ""
            End If
        Else
            If TextBox6.Text = "" Or TextBox18.Text = "" Or TextBox5.Text = "" Or TextBox7.Text = "" Or TextBox8.Text = "" Or TextBox19.Text = "" Then
                MsgBox("No se permiten campos vacios")
            Else

                cn.consulta("Update puertos set Puerto='" & TextBox6.Text & "' where Id=" & idEditar, "contactos")
                cn.consulta("Update puertos set Instancia='" & TextBox18.Text & "' where Id=" & idEditar, "contactos")
                cn.consulta("Update puertos set Descripcion='" & TextBox5.Text & "' where Id=" & idEditar, "contactos")
                cn.consulta("Update puertos set Socket='" & TextBox7.Text & "' where Id=" & idEditar, "contactos")
                cn.consulta("Update puertos set IP='" & TextBox8.Text & "' where Id=" & idEditar, "contactos")
                cn.consulta("Update puertos set Estado='" & TextBox19.Text & "' where Id=" & idEditar, "contactos")

                TextBox6.Text = ""
                TextBox18.Text = ""
                TextBox5.Text = ""
                TextBox7.Text = ""
                TextBox8.Text = ""
                TextBox19.Text = ""

                DataGridView2.ClearSelection()
                MsgBox("Se han editado los datos")
                editar = False
                idEditar = 0
            End If

        End If
    End Sub

    'elimina fila de puertos ts****************************************************
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        If DataGridView2.SelectedRows.Count > 0 Then
            Dim result = MessageBox.Show("¿Desea eliminar este registro?", "Alerta", MessageBoxButtons.YesNo)
            If result = Windows.Forms.DialogResult.Yes Then

                Dim id As Integer = 0
                id = CInt(DataGridView2.Rows(DataGridView2.SelectedRows(0).Index).Cells(0).Value)
                cn.consulta("delete from puertos where id=" & id, "contactos")
                Call consulta()


            End If
        Else
            MsgBox("No se ha seleccionado ninguna fila")
        End If
    End Sub

    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++indica que la caja es un buscador+++++++++++++


    Private Sub TextBox10_Enter(sender As Object, e As EventArgs) Handles TextBox10.Enter
        Dim txt As TextBox = TryCast(sender, TextBox)
        If txt.Text.Equals(buscar) Then
            txt.Text = String.Empty
            txt.ForeColor = SystemColors.WindowText
        End If

    End Sub

    Private Sub TextBox10_Leave(sender As Object, e As EventArgs) Handles TextBox10.Leave
        Dim txt As textBox = TryCast(sender, TextBox)
        If txt.Text.Trim() = String.Empty Then
            txt.Text = buscar
            txt.ForeColor = Color.Red
            Call consulta()
        End If

    End Sub


    Private Sub TextBox17_Enter(sender As Object, e As EventArgs) Handles TextBox17.Enter
        Dim txt As TextBox = TryCast(sender, TextBox)
        If txt.Text.Equals(buscar) Then
            txt.Text = String.Empty
            txt.ForeColor = SystemColors.WindowText
        End If

    End Sub

    Private Sub TextBox17_Leave(sender As Object, e As EventArgs) Handles TextBox17.Leave
        Dim txt As TextBox = TryCast(sender, TextBox)
        If txt.Text.Trim() = String.Empty Then
            txt.Text = buscar
            txt.ForeColor = Color.Red
            Call consulta()
        End If

    End Sub


    Private Sub TextBox20_Enter(sender As Object, e As EventArgs) Handles TextBox20.Enter
        Dim txt As TextBox = TryCast(sender, TextBox)
        If txt.Text.Equals(buscar) Then
            txt.Text = String.Empty
            txt.ForeColor = SystemColors.WindowText
        End If

    End Sub

    Private Sub TextBox20_Leave(sender As Object, e As EventArgs) Handles TextBox20.Leave
        Dim txt As TextBox = TryCast(sender, TextBox)
        If txt.Text.Trim() = String.Empty Then
            txt.Text = buscar
            txt.ForeColor = Color.Red
            Call consulta()
        End If

    End Sub


    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

    ' abre el excel de guardias
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Process.Start("Excel.exe", "\\192.168.6.21\Bitacora\DOCUMENTOS_AUDITORIA\Guardias.xlsx")
    End Sub

    'abre excel lideres de cuenta
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Process.Start("Excel.exe", "\\192.168.6.21\Bitacora\DOCUMENTOS_AUDITORIA\Lideres.xlsx")
    End Sub


    '+++++++++++++++++++++++++++++++++++++consultas de busquedas++++++++++++++++++++++++++++++++++++++++++++++++++++
    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged
        sp = TextBox10.Text
        DataGridView1.DataSource = Nothing
        dataSt = cn.consulta("select * from Contactos
                              where
                              cliente like'" & "%" & sp & "%" &
                              "'or pais Like'" & "%" & sp & "%" &
                              "'", "contactos")
        DataGridView1.DataSource = dataSt.Tables(0)
        DataGridView1.Columns.Item(0).Visible = False
    End Sub

    Private Sub TextBox17_TextChanged(sender As Object, e As EventArgs) Handles TextBox17.TextChanged
        sp = TextBox17.Text
        DataGridView3.DataSource = Nothing
        dataSt = cn.consulta("select * from Evertec
                              where
                              Nombre like'" & "%" & sp & "%" &
                              "'or Area Like'" & "%" & sp & "%" &
                              "'", "contactos")
        DataGridView3.DataSource = dataSt.Tables(0)
        DataGridView3.Columns.Item(0).Visible = False
    End Sub

    Private Sub TextBox20_TextChanged(sender As Object, e As EventArgs) Handles TextBox20.TextChanged
        sp = TextBox20.Text
        DataGridView2.DataSource = Nothing
        dataSt = cn.consulta("select * from puertos
                              where
                              puerto like'" & sp &
                              "'or Instancia like'" & sp & "%" &
                              "'or Descripcion like'" & "%" & sp & "%" &
                              "'or Socket like'" & sp &
                              "'order by puerto", "contactos")

        DataGridView2.DataSource = dataSt.Tables(0)
        DataGridView2.Columns.Item(0).Visible = False
    End Sub

    '*********************************************************************************************************************************

    Private Sub TextBox10_DoubleClick(sender As Object, e As EventArgs) Handles TextBox10.DoubleClick
        TextBox10.Text = ""
    End Sub
    Private Sub TextBox17_DoubleClick(sender As Object, e As EventArgs) Handles TextBox17.DoubleClick
        TextBox17.Text = ""
    End Sub
    Private Sub TextBox20_DoubleClick(sender As Object, e As EventArgs) Handles TextBox20.DoubleClick
        TextBox20.Text = ""
    End Sub


    Private Sub Textbox20_GotFocus(sender As Object, e As EventArgs) Handles TextBox20.GotFocus

        TextBox6.Text = ""
        TextBox18.Text = ""
        TextBox5.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox19.Text = ""
    End Sub

    Private Sub Textbox17_GotFocus(sender As Object, e As EventArgs) Handles TextBox17.GotFocus

        TextBox9.Text = ""
        TextBox16.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
    End Sub

    Private Sub Textbox10_GotFocus(sender As Object, e As EventArgs) Handles TextBox10.GotFocus
        TextBox14.Text = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox15.Text = ""
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class
