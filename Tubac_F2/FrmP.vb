Imports System.Windows.Forms
Imports System.IO
Imports System.Data.SqlClient
Imports Tubac_F2.BatchsFase2
Imports SAPbobsCOM
Imports SAPbouiCOM

Public Class FrmP
#Region "Variables"
    Public con As New Conexion
    Public itemcode As String
    Dim valBarCode As String
    Dim oCompany As SAPbobsCOM.Company
    Dim connectionString As String = Conexion.ObtenerConexion.ConnectionString
    Public Ready As Boolean
    Private Const CP_NOCLOSE_BUTTON As Integer = &H200
    Public Shared oInvGenExit As SAPbobsCOM.Documents
    Public Shared SQL_Conexion As SqlConnection = New SqlConnection()
    Public Shared ba As New List(Of String)
    Public Shared quantity As New List(Of Integer)

#End Region
    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim myCp As CreateParams = MyBase.CreateParams
            myCp.ClassStyle = myCp.ClassStyle Or CP_NOCLOSE_BUTTON
            Return myCp
        End Get
    End Property
    Public Sub New(ByVal user As String)
        MyBase.New()
        InitializeComponent()
        '  Note which form has called this one
        ToolStripStatusLabel1.Text = user
    End Sub
    Private Sub FrmFase1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox2.Select()
        cargaORDER()
    End Sub
    Public Function cargaORDER()
        Dim SQL_da As SqlDataAdapter = New SqlDataAdapter("SELECT T0.DocNum FROM OWOR T0 where T0.Type = 'P' and T0.Status = 'R'", con.ObtenerConexion())
        Dim DT_dat As System.Data.DataTable = New System.Data.DataTable()
        SQL_da.Fill(DT_dat)
        DGV.DataSource = DT_dat
        con.ObtenerConexion.Close()
    End Function



    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Dim SQL_da As SqlDataAdapter = New SqlDataAdapter("SELECT T0.DocNum FROM OWOR T0 where T0.Type = 'P' and T0.Status = 'R' and T0.DocNum LIKE '" + TextBox2.Text + "%' ORDER BY T0.DocNum", con.ObtenerConexion())
        Dim DT_dat As System.Data.DataTable = New System.Data.DataTable()
        SQL_da.Fill(DT_dat)
        DGV.DataSource = DT_dat
        con.ObtenerConexion.Close()
    End Sub

    Private Sub generaEntrada(docE As String)

        Dim iResult As Integer = -1
        Dim contlinea As Integer
        Dim sql As String
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim oReturn As Integer = -1

        Try
            '--------------------------------------------------------------------------------------------------------------
            'EMISION DE PRODUCCION -----------------------------------------------------------------------------------------
            '--------------------------------------------------------------------------------------------------------------
            'Contador de lineas para la orden
            sql = ("SELECT count(T0.DocNum) FROM OWOR T0  INNER JOIN WOR1 T1 ON T0.DocEntry = T1.DocEntry WHERE T0.DocNum = '" + docE + "'")
            oRecordSet = con.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(sql)
            If oRecordSet.RecordCount > 0 Then
                contlinea = oRecordSet.Fields.Item(0).Value
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
            oRecordSet = Nothing
            GC.Collect()

            Dim frm As New Tubac_F2.BatchsFase2
            oInvGenExit = con.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
            Dim wo As ProductionOrders = con.oCompany.GetBusinessObject(BoObjectTypes.oProductionOrders)
            wo.GetByKey(docE)

            Dim i As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
            Dim existe As Boolean = DGV2.Columns.Cast(Of DataGridViewColumn).Any(Function(x) x.Name = "CHK")
            If existe = False Then
                DGV2.Columns.Add(i)
                i.HeaderText = "CHK"
                i.Name = "CHK"
                i.Width = 32
                i.DisplayIndex = 0
            End If
            Dim result As Integer = MessageBox.Show("Desea Ingresar la Emision?", "Atencion", MessageBoxButtons.YesNoCancel)
            If result = DialogResult.Cancel Then
                MessageBox.Show("Cancelado")
            ElseIf result = DialogResult.No Then
                MessageBox.Show("No se realizara la orden")
            ElseIf result = DialogResult.Yes Then

                oInvGenExit.DocDate = Now ' it could be date.Today.
                oInvGenExit.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items

                Dim contbase As Integer = 0
                For Each row As DataGridViewRow In DGV2.Rows
                    Dim chk As DataGridViewCheckBoxCell = row.Cells("CHK")
                    If chk.Value IsNot Nothing AndAlso chk.Value = True Then
                        ba.Clear()
                        quantity.Clear()
                        oInvGenExit.Lines.BaseType = 202
                        oInvGenExit.Lines.BaseEntry = docE
                        oInvGenExit.Lines.BaseLine = contbase
                        oInvGenExit.Lines.AccountCode = "_SYS00000000039"
                        WO.Lines.SetCurrentLine(contbase)
                        frm.load(DGV2.Rows(chk.RowIndex).Cells.Item(1).Value.ToString, wo.Lines.PlannedQuantity, docE)
                        frm.ShowDialog()
                        Dim cont As Integer
                        For cont = 0 To ba.Count - 1
                            oInvGenExit.Lines.BatchNumbers.BatchNumber = ba.Item(cont)
                            oInvGenExit.Lines.BatchNumbers.Quantity = quantity.Item(cont)
                            oInvGenExit.Lines.BatchNumbers.Add()
                        Next
                        contbase = contbase + 1
                        oInvGenExit.Lines.Add()
                    End If
                Next
                oReturn = oInvGenExit.Add
                If oReturn <> 0 Then
                    MessageBox.Show(con.oCompany.GetLastErrorDescription)
                End If
            End If

        Catch ex As Exception
            MsgBox("Error: " + ex.Message.ToString)
        End Try
    End Sub
    Private Sub GR_from_PO()
        Try
            If con.MakeConnectionSAP() Then
                generaEntrada(txtOrder.Text)
            Else
                con.MakeConnectionSAP()
                If con.Connected Then
                    generaEntrada(txtOrder.Text)
                Else
                    MessageBox.Show("Error de Conexion, intente Nuevamente")
                End If
            End If
        Catch ex As Exception
            MsgBox("Error: " + ex.Message.ToString)
        End Try
    End Sub

    Private Sub DGV_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV.CellContentClick
        txtOrder.Text = DGV(0, DGV.CurrentCell.RowIndex).Value.ToString()
    End Sub

    Private Sub txtOrder_TextChanged(sender As Object, e As EventArgs) Handles txtOrder.TextChanged
        Dim i As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
        Dim existe As Boolean = DGV2.Columns.Cast(Of DataGridViewColumn).Any(Function(x) x.Name = "CHK")
        If existe = False Then
            DGV2.Columns.Add(i)
            i.HeaderText = "CHK"
            i.Name = "CHK"
            i.Width = 32
            i.DisplayIndex = 0
        End If

        Dim SQL_da As SqlDataAdapter = New SqlDataAdapter("SELECT T0.ItemCode, T0.BaseQty, isnull(T0.LineNum,0) FROM WOR1 T0 where T0.[DocEntry] like '" + txtOrder.Text + "%'", con.ObtenerConexion())
        Dim DT_dat As System.Data.DataTable = New System.Data.DataTable()
        SQL_da.Fill(DT_dat)
        DGV2.DataSource = DT_dat
        con.ObtenerConexion.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        GR_from_PO()
        DGV.DataSource = Nothing
        DGV2.DataSource = Nothing
        TextBox2.Clear()
        txtOrder.Clear()
    End Sub

    Private Sub btnFinalizar_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim result As Integer = MessageBox.Show("Desea limpiar el objeto?", "Atencion", MessageBoxButtons.YesNoCancel)
        If result = DialogResult.Cancel Then
            MessageBox.Show("Cancelado")
        ElseIf result = DialogResult.No Then
            MessageBox.Show("Puede continuar!")
        ElseIf result = DialogResult.Yes Then
            TextBox2.Clear()
            txtOrder.Clear()
            DGV.DataSource = Nothing
            DGV2.DataSource = Nothing
            MessageBox.Show("Inicie un objeto nuevo")
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim result As Integer = MessageBox.Show("Desea salir del modulo?", "Atencion", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            MessageBox.Show("Puede continuar")
        ElseIf result = DialogResult.Yes Then
            MessageBox.Show("Finalizando modulo")
            Try
                con.oCompany.Disconnect()
            Catch
            End Try
            System.Windows.Forms.Application.Exit()
            Me.Close()
        End If
    End Sub
End Class
