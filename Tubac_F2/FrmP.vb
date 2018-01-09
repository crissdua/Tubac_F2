Imports System.Windows.Forms
Imports System.IO
Imports System.Data.SqlClient
Imports System.Data
Imports System.Drawing.Text
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Windows.Forms
Imports CrystalDecisions.ReportSource
Imports System.Web.UI.WebControls
Public Class FrmP
#Region "Variables"
    Public con As New Conexion
    Dim objectCode As String
    Dim itemcode As String
    Dim oCompany As SAPbobsCOM.Company
    Dim connectionString As String = Conexion.ObtenerConexion.ConnectionString
    Public Shared PO As SAPbobsCOM.Documents
    Public Shared GoodsReceiptPO As SAPbobsCOM.Documents
    Public Shared SQL_Conexion As SqlConnection = New SqlConnection()

#Region "Listas"
    Dim batch As New List(Of String)
    Dim descripcion As New List(Of String)
    Dim anchotira As New List(Of Double)
    Dim pesoreal As New List(Of Double)
    Dim bobina As New List(Of String)
    Dim heat As New List(Of String)
    Dim coil As New List(Of String)
    Dim ordencorte As New List(Of String)
#End Region
#Region "Fuentes"
    Private _Font As Font
    Private PATH_FONTS As String = Application.StartupPath + "\Fonts"
#End Region
    Private Const CP_NOCLOSE_BUTTON As Integer = &H200
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

    Private Sub imprime(bat As String, desc As String, anch As Double, pes As Double, bob As String, het As String, coi As String, ordr As String)
        Dim Report1 As New CrystalDecisions.CrystalReports.Engine.ReportDocument()
        Report1.PrintOptions.PaperOrientation = PaperOrientation.Portrait
        Report1.Load("C:\Users\Cristhiam\Desktop\Informe3.rpt", CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault.OpenReportByDefault)
        Report1.SetParameterValue("CodBatch", bat)
        'Report1.SetParameterValue("CodBatch", txtBarcode.Text)
        Report1.SetParameterValue("descripcion", desc)
        Report1.SetParameterValue("anchotira", anch)
        Report1.SetParameterValue("pesoreal", pes)
        Report1.SetParameterValue("bobina", bob)
        Report1.SetParameterValue("heat", het)
        Report1.SetParameterValue("coil", coi)
        Report1.SetParameterValue("ordencorte", ordr)
        Report1.SetParameterValue("fechacorte", Now.ToShortDateString)
        'CrystalReportViewer1.ReportSource = Report1
        Report1.PrintToPrinter(1, False, 0, 0)
    End Sub
    Private Sub generaEntrada()
        batch.Clear()
        descripcion.Clear()
        anchotira.Clear()
        pesoreal.Clear()
        bobina.Clear()
        heat.Clear()
        coil.Clear()
        ordencorte.Clear()
        Dim iResult As Integer = -1
        Dim iResult2 As Integer = -1
        Dim sResult As String = String.Empty
        Dim sOutput As String = String.Empty
        Dim sql As String
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim sql2 As String
        Dim oRecordSet2 As SAPbobsCOM.Recordset
        Dim docentry As String

        Try
            '------------------OBTIENE DOCENTRY------------
            sql = ("SELECT T0.DocEntry FROM OPOR T0 WHERE T0.DocNum = '" + txtOrder.Text + "'")
            oRecordSet = con.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(sql)
            If oRecordSet.RecordCount > 0 Then
                docentry = oRecordSet.Fields.Item(0).Value
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
            oRecordSet = Nothing
            GC.Collect()

            PO = con.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
            GoodsReceiptPO = con.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)

            '----------------------------------------------
            PO.GetByKey(docentry)
            '----------------------------------------------

            GoodsReceiptPO.CardCode = PO.CardCode
            GoodsReceiptPO.CardName = PO.CardName
            '----------- LINEAS -----------------------------

            Dim itemcode As String
            Dim quantity As Double
            Dim i As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
            Dim existe As Boolean = DGV2.Columns.Cast(Of DataGridViewColumn).Any(Function(x) x.Name = "CHK")
            If existe = False Then
                DGV2.Columns.Add(i)
                i.HeaderText = "CHK"
                i.Name = "CHK"
                i.Width = 32
                i.DisplayIndex = 0
            End If
            Dim result As Integer = MessageBox.Show("Desea Ingresar la Orden?", "Atencion", MessageBoxButtons.YesNoCancel)
            If result = DialogResult.Cancel Then
                MessageBox.Show("Cancelado")
            ElseIf result = DialogResult.No Then
                MessageBox.Show("No se realizara la orden")
            ElseIf result = DialogResult.Yes Then
                For Each row As DataGridViewRow In DGV2.Rows
                    Dim chk As DataGridViewCheckBoxCell = row.Cells("CHK")
                    If chk.Value IsNot Nothing AndAlso chk.Value = True Then

                        PO.Lines.SetCurrentLine(DGV2.Rows(chk.RowIndex).Cells.Item(3).Value.ToString)
                        PO.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Close
                        'GoodsReceiptPO.Lines.ItemCode = DGV2.Rows(chk.RowIndex).Cells.Item(1).Value.ToString
                        'GoodsReceiptPO.Lines.ItemDescription = PO.Lines.ItemDescription
                        GoodsReceiptPO.Lines.Quantity = DGV2.Rows(chk.RowIndex).Cells.Item(2).Value.ToString
                        GoodsReceiptPO.Lines.BaseEntry = PO.DocEntry
                        GoodsReceiptPO.Lines.BaseLine = DGV2.Rows(chk.RowIndex).Cells.Item(3).Value.ToString
                        GoodsReceiptPO.Lines.BaseType = Convert.ToInt32(PO.DocObjectCodeEx)
                        GoodsReceiptPO.Lines.BatchNumbers.SetCurrentLine(0)
                        GoodsReceiptPO.Lines.BatchNumbers.BatchNumber = "batchIMP" & DGV2.Rows(chk.RowIndex).Cells.Item(1).Value.ToString
                        '-----------------------------------------------------------------------------llena listas de datos
                        batch.Add("batchIMP" & DGV2.Rows(chk.RowIndex).Cells.Item(1).Value.ToString)
                        descripcion.Add(PO.Lines.ItemDescription)
                        anchotira.Add(PO.Lines.UserFields.Fields.Item("U_ancho").Value)
                        pesoreal.Add(PO.Lines.UserFields.Fields.Item("U_peso").Value)
                        bobina.Add(PO.Lines.UserFields.Fields.Item("U_bobina").Value)
                        heat.Add(PO.Lines.UserFields.Fields.Item("U_heat").Value)
                        coil.Add(PO.Lines.UserFields.Fields.Item("U_coil").Value)
                        ordencorte.Add(PO.DocNum)
                        '---------------------------------------------------------------------------------------------------
                        GoodsReceiptPO.Lines.BatchNumbers.Quantity = Convert.ToDouble(DGV2.Rows(chk.RowIndex).Cells.Item(2).Value.ToString)
                        GoodsReceiptPO.Lines.BatchNumbers.AddmisionDate = Now
                        GoodsReceiptPO.Lines.BatchNumbers.Add()


                        GoodsReceiptPO.Lines.Add()
                    End If
                Next

            End If
            '---------------------------------------Ingresa Mercaderia----------------------
            GoodsReceiptPO.Comments = PO.DocEntry
            iResult = GoodsReceiptPO.Add()
            If iResult <> 0 Then
                MessageBox.Show(con.oCompany.GetLastErrorDescription)
            Else
                PO.Comments = PO.DocEntry
                iResult2 = PO.Update() '---------------------------- Actualiza el pedido (las lineas del pedido)
                If iResult2 <> 0 Then
                    MessageBox.Show(con.oCompany.GetLastErrorDescription)
                End If
                '-------------------------------IMPRIME BATCH--------------------------------------
                'bat As String, desc As String, anch As Double, pes As Double, bob As String, het As String, coi As String, ordr As String
                Dim cont As Integer
                For cont = 0 To batch.Count - 1
                    'imprime(FormatBarCode(batch.Item(cont)), descripcion(cont), anchotira(cont), pesoreal(cont), bobina(cont), heat(cont), coil(cont), ordencorte(cont))
                Next
                '-----------------------------------------------------------------------------------
            End If
            con.oCompany.Disconnect()
        Catch ex As Exception
            MsgBox("Error: " + ex.Message.ToString)
            con.oCompany.Disconnect()
        End Try
    End Sub
    Private Sub GR_from_PO()
        Try
            If con.MakeConnectionSAP() Then
                generaEntrada()
            Else
                con.MakeConnectionSAP()
                If con.Connected Then
                    generaEntrada()
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
            PO = Nothing
            GoodsReceiptPO = Nothing
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
            Application.Exit()
            Me.Close()
        End If
    End Sub
End Class
