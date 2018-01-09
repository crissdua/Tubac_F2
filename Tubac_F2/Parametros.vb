Imports System.Xml
Imports System.Windows.Forms
Imports System.IO
Public Class Parametros
    Dim oCompany As SAPbobsCOM.Company = Login.con.oCompany
    Dim Connected As Boolean = Login.con.Connected
    Private Sub Parametros_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim entra As String = Application.StartupPath + "\conexion.xml"
        Dim Xml As XmlDocument = New XmlDocument()
        Xml.Load(entra)
        Dim ArticleList As XmlNodeList = Xml.SelectNodes("body/SQL")
        For Each xnDoc As XmlNode In ArticleList
            TextBox1.Text = xnDoc.SelectSingleNode("DbUserName").InnerText
            TextBox2.Text = xnDoc.SelectSingleNode("DbPassword").InnerText
            TextBox3.Text = xnDoc.SelectSingleNode("Server").InnerText
            TextBox4.Text = xnDoc.SelectSingleNode("CompanyDB").InnerText
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim entra As String = Application.StartupPath + "\conexion.xml"
        Dim Xml As XmlDocument = New XmlDocument()
        Xml.Load(entra)
        Dim ArticleList As XmlNodeList = Xml.SelectNodes("body/SQL")
        For Each xnDoc As XmlNode In ArticleList
            xnDoc.ChildNodes(0).InnerText = TextBox1.Text
            xnDoc.ChildNodes(1).InnerText = TextBox2.Text
            xnDoc.ChildNodes(2).InnerText = TextBox3.Text
            xnDoc.ChildNodes(3).InnerText = TextBox4.Text
        Next
        Xml.Save(entra)
        MessageBox.Show("Parametros Guardados Exitosamente")
        Me.Hide()
        Dim frm As New Login
        frm.Show()
        GC.Collect()
        Connected = False
        oCompany.Disconnect()
    End Sub
End Class