Imports System
Imports System.Xml
Imports System.Windows.Forms
Imports System.IO
Imports System.Data.SqlClient
Public Class Conexion
    Public Connected As Boolean
    Public connectionString As String
    Public Property _oCompany As SAPbobsCOM.Company
    Public Property oCompany() As SAPbobsCOM.Company
        Get
            Return _oCompany
        End Get
        Set(ByVal value As SAPbobsCOM.Company)
            _oCompany = value
        End Set
    End Property
    Public Function MakeConnectionSAP() As Boolean
        Connected = False
        Try
            Connected = -1
            oCompany = New SAPbobsCOM.Company
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
            Dim entra As String = Application.StartupPath + "\conexion.xml"
            Dim Xml As XmlDocument = New XmlDocument()
            Xml.Load(entra)
            Dim ArticleList As XmlNodeList = Xml.SelectNodes("body/HANA")
            For Each xnDoc As XmlNode In ArticleList
                oCompany.UserName = xnDoc.SelectSingleNode("UserName").InnerText
                oCompany.Password = xnDoc.SelectSingleNode("Password").InnerText
                oCompany.DbUserName = xnDoc.SelectSingleNode("DbUserName").InnerText
                oCompany.DbPassword = xnDoc.SelectSingleNode("DbPassword").InnerText
                oCompany.Server = xnDoc.SelectSingleNode("Server").InnerText
                oCompany.CompanyDB = xnDoc.SelectSingleNode("CompanyDB").InnerText
                oCompany.LicenseServer = xnDoc.SelectSingleNode("Server").InnerText + ":30000"
            Next

            oCompany.UseTrusted = False
            Connected = oCompany.Connect()

            If Connected <> 0 Then
                Connected = False
                MsgBox(oCompany.GetLastErrorDescription)
            Else
                Connected = True
            End If
            Return Connected
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Shared Function ObtenerConexion() As SqlConnection
        Dim ip As String
        Dim db As String
        Dim user As String
        Dim password As String
        Dim entra As String = Application.StartupPath + "\Conexion.xml"
        Dim Xml As XmlDocument = New XmlDocument()
        Xml.Load(entra)
        Dim ArticleList As XmlNodeList = Xml.SelectNodes("body/SQL")
        For Each xnDoc As XmlNode In ArticleList
            ip = xnDoc.SelectSingleNode("ip").InnerText
            db = xnDoc.SelectSingleNode("db").InnerText
            user = xnDoc.SelectSingleNode("user").InnerText
            password = xnDoc.SelectSingleNode("password").InnerText
        Next

        'Dim connectionString As String = "Server=DESKTOP-G67F328;Database=2016Tubex_Productiva;Trusted_Connection=True;" '"Server=" + ip + ";" + "Database=" + db + ";" + "User id=" + user + ";" + "Password=" + password + ";"
        Dim connectionString As String = "Server=" + ip + ";" + "Database=" + db + ";" + "User id=" + user + ";" + "Password=" + password + ";"
        Dim con As SqlConnection = New SqlConnection(connectionString)
        con.Open()
        Return con
    End Function
End Class
