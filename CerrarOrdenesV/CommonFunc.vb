'//****************************************************************************
'//
'// DESCRIPCION FUNCIONAL: Funciones comunes a todos los objetos
'//
'//****************************************************************************
Option Strict Off
Option Explicit On
Imports System.Windows.Forms
Imports Microsoft.Win32

Module CommonFunc
    '//----- OBTIENE UN QUERY CON SOLO UN PARAMETRO
    '//----- CARGA UNA FORMA A PARTIR DE UN ARCHIVO XML
    Public Function loadFormXML(ByVal SBOApplication As SAPbouiCOM.Application,
                                ByVal psFormUID As String, ByVal psFile As String) As Integer
        Dim loXMLDoc As MSXML2.DOMDocument
        Dim loForm As SAPbouiCOM.Form
        Try
            loXMLDoc = New MSXML2.DOMDocument
            '//BUSCA LA FORMA
            loForm = searchForm(SBOApplication, psFormUID)
            If (loForm Is Nothing) Then
                '//CARGA EL XML
                If (Not loXMLDoc.load(psFile)) Then
                    SBOApplication.MessageBox("No pudo abrir el archivo " & psFile)
                    Return -1
                End If
                '//CARGA LA FORMA EN SBO
                'MsgBox(loXMLDoc.childNodes.item(1).childNodes.item(0).childNodes.item(0).childNodes.item(0).attributes.getNamedItem("uid").xml.ToString)
                'loXMLDoc.childNodes.item(1).childNodes.item(0).childNodes.item(0).childNodes.item(0).attributes.getNamedItem("FormType").nodeValue = "Chidey"
                'loXMLDoc.childNodes.item(1).childNodes.item(0).childNodes.item(0).childNodes.item(0).attributes.getNamedItem("uid").nodeValue = "Chidey"
                ' MsgBox(loXMLDoc.childNodes.item(1).childNodes.item(0).childNodes.item(0).childNodes.item(0).attributes.getNamedItem("uid").xml.ToString)

                SBOApplication.LoadBatchActions(loXMLDoc.xml)
                loadFormXML = 0
            End If
            '//MUEVE EL FOCO A LA FORMA
            loForm = SBOApplication.Forms.Item(psFormUID)
            loForm.Select()
        Catch ex As Exception
            SBOApplication.MessageBox("CommonFunc. Error al abrir la forma " & psFormUID & ". " & ex.Message)
            Return -1
        End Try
    End Function

    '//----- BUSCA UNA FORMA INDICADA EN LA APLICACION
    Public Function searchForm(ByVal SBOApplication As SAPbouiCOM.Application,
                               ByVal psFormUID As String) As SAPbouiCOM.Form
        Try
            searchForm = SBOApplication.Forms.Item(psFormUID)
        Catch ex As Exception
            searchForm = Nothing
        End Try
    End Function

    '//----- LEE EL PATH INDICADO
    Public Function ReadPath(ByVal psApplName As String) As String
        Dim sAns As String
        Dim sErr As String = ""
        sAns = My.Application.Info.DirectoryPath
        'sAns = RegValue(RegistryHive.CurrentUser, "BBConsulting", psApplName, sErr)
        ReadPath = sAns
        If Not (sAns <> "") Then
            MessageBox.Show("CommonFunc. Al obtener el valor del registro. " & sErr)
            ReadPath = ""
        End If
    End Function

    Public Function SQLSentence(ByVal plIdx As Long, ByVal psCambio1 As String) As String
        Try
            SQLSentence = SQLSentence(plIdx, psCambio1, "")
        Catch ex As Exception
            SQLSentence = ""
        End Try
    End Function
    '//----- OBTIENE UN QUERY CON DOS PARAMETROS
    Public Function SQLSentence(ByVal plIdx As Long, ByVal psCambio1 As String, ByVal psCambio2 As String) As String
        Try
            SQLSentence = SQLSentence(plIdx, psCambio1, psCambio2, "")
        Catch ex As Exception
            SQLSentence = ""
        End Try
    End Function
    '//----- OBTIENE UN QUERY CON TRES PARAMETROS
    Public Function SQLSentence(ByVal plIdx As Long, ByVal psCambio1 As String, ByVal psCambio2 As String, ByVal psCambio3 As String) As String
        Try
            SQLSentence = SQLSentence(plIdx, psCambio1, psCambio2, psCambio3, "")
        Catch ex As Exception
            SQLSentence = ""
        End Try
    End Function
    '//----- OBTIENE UN QUERY CON CUATRO PARAMETROS
    Public Function SQLSentence(ByVal plIdx As Long, ByVal psCambio1 As String, ByVal psCambio2 As String, ByVal psCambio3 As String, ByVal psCambio4 As String) As String
        Try
            SQLSentence = SQLSentence(plIdx, psCambio1, psCambio2, psCambio3, psCambio4, "")
        Catch ex As Exception
            SQLSentence = ""
        End Try
    End Function
    '//----- OBTIENE UN QUERY CON 5 PARAMETROS
    Public Function SQLSentence(ByVal plIdx As Long, ByVal psCambio1 As String, ByVal psCambio2 As String, ByVal psCambio3 As String,
    ByVal psCambio4 As String, ByVal psCambio5 As String) As String
        Try
            SQLSentence = SQLSentence(plIdx, psCambio1, psCambio2, psCambio3, psCambio4, psCambio5, "")
        Catch ex As Exception
            SQLSentence = ""
        End Try
    End Function
    '//----- OBTIENE UN QUERY CON 6 PARAMETROS
    Public Function SQLSentence(ByVal plIdx As Long, ByVal psCambio1 As String, ByVal psCambio2 As String, ByVal psCambio3 As String,
    ByVal psCambio4 As String, ByVal psCambio5 As String, ByVal psCambio6 As String) As String
        Try
            SQLSentence = SQLSentence(plIdx, psCambio1, psCambio2, psCambio3, psCambio4, psCambio5, psCambio6, "")
        Catch ex As Exception
            SQLSentence = ""
        End Try
    End Function
    '//----- OBTIENE UN QUERY CON 7 PARAMETROS
    Public Function SQLSentence(ByVal plIdx As Long, ByVal psCambio1 As String, ByVal psCambio2 As String, ByVal psCambio3 As String,
    ByVal psCambio4 As String, ByVal psCambio5 As String, ByVal psCambio6 As String, ByVal psCambio7 As String) As String
        Try
            SQLSentence = "BBCsp_execAgregarAdenda " & plIdx & ", '" & psCambio1 & "', '" &
            psCambio2 & "', '" & psCambio3 & "', '" & psCambio4 & "', '" &
            psCambio5 & "', '" & psCambio6 & "', '" & psCambio7 & "'"
        Catch ex As Exception
            SQLSentence = ""
        End Try
    End Function
End Module