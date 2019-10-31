Imports System.Windows.Forms

Public Class CerrarOV
    'Dim oLinq_InterfazEDI As LINQ_InterfazEDIDataContext '--- contexto de Base de Fltes
    Private cSBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private cSBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION
    Private coForm As SAPbouiCOM.Form           '//FORMA
    Private oGrid As SAPbouiCOM.Grid
    Private csFormUID As String
    Private csDirectory As String
    Private oItem As SAPbouiCOM.Item
    Private oDBDataSource As SAPbouiCOM.DBDataSource
    Private oUserDataSource As SAPbouiCOM.UserDataSource
    Private oColumn As SAPbouiCOM.GridColumn
    Public Fecha1, Fecha2, Serie As String
    Private stFecha1, stFecha2, stSerie As String

    Public Sub New()
        MyBase.New()
        cSBOApplication = oCatchingEvents.SBOApplication
        cSBOCompany = oCatchingEvents.SBOCompany
        csDirectory = oCatchingEvents.csDirectory
        '--- crea objeto LinqContext
        'oLinq_InterfazEDI = New LINQ_InterfazEDIDataContext

    End Sub

    '/// Se Crea la Forma
    Public Sub CreateForm()

        Dim CP As SAPbouiCOM.FormCreationParams
        Dim oStat As SAPbouiCOM.StaticText
        Dim oBtn As SAPbouiCOM.Button
        Dim lsItemRef As String
        Dim loDS As SAPbouiCOM.UserDataSource
        Dim oEditTxt As SAPbouiCOM.EditText
        Dim loForm As SAPbouiCOM.Form
        Dim oComboBox As SAPbouiCOM.ComboBox
        Dim stQueryS, serie, serien As String
        Dim oRecSetS As SAPbobsCOM.Recordset

        Try

            loForm = searchForm(cSBOApplication, "frmCOVSA2")
            If (loForm Is Nothing) Then

                CP = cSBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
                CP.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
                'CP.FormType = 60004
                CP.UniqueID = "frmCOVSA2"

                coForm = cSBOApplication.Forms.AddEx(CP)


                '   Define ancho y largo de la forma
                coForm.Height = 350
                coForm.Width = 800
                coForm.Title = "Cierre de Ordenes de Ventas"

                ' Agrega texto 1
                oItem = coForm.Items.Add("StatHead", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                oItem.Left = 20
                oItem.Top = 20
                oItem.Width = 60

                oStat = oItem.Specific
                oStat.Caption = "Desde: "

                'Fecha1
                lsItemRef = "StatHead"
                oItem = coForm.Items.Add("Campodate1", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                oItem.Left = coForm.Items.Item(lsItemRef).Left + coForm.Items.Item(lsItemRef).Width + 10
                oItem.Top = coForm.Items.Item(lsItemRef).Top
                oItem.Width = 80
                oItem.Height = coForm.Items.Item(lsItemRef).Height

                loDS = coForm.DataSources.UserDataSources.Add("dsFecha1", SAPbouiCOM.BoDataType.dt_DATE)
                oEditTxt = coForm.Items.Item("Campodate1").Specific
                oEditTxt.DataBind.SetBound(True, "", "dsFecha1")

                ' Agrega texto 2
                lsItemRef = "Campodate1"
                oItem = coForm.Items.Add("StatHeadH", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                oItem.Left = coForm.Items.Item(lsItemRef).Left + coForm.Items.Item(lsItemRef).Width + 15
                oItem.Top = coForm.Items.Item(lsItemRef).Top
                oItem.Width = 60
                oItem.Height = coForm.Items.Item(lsItemRef).Height

                oStat = oItem.Specific
                oStat.Caption = "Hasta: "

                'Fecha2
                lsItemRef = "StatHeadH"
                oItem = coForm.Items.Add("Campodate2", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                oItem.Left = coForm.Items.Item(lsItemRef).Left + coForm.Items.Item(lsItemRef).Width + 10
                oItem.Top = coForm.Items.Item(lsItemRef).Top
                oItem.Width = 80
                oItem.Height = coForm.Items.Item(lsItemRef).Height

                loDS = coForm.DataSources.UserDataSources.Add("dsFecha2", SAPbouiCOM.BoDataType.dt_DATE)
                oEditTxt = coForm.Items.Item("Campodate2").Specific
                oEditTxt.DataBind.SetBound(True, "", "dsFecha2")


                ' Agrega texto 3
                lsItemRef = "Campodate2"
                oItem = coForm.Items.Add("StatHeadS", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                oItem.Left = coForm.Items.Item(lsItemRef).Left + coForm.Items.Item(lsItemRef).Width + 15
                oItem.Top = coForm.Items.Item(lsItemRef).Top
                oItem.Width = 60
                oItem.Height = coForm.Items.Item(lsItemRef).Height

                oStat = oItem.Specific
                oStat.Caption = "Serie: "


                'Series
                lsItemRef = "StatHeadS"
                oItem = coForm.Items.Add("Series", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                oItem.Left = coForm.Items.Item(lsItemRef).Left + coForm.Items.Item(lsItemRef).Width + 10
                oItem.Top = coForm.Items.Item(lsItemRef).Top
                oItem.Width = 80
                oItem.Height = coForm.Items.Item(lsItemRef).Height

                oItem.DisplayDesc = False

                oComboBox = oItem.Specific

                '// bind the Combo Box item to the defined used data source
                loDS = coForm.DataSources.UserDataSources.Add("CombSource", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                oComboBox.DataBind.SetBound(True, "", "CombSource")

                oRecSetS = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                stQueryS = "Select T0.""Series"",T0.""SeriesName"" from ""NNM1"" T0 where T0.""ObjectCode""=17 group by T0.""Series"",T0.""SeriesName"" order by T0.""Series"""
                oRecSetS.DoQuery(stQueryS)

                If oRecSetS.RecordCount > 0 Then

                    For i = 1 To oRecSetS.RecordCount

                        If i = 1 Then

                            oRecSetS.MoveFirst()
                            serie = oRecSetS.Fields.Item("Series").Value
                            serien = oRecSetS.Fields.Item("SeriesName").Value
                            oComboBox.ValidValues.Add(serie, serien)
                            oRecSetS.MoveNext()

                        Else

                            serie = oRecSetS.Fields.Item("Series").Value
                            serien = oRecSetS.Fields.Item("SeriesName").Value
                            oComboBox.ValidValues.Add(serie, serien)
                            oRecSetS.MoveNext()

                        End If

                    Next

                End If

                'Agrega Boton Buscar

                lsItemRef = "Series"
                oItem = coForm.Items.Add("btBuscar", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                oItem.Left = coForm.Items.Item(lsItemRef).Left + coForm.Items.Item(lsItemRef).Width + 50
                oItem.Top = coForm.Items.Item(lsItemRef).Top
                oItem.Width = 80
                oItem.Height = coForm.Items.Item(lsItemRef).Height
                oBtn = oItem.Specific
                oBtn.Caption = "Buscar"
                oItem.Visible = True


                ' Add a Grid item to the form

                oItem = coForm.Items.Add("MyGrid", SAPbouiCOM.BoFormItemTypes.it_GRID)
                ' Set the grid dimentions and position
                oItem.Left = 20
                oItem.Top = 50
                oItem.Width = 650
                oItem.Height = 200

                ' Set the grid data
                oGrid = oItem.Specific

                coForm.DataSources.DataTables.Add("MyDataTable")
                coForm.DataSources.DataTables.Item(0).ExecuteQuery("Select '' as ""Check"",'' as ""Ref"",'' as ""DocNum"",'' as ""Fecha"",'' as ""Cliente"",'' as ""Nombre de Cliente"",'' as ""Total"" from dummy;")
                oGrid.DataTable = coForm.DataSources.DataTables.Item("MyDataTable")

                ocultarColumnas("Ref")

                oGrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox

                ' Set columns size
                oGrid.Columns.Item(0).Width = 40
                oGrid.Columns.Item(1).Width = 70
                oGrid.Columns.Item(2).Width = 50
                oGrid.Columns.Item(3).Width = 70
                oGrid.Columns.Item(4).Width = 100
                oGrid.Columns.Item(5).Width = 300
                oGrid.Columns.Item(6).Width = 70

                'Agrega Boton Buscar

                lsItemRef = "btBuscar"
                oItem = coForm.Items.Add("btCerrar", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                oItem.Left = 20
                oItem.Top = 270
                oItem.Width = 80

                oBtn = oItem.Specific
                oBtn.Caption = "Cerrar Docs."
                oItem.Visible = True

                SaveAsXML()

            End If

        Catch ex As Exception

            cSBOApplication.MessageBox("CreateForm: " & ex.Message)

        End Try


    End Sub

    Private Sub SaveAsXML()

        Dim oXmlDoc As Xml.XmlDocument

        oXmlDoc = New Xml.XmlDocument

        Dim sXmlString As String
        Dim csFormUID As String

        csFormUID = "frmCOVSA2"

        '// get the form as an XML string
        sXmlString = coForm.GetAsXML

        '// load the form's XML string to the
        '// XML document object
        oXmlDoc.LoadXml(sXmlString)

        '// save the XML Document
        Dim sPath As String

        sPath = IO.Directory.GetParent(Application.StartupPath).ToString

        oXmlDoc.Save((sPath & "\Debug\frmCOVSA2.srf"))


    End Sub

    Public Sub openForm(ByVal psDirectory As String)

        Try
            csFormUID = "frmCOVSA2"

            '//CARGA LA FORMA
            If (loadFormXML(cSBOApplication, csFormUID, psDirectory + "\" + csFormUID + ".srf") <> 0) Then
                Err.Raise(-1, 1, "")
            End If

            '//ESTABLECE LOS DATASOURCES
            If (setForm(csFormUID) <> 0) Then
                Err.Raise(-1, 1, "")
            End If

            coForm.Refresh()
            coForm.Visible = True

            'Me.getDatosIniciales()

        Catch ex As Exception
            If (ex.Message <> "") Then
                cSBOApplication.MessageBox("openForm: No se pudo iniciar la forma. " & ex.Message)
            End If
            Me.close()
        End Try
    End Sub

    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '//----- CIERRA LA VENTANA
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Public Function close() As Integer
        coForm.Close()
    End Function

    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '//----- GUARDA LA REFERENCIA DE LA FORMA
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Public Function setForm(ByVal psFormUID As String) As Integer
        Try
            setForm = 0
            '//ESTABLECE LA REFERENCIA A LA FORMA
            coForm = cSBOApplication.Forms.Item(psFormUID)
            '//OBTIENE LA REFERENCIA A LOS USER DATA SOURCES
            setForm = getUserDataSources()
        Catch ex As Exception
            cSBOApplication.Box("Al referenciar a la forma. " & ex.Message)
            setForm = -1
        End Try
    End Function

    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '//----- OBTIENE LA REFERENCIA A LOS USERDATASOURCES
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Function getUserDataSources() As Integer
        Try
            coForm.Freeze(True)
            getUserDataSources = 0
            '//SI YA EXISTEN LOS DATASOURCES, SOLO LOS ASOCIA
            If (coForm.DataSources.UserDataSources.Count() > 0) Then
            Else '//EN CASO DE QUE NO EXISTAN, LOS CREA

            End If
            coForm.Freeze(False)
        Catch ex As Exception
            cSBOApplication.MessageBox("Al referenciar los UserDataSources" & ex.Message)
            getUserDataSources = -1
        End Try
    End Function

    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '//----- ASOCIA LOS USERDATA A LA MATRIZ
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Function bindUserDataSources() As Integer
        Dim loText As SAPbouiCOM.EditText
        Dim oCheckBox As SAPbouiCOM.CheckBox
        Dim loDS As SAPbouiCOM.UserDataSource
        Dim oCombo As SAPbouiCOM.ComboBox

        Try
            bindUserDataSources = 0

        Catch ex As Exception
            cSBOApplication.MessageBox("Al crear los UserDataSources. " & ex.Message)
            bindUserDataSources = -1
        Finally
            loText = Nothing
            loDS = Nothing
            oCheckBox = Nothing
        End Try
    End Function

    Public Function getDatosIniciales()
        Dim oRecSet As SAPbobsCOM.Recordset
        Dim oCombo As SAPbouiCOM.ComboBox
        Try

        Catch ex As Exception
            cSBOApplication.MessageBox("getDatosIniciales. " & ex.Message)
        Finally
            oRecSet = Nothing
            oCombo = Nothing
        End Try
    End Function

    Private Function getFromList()
        Dim oRecSet As SAPbobsCOM.Recordset
        Dim oCombo As SAPbouiCOM.ComboBox
        Dim i As Integer
        Try

        Catch ex As Exception
            cSBOApplication.MessageBox("getFromList. " & ex.Message)
        Finally
            oRecSet = Nothing
        End Try
    End Function

    Public Function DatosGrid() As String

        Try

            coForm = cSBOApplication.Forms.Item("frmCOVSA2")

            stFecha1 = coForm.DataSources.UserDataSources.Item("dsFecha1").Value
            stFecha2 = coForm.DataSources.UserDataSources.Item("dsFecha2").Value
            stSerie = coForm.DataSources.UserDataSources.Item("CombSource").Value

            'MsgBox(stFecha1 & " " & stFecha2 & " " & stSerie)

            If (stFecha1 = "") And (stFecha2 = "") And (stSerie = "") Then

                oGrid = coForm.Items.Item("MyGrid").Specific
                coForm.DataSources.DataTables.Item(0).ExecuteQuery("Select '' as ""Check"",T1.""DocEntry"" as ""Ref"",T1.""DocNum"",TO_DATE(T1.""DocDate"") as ""Fecha"",T1.""CardCode"" as ""Cliente"",T2.""CardName"" as ""Nombre de Cliente"",T1.""DocTotal"" as ""Total"" from ""ORDR"" T1 inner join ""OCRD"" T2 on T2.""CardCode""=T1.""CardCode"" inner join ""RDR1"" T3 on T3.""DocEntry""=T1.""DocEntry"" inner join ""OITM"" T4 on T4.""ItemCode""=T3.""ItemCode"" inner join ""OITW"" T5 on T5.""WhsCode""=T3.""WhsCode"" and T5.""ItemCode""=T4.""ItemCode"" where T1.""DocStatus""='O' and T4.""validFor""='Y' and T5.""Locked""='N' and T1.""U_tekStatus""='-1' group by T1.""DocEntry"",T1.""DocNum"",T1.""DocDate"",T1.""CardCode"",T2.""CardName"",T1.""DocTotal"" order by T1.""DocDate"",T1.""DocNum""")
                oGrid.DataTable = coForm.DataSources.DataTables.Item("MyDataTable")

                ocultarColumnas("Ref")

                oColumn = oGrid.Columns.Item("Ref")
                oColumn.LinkedObjectType = 17

                oGrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox

                ' Set columns size
                oGrid.Columns.Item(0).Width = 40
                oGrid.Columns.Item(1).Width = 70
                oGrid.Columns.Item(2).Width = 50
                oGrid.Columns.Item(3).Width = 70
                oGrid.Columns.Item(4).Width = 100
                oGrid.Columns.Item(5).Width = 300
                oGrid.Columns.Item(6).Width = 70

            Else

                Fecha1 = ArreglarFecha1(stFecha1)
                Fecha2 = ArreglarFecha2(stFecha2)
                Serie = stSerie

                oGrid = coForm.Items.Item("MyGrid").Specific
                coForm.DataSources.DataTables.Item(0).ExecuteQuery("Select '' as ""Check"",T1.""DocEntry"" as ""Ref"",T1.""DocNum"",TO_DATE(T1.""DocDate"") as ""Fecha"",T1.""CardCode"" as ""Cliente"",T2.""CardName"" as ""Nombre de Cliente"",T1.""DocTotal"" as ""Total"" from ""ORDR"" T1 inner join ""OCRD"" T2 on T2.""CardCode""=T1.""CardCode"" inner join ""RDR1"" T3 on T3.""DocEntry""=T1.""DocEntry"" inner join ""OITM"" T4 on T4.""ItemCode""=T3.""ItemCode"" inner join ""OITW"" T5 on T5.""WhsCode""=T3.""WhsCode"" and T5.""ItemCode""=T4.""ItemCode"" inner join ""NNM1"" T6 on T6.""Series""=T1.""Series"" where T1.""DocStatus""='O' and T1.""DocDate"" between '" & Fecha1 & "' and '" & Fecha2 & "' and T4.""validFor""='Y' and T5.""Locked""='N' and T1.""U_tekStatus""='-1' and T6.""Series""=" & Serie & " group by T1.""DocEntry"",T1.""DocNum"",T1.""DocDate"",T1.""CardCode"",T2.""CardName"",T1.""DocTotal"" order by T1.""DocDate"",T1.""DocNum""")
                oGrid.DataTable = coForm.DataSources.DataTables.Item("MyDataTable")

                ocultarColumnas("Ref")

                oColumn = oGrid.Columns.Item("Ref")
                oColumn.LinkedObjectType = 17

                oGrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox

                ' Set columns size
                oGrid.Columns.Item(0).Width = 40
                oGrid.Columns.Item(1).Width = 70
                oGrid.Columns.Item(2).Width = 50
                oGrid.Columns.Item(3).Width = 70
                oGrid.Columns.Item(4).Width = 100
                oGrid.Columns.Item(5).Width = 300
                oGrid.Columns.Item(6).Width = 70

            End If

        Catch ex As Exception
            cSBOApplication.MessageBox("DatosGrid: " & ex.Message)
        End Try

    End Function

    Public Function cerrarOrdenesVentas()
        Dim oGrid As SAPbouiCOM.Grid
        Dim oDataTable As SAPbouiCOM.DataTable
        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        Dim docNum As String
        Dim docEntry As String

        Try
            coForm = cSBOApplication.Forms.Item("frmCOVSA2")
            oGrid = coForm.Items.Item("MyGrid").Specific
            oDataTable = oGrid.DataTable
            oRecSetH = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            For i = 0 To oDataTable.Rows.Count - 1
                If oDataTable.GetValue("Check", i) = "Y" Then
                    docNum = oDataTable.GetValue("DocNum", i)
                    stQueryH = "Select T1.""DocEntry"" from ""ORDR"" T1 where T1.""DocNum""=" & docNum
                    oRecSetH.DoQuery(stQueryH)

                    If oRecSetH.RecordCount > 0 Then
                        oRecSetH.MoveFirst()
                        docEntry = oRecSetH.Fields.Item("DocEntry").Value

                        Dim RetVal As Long
                        Dim oOrder As SAPbobsCOM.Documents
                        oOrder = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)

                        'Retrieve the document record to close from the database
                        RetVal = oOrder.GetByKey(docEntry)

                        'Close the record
                        RetVal = oOrder.Close()


                    End If

                End If

            Next

            DatosGrid()

        Catch ex As Exception
            cSBOApplication.MessageBox("cerrarOrdenesVentas: " & ex.Message)
            Return -1
        End Try
    End Function

    '---- funcion que oculta columnas dependiendo de vista
    Private Function ocultarColumnas(ByVal stListaCols As String) As Integer
        Dim oGrid As SAPbouiCOM.Grid
        Dim listaCols As String()
        Dim i As Integer

        Try

            coForm = cSBOApplication.Forms.Item("frmCOVSA2")
            oGrid = coForm.Items.Item("MyGrid").Specific

            '--- separa columnas
            listaCols = Split(stListaCols, ",")

            '--- recorremos lista y oculta columnas
            For i = 0 To listaCols.Length - 1
                oGrid.Columns.Item(listaCols(i)).Editable = False
            Next

        Catch ex As Exception
            ocultarColumnas = -1
            cSBOApplication.MessageBox("OcultarColumnas. " & ex.Message)
        Finally
            oGrid = Nothing
        End Try
    End Function

    Public Function ArreglarFecha1(ByVal stFecha1 As String) As String

        Try
            Dim oRecSetF1 As SAPbobsCOM.Recordset
            Dim stQueryF1 As String

            oRecSetF1 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            stQueryF1 = "select substring('" & stFecha1 & "',7,4)as ""año1"",substring('" & stFecha1 & "',4,2) as ""mes1"",substring('" & stFecha1 & "',1,2)as ""dia1"" from dummy;"
            oRecSetF1.DoQuery(stQueryF1)

            If oRecSetF1.RecordCount > 0 Then
                oRecSetF1.MoveFirst()
                Fecha1 = oRecSetF1.Fields.Item("año1").Value & "-" & oRecSetF1.Fields.Item("mes1").Value & "-" & oRecSetF1.Fields.Item("dia1").Value
            End If

            'MsgBox(Fecha1)
            Return Fecha1

        Catch ex As Exception
            cSBOApplication.MessageBox("ArreglasFecha1. " & ex.Message)
        End Try
    End Function

    Public Function ArreglarFecha2(ByVal stFecha2 As String)

        Try
            Dim stQueryF2 As String
            Dim oRecSetF2 As SAPbobsCOM.Recordset

            oRecSetF2 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            stQueryF2 = "select substring('" & stFecha2 & "',7,4)as ""año2"",substring('" & stFecha2 & "',4,2) as ""mes2"",substring('" & stFecha2 & "',1,2)as ""dia2"" from dummy;"
            oRecSetF2.DoQuery(stQueryF2)

            If oRecSetF2.RecordCount > 0 Then
                oRecSetF2.MoveFirst()
                Fecha2 = oRecSetF2.Fields.Item("año2").Value & "-" & oRecSetF2.Fields.Item("mes2").Value & "-" & oRecSetF2.Fields.Item("dia2").Value
            End If

            'MsgBox(Fecha2)
            Return Fecha2

        Catch ex As Exception
            cSBOApplication.MessageBox("ArreglasFecha2. " & ex.Message)
        End Try
    End Function

End Class
