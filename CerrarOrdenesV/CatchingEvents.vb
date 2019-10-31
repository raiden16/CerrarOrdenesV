Option Strict Off
Option Explicit On
Imports System.Xml
Imports System.IO
Imports System.Windows.Forms

Friend Class CatchingEvents

    Public WithEvents SBOApplication As SAPbouiCOM.Application
    Public SBOCompany As SAPbobsCOM.Company
    Friend csDirectory As String
    Private oForm As SAPbouiCOM.Form
    Private oGrid As SAPbouiCOM.Grid

    Public Sub New()
        MyBase.New()
        SetAplication()
        SetConnectionContext()
        ConnectSBOCompany()

        'Agregar Menu
        addMenuItems()

        setFilters()

    End Sub

    Private Sub SetAplication()
        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String
        Try

            SboGuiApi = New SAPbouiCOM.SboGuiApi
            sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
            SboGuiApi.Connect(sConnectionString)
            SBOApplication = SboGuiApi.GetApplication()

        Catch ex As Exception
            MsgBox(ex.Message)
            SBOApplication.MessageBox("Falló la conexión con la aplicación SBO")
            System.Windows.Forms.Application.Exit()
            End
        End Try
    End Sub

    Private Sub SetConnectionContext()
        Try
            SBOCompany = SBOApplication.Company.GetDICompany
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con el DI")
            SBOApplication = Nothing
            System.Windows.Forms.Application.Exit()
            End
        End Try
    End Sub

    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '// CONEXION CON LA BASE DE DATOS
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub ConnectSBOCompany()
        Dim loRecSet As SAPbobsCOM.Recordset
        Try
            '/////// Oobtiene la direccion de la aplicacion de acuerdo a registro del sistema
            csDirectory = ReadPath("Cierre de Ordenes de Ventas")
            If (csDirectory = "") Then
                System.Windows.Forms.Application.Exit()
                End
            End If
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con la BD. " & ex.Message)
            SBOApplication = Nothing
            System.Windows.Forms.Application.Exit()
            End '//termina aplicación
        Finally
            loRecSet = Nothing
        End Try
    End Sub

    Private Sub addMenuItems()
        Dim loForm As SAPbouiCOM.Form = Nothing
        Dim loMenus As SAPbouiCOM.Menus
        Dim loMenusRoot As SAPbouiCOM.Menus
        Dim loMenuItem As SAPbouiCOM.MenuItem

        Try
            '////// Obtiene referencia de la forma Principal de Modulos
            loForm = SBOApplication.Forms.GetForm(169, 1)

            loForm.Freeze(True)

            '////// Obtiene la referencia de los Menus de SBO
            loMenus = SBOApplication.Menus.Item(6).SubMenus

            '////// Adiciona un Nuevo Menu para la Aplicacion de VectorSBO
            If loMenus.Exists("COV01") Then
                loMenus.RemoveEx("COV01")
            End If

            loMenuItem = loMenus.Add("COV01", "Cierre de Ordenes de Ventas", SAPbouiCOM.BoMenuType.mt_POPUP, loMenus.Count)
            'MsgBox(csDirectory & "\" & "checkb.png")
            'loMenuItem.Image = csDirectory & "\" & "check2.png"

            loMenuItem.Image = Application.StartupPath & "\" & "check2.png"

            loMenusRoot = loMenuItem.SubMenus

            '////// Adiciona un menu Item
            If loMenusRoot.Exists("COV11") Then
                loMenusRoot.RemoveEx("COV11")
            End If
            loMenuItem = loMenusRoot.Add("COV11", "Cerrar Ordenes de Ventas", SAPbouiCOM.BoMenuType.mt_STRING, loMenusRoot.Count)
            loMenus = loMenuItem.SubMenus

            loForm.Freeze(False)
            loForm.Update()

        Catch ex As Exception
            If (Not loForm Is Nothing) Then
                loForm.Freeze(False)
                loForm.Update()
            End If
            SBOApplication.MessageBox("CatchingEvents. Error al agregar las opciones del menú. " & ex.Message)
            End
        Finally
            loMenus = Nothing
            loMenusRoot = Nothing
            loMenuItem = Nothing
        End Try
    End Sub

    Private Sub setFilters()
        Dim lofilter As SAPbouiCOM.EventFilter
        Dim lofilters As SAPbouiCOM.EventFilters

        Try
            lofilters = New SAPbouiCOM.EventFilters

            lofilter = lofilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
            lofilter.AddEx("60004")
            lofilter.AddEx("133")

            lofilter = lofilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)

            SBOApplication.SetFilter(lofilters)
        Catch ex As Exception
            SBOApplication.MessageBox("SetFilter: " & ex.Message)
        End Try

    End Sub


    ''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    ''// METODOS PARA EVENTOS DE LA APLICACION
    ''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBOApplication.AppEvent
        Select Case EventType
            Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                System.Windows.Forms.Application.Exit()
                End
        End Select

    End Sub

    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '// CONTROLADOR DE EVENTOS MENU
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub SBOApplication_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBOApplication.MenuEvent
        Dim oCerrarOV As CerrarOV

        Try
            '//ANTES DE PROCESAR SBO
            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    '//////////////////////////////////SubMenu de Crear traslado inventario////////////////////////
                    Case "COV11"

                        'Crea Forma                       
                        oCerrarOV = New CerrarOV
                        oCerrarOV.CreateForm()

                        oCerrarOV = New CerrarOV
                        oCerrarOV.openForm(csDirectory)

                End Select
            End If

        Catch ex As Exception
            SBOApplication.MessageBox("clsCatchingEvents. MenuEvent" & ex.Message)
        Finally
            'oReservaPedido = Nothing
        End Try
    End Sub


    Private Sub SBOApplication_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBOApplication.AppEvent
        Select Case EventType
            Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                System.Windows.Forms.Application.Exit()
                End
        End Select
    End Sub


    Private Sub SBOApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBOApplication.ItemEvent
        Try
            If pVal.Before_Action = False And pVal.FormTypeEx <> "" Then

                Select Case pVal.FormTypeEx
                    '////////////////FORMA PARA ACTIVAR LICENCIA
                    Case "60004"
                        FrmPagoSBOControllerAfter(FormUID, pVal)
                End Select
            End If

        Catch ex As Exception
            SBOApplication.MessageBox("clsCatchingEvents. MenuEvent" & ex.Message)
        Finally
        End Try
    End Sub


    Private Sub FrmPagoSBOControllerAfter(ByVal FormUID As String, ByVal pVal As SAPbouiCOM.ItemEvent)

        Dim oFrmCerrarOrdenVentas As CerrarOV

        Try

            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    Select Case pVal.ItemUID

                        Case "btBuscar"

                            oFrmCerrarOrdenVentas = New CerrarOV
                            oFrmCerrarOrdenVentas.DatosGrid()

                        Case "btCerrar"

                            oForm = SBOApplication.Forms.Item("frmCOVSA2")
                            oGrid = oForm.Items.Item("MyGrid").Specific
                            Dim oDataTable As SAPbouiCOM.DataTable
                            oDataTable = oGrid.DataTable

                            If (oDataTable.Rows.Count = 0) Then

                                SBOApplication.MessageBox("Selecciona alguna Orden de Venta.")

                            Else

                                oFrmCerrarOrdenVentas = New CerrarOV
                                oFrmCerrarOrdenVentas.cerrarOrdenesVentas()

                            End If

                    End Select

            End Select

        Catch ex As Exception
            SBOApplication.MessageBox("clsCatchingEvents. Error en forma de Panel General. " & ex.Message)
        Finally

        End Try
    End Sub

End Class
