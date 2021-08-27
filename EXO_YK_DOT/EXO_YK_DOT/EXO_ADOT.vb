Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_ADOT
    Inherits EXO_Generales.EXO_DLLBase
#Region "Variables públicas"
    Public Shared _sDocEntry As String = ""
    Public Shared _sDocNum As String = ""
    Public Shared _sObjType As String = ""
    Public Shared _dtArt As System.Data.DataTable
#End Region
    Public Sub New(ByRef general As EXO_Generales.EXO_General, actualizar As Boolean)
        MyBase.New(general, actualizar)

        If actualizar Then
            cargaCampos()
        End If
    End Sub
    Public Sub cargaCampos()
        If objGlobal.conexionSAP.esAdministrador() Then
            objGlobal.conexionSAP.escribeMensaje("El usuario es administrador")
            'Definicion descuentos financieros
            Dim contenidoXML As String

            Try
                contenidoXML = objGlobal.Functions.leerEmbebido(Me.GetType(), "UDO_EXO_ADOT.xml")
                objGlobal.conexionSAP.refCompañia.LoadBDFromXML(contenidoXML)
                objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - Validado UDO_EXO_ADOT", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            Catch exCOM As System.Runtime.InteropServices.COMException
                objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Catch ex As Exception
                objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Finally

            End Try
        Else
            objGlobal.conexionSAP.escribeMensaje("(EXO) - El usuario NO es administrador")
        End If
    End Sub
    Public Overrides Function filtros() As EventFilters
        Dim filtrosXML As Xml.XmlDocument = New Xml.XmlDocument
        filtrosXML.LoadXml(objGlobal.Functions.leerEmbebido(Me.GetType(), "XML_FILTROS.xml"))
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(filtrosXML.OuterXml)

        Return filtro
    End Function
    Public Overrides Function menus() As XmlDocument
        Return Nothing
    End Function
    Public Overrides Function SBOApp_ItemEvent(ByRef infoEvento As EXO_Generales.EXO_infoItemEvent) As Boolean
        Dim res As Boolean = True
        Dim oForm As SAPbouiCOM.Form = SboApp.Forms.Item(infoEvento.FormUID)

        Dim Path As String = ""
        Dim XmlDoc As New System.Xml.XmlDocument
        Dim EXO_Functions As New EXO_BasicDLL.EXO_Generic_Forms_Functions(Me.objGlobal.conexionSAP)

        Try
            If infoEvento.FormTypeEx = "UDO_FT_EXO_ADOT" Then
                If infoEvento.InnerEvent = True Then
                    Select Case infoEvento.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                        Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                            If infoEvento.BeforeAction = False Then
                                If EventHandler_Form_Visible(infoEvento) = False Then
                                    GC.Collect()
                                    Return False
                                End If
                            End If
                        Case BoEventTypes.et_LOST_FOCUS
                            If infoEvento.BeforeAction = False Then
                                If EventHandler_LOST_FOCUS(infoEvento) = False Then
                                    GC.Collect()
                                    Return False
                                End If
                                'If infoEvento.ItemUID = "0_U_G" AndAlso infoEvento.ColUID = "C_0_1" Then
                                '    If CargarComboDOT(objGlobal, oForm, infoEvento.Row) = False Then
                                '        Exit Function
                                '    End If
                                'End If
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                            If infoEvento.BeforeAction = False Then
                                If EventHandler_VALIDATE_after(infoEvento) = False Then
                                    GC.Collect()
                                    Return False
                                End If
                            End If
                    End Select
                Else
                    Select Case infoEvento.EventType
                        Case BoEventTypes.et_COMBO_SELECT
                            If infoEvento.BeforeAction = False And infoEvento.ActionSuccess = True Then
                                'If infoEvento.ItemUID = "0_U_G" AndAlso infoEvento.ColUID = "C_0_1" Then
                                '    If CargarComboDOT(objGlobal, oForm, infoEvento.Row) = False Then
                                '        Exit Function
                                '    End If
                                'ElseIf infoEvento.ItemUID = "0_U_G" AndAlso infoEvento.ColUID = "C_0_2" Then
                                '    If ControlarDOT(oForm, infoEvento.Row) = False Then
                                '        Exit Function
                                '    End If
                                'End If
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                            If infoEvento.BeforeAction = False Then

                            End If
                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                            If infoEvento.BeforeAction = False Then

                            Else
                                'If EventHandler_Click_Before(objGlobal, infoEvento) = False And infoEvento.ActionSuccess = True Then
                                '    Return False
                                'End If
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                            If infoEvento.BeforeAction = False Then
                                If EventHandler_ItemPressed_After(objGlobal, infoEvento) = False Then
                                    Return False
                                End If
                            Else
                                If EventHandler_ItemPressed_Before(objGlobal, infoEvento) = False Then
                                    Return False
                                End If
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_CLICK

                    End Select
                End If
            End If


        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Return False
        Finally

            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try

        Return res
    End Function
    Private Function EventHandler_VALIDATE_after(ByRef pVal As EXO_Generales.EXO_infoItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oConds As SAPbouiCOM.Conditions = Nothing
        Dim oCond As SAPbouiCOM.Condition = Nothing

        EventHandler_VALIDATE_after = False

        Try
            If pVal.ActionSuccess = True Then
                'Recuperar el formulario
                oForm = SboApp.Forms.Item(pVal.FormUID)

                If pVal.ItemUID = "0_U_G" And pVal.ColUID = "C_0_3" And pVal.ItemChanged Then
                    Dim iRegistros As Integer = oForm.DataSources.DBDataSources.Item("@EXO_ADOTL").Size
                    Dim iRegActivo As Integer = oForm.DataSources.DBDataSources.Item("@EXO_ADOTL").Offset + 1
                    If iRegistros = iRegActivo Then
                        CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).AddRow()
                        CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).ClearRowData(iRegActivo + 1)
                        CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).FlushToDataSource()
                    End If
                End If
            End If

            EventHandler_VALIDATE_after = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oConds, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCond, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function EventHandler_LOST_FOCUS(ByRef pVal As EXO_Generales.EXO_infoItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sItemCode As String = "" : Dim sDOT As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.conexionSAP.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        EventHandler_LOST_FOCUS = False

        Try

            oForm = SboApp.Forms.Item(pVal.FormUID)
            If pVal.ItemUID = "0_U_G" And pVal.ColUID = "C_0_2" Then
                sItemCode = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value.ToString
                sDOT = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_2").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value.ToString
                If sDOT.Trim.Length = 3 Then
                    sDOT = "0" & sDOT.Trim
                Else
                    sDOT = sDOT.Trim
                End If
                CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_2").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = sDOT
                If sItemCode.Trim <> "" And sDOT.Trim <> "" Then
                    'Buscamos la cantidad de DOT que tenemos en el momento
                    sSQL = " SELECT ""U_EXO_CANT"" FROM ""@EXO_DOTL"" WHERE ""Code""='" & sItemCode.Trim & "' and ""U_EXO_DOT""='" & sDOT.Trim & "' "
                    oRs.DoQuery(sSQL)
                    If oRs.RecordCount > 0 Then
                        Dim sCANT As String = oRs.Fields.Item("U_EXO_CANT").Value.ToString.Trim
                        CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_4").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = sCANT
                    Else
                        CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_4").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = 0
                    End If
                End If
            End If
            EventHandler_LOST_FOCUS = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Private Function EventHandler_Form_Visible(ByRef pVal As EXO_Generales.EXO_infoItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oItem As SAPbouiCOM.Item

        EventHandler_Form_Visible = False

        Try
            If pVal.ActionSuccess = True Then
                'Recuperar el formulario
                oForm = SboApp.Forms.Item(pVal.FormUID)

                If oForm.Visible = True Then
                    If oForm.Mode = BoFormMode.fm_ADD_MODE Then
                        oForm.DataSources.DBDataSources.Item("@EXO_ADOT").SetValue("U_EXO_DOCENTRY", 0, _sDocEntry)
                        oForm.DataSources.DBDataSources.Item("@EXO_ADOT").SetValue("U_EXO_DOCNUM", 0, _sDocNum)
                        oForm.DataSources.DBDataSources.Item("@EXO_ADOT").SetValue("U_EXO_OTYPE", 0, _sObjType)
                        Dim sStatus As String = oForm.DataSources.DBDataSources.Item("@EXO_ADOT").GetValue("Status", 0)
                        If sStatus = "C" Then
                            oForm.Mode = BoFormMode.fm_VIEW_MODE
                        End If
                    ElseIf oForm.Mode = BoFormMode.fm_FIND_MODE Then
                        oForm.DataSources.DBDataSources.Item("@EXO_ADOT").SetValue("U_EXO_DOCENTRY", 0, _sDocEntry)
                        oForm.DataSources.DBDataSources.Item("@EXO_ADOT").SetValue("U_EXO_DOCNUM", 0, _sDocNum)
                        oForm.DataSources.DBDataSources.Item("@EXO_ADOT").SetValue("U_EXO_OTYPE", 0, _sObjType)
                    End If

                    'Filtrar Choosfromlist Artículos
#Region "ChooseFromList Artículos"
                    Dim oConds As SAPbouiCOM.Conditions = Nothing
                    Dim oCond As SAPbouiCOM.Condition = Nothing
                    Dim sCode As String = ""
                    oConds = New SAPbouiCOM.Conditions
                    For Each row As DataRow In _dtArt.Rows
                        oCond = oConds.Add
                        oCond.Alias = "ItemCode"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        sCode = CStr(row("Articulo"))
                        oCond.CondVal = sCode
                        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    Next
                    If oConds.Count > 0 Then oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_NONE
                    oForm.ChooseFromLists.Item("CFL_0").SetConditions(oConds)
#End Region


                    oItem = oForm.Items.Item("btnCom")
                    oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                    oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                    oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                    oItem = oForm.Items.Item("btnConf")
                    oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                    oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                    oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                End If
            End If

            EventHandler_Form_Visible = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    'Private Function EventHandler_Click_Before(ByRef objGlobal As EXO_Generales.EXO_General, ByRef pVal As EXO_Generales.EXO_infoItemEvent) As Boolean
    '    Dim oForm As SAPbouiCOM.Form = Nothing
    '    EventHandler_Click_Before = False

    '    Try
    '        oForm = SboApp.Forms.Item(pVal.FormUID)
    '        If pVal.ColUID = "C_0_2" Then 'Columna 2             
    '            'Refrescamos el combo
    '            'Cargamos combo SPECI
    '            oForm.Freeze(True)

    '            If CargarComboDOT(objGlobal, oForm, pVal.Row) = False Then
    '                Exit Function
    '            End If

    '        End If

    '        EventHandler_Click_Before = True

    '    Catch exCOM As System.Runtime.InteropServices.COMException
    '        oForm.Freeze(False)
    '        Throw exCOM
    '    Catch ex As Exception
    '        oForm.Freeze(False)
    '        Throw ex
    '    Finally
    '        oForm.Freeze(False)
    '        EXO_CleanCOM.CLiberaCOM.Form(oForm)

    '    End Try
    'End Function
    Private Function EventHandler_ItemPressed_Before(ByRef objGlobal As EXO_Generales.EXO_General, ByRef pVal As EXO_Generales.EXO_infoItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        EventHandler_ItemPressed_Before = False
        Dim sObjType As String = ""
        Try
            oForm = SboApp.Forms.Item(pVal.FormUID)
            Select Case pVal.ItemUID
                Case "1"
                    'Limpiamos el registro vacío
                    If oForm.Mode = BoFormMode.fm_ADD_MODE Or oForm.Mode = BoFormMode.fm_UPDATE_MODE Then
                        Dim iRegistros As Integer = oForm.DataSources.DBDataSources.Item("@EXO_ADOTL").Size
                        Dim sDOT As String = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_2").Cells.Item(iRegistros).Specific, SAPbouiCOM.EditText).Value
                        If sDOT.Trim = "" Then
                            CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).DeleteRow(iRegistros)
                        End If
                    End If
            End Select

            EventHandler_ItemPressed_Before = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByRef objGlobal As EXO_Generales.EXO_General, ByRef pVal As EXO_Generales.EXO_infoItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        EventHandler_ItemPressed_After = False
        Dim sObjType As String = ""
        Try
            oForm = SboApp.Forms.Item(pVal.FormUID)
            sObjType = oForm.DataSources.DBDataSources.Item("@EXO_ADOT").GetValue("U_EXO_OTYPE", 0)
            Select Case pVal.ItemUID
                Case "btnCom"  'Botón comprobar
                    'sumará las cantidades agrupadas por articulo del documento y 
                    'realizará lo mismo con la tabla de DOTs, si falta información mostrará mensaje de aviso
                    If oForm.Mode = BoFormMode.fm_OK_MODE Then
                        If comprobar(oForm) = True Then
                            objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - Se ha comprobado los datos. Las cantidades son correctas.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            objGlobal.conexionSAP.SBOApp.MessageBox("Se ha comprobado los datos. Las cantidades son correctas.")
                        End If
                    Else
                        objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - Para realizar la comprobación, debe grabar primero.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        objGlobal.conexionSAP.SBOApp.MessageBox("Para realizar la comprobación, debe grabar primero.")
                    End If
                Case "btnConf" 'Botón confirmar
                    'Realizará primero la comprobación y si es correcta realizará el 
                    'proceso de sumar los dots en el maestro de articulo y se marcará el documento 
                    'como cerrado para que no se pueda volver a procesar
                    If objGlobal.conexionSAP.SBOApp.MessageBox("¿Desea confirmar los DOTs asignados a este documento?", 1, "Sí", "No") = 1 Then
                        Dim sDocEntryDOT As String = oForm.DataSources.DBDataSources.Item("@EXO_ADOT").GetValue("DocEntry", 0)
                        If comprobar(oForm) = True Then
                            objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - Se ha comprobado los datos. Las cantidades son correctas.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            objGlobal.conexionSAP.SBOApp.MessageBox("Se ha comprobado los datos. Las cantidades son correctas.")
#Region "Recorremos las líneas y sumamos a UDO DOT"
                            oRs = CType(objGlobal.conexionSAP.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                            sSQL = "SELECT ""U_EXO_ART"",""U_EXO_DOT"",SUM(""U_EXO_CANT"") ""CANT"" FROM ""@EXO_ADOTL"" "
                            sSQL &= " WHERE DocEntry=" & sDocEntryDOT
                            sSQL &= " GROUP BY ""U_EXO_ART"", ""U_EXO_DOT"" "
                            oRs.DoQuery(sSQL)
                            If oRs.RecordCount > 0 Then
                                For i = 0 To oRs.RecordCount - 1
                                    Select Case sObjType
                                        Case "20", "59"
                                            sSQL = "UPDATE ""@EXO_DOTL"" SET ""U_EXO_CANT""=""U_EXO_CANT"" + " & oRs.Fields.Item("CANT").Value.ToString
                                            sSQL &= " WHERE ""Code""='" & oRs.Fields.Item("U_EXO_ART").Value.ToString & "' "
                                            sSQL &= " AND ""U_EXO_DOT""='" & oRs.Fields.Item("U_EXO_DOT").Value.ToString & "' "
                                        Case "21", "60"
                                            sSQL = "UPDATE ""@EXO_DOTL"" SET ""U_EXO_CANT""=""U_EXO_CANT"" - " & oRs.Fields.Item("CANT").Value.ToString
                                            sSQL &= " WHERE ""Code""='" & oRs.Fields.Item("U_EXO_ART").Value.ToString & "' "
                                            sSQL &= " AND ""U_EXO_DOT""='" & oRs.Fields.Item("U_EXO_DOT").Value.ToString & "' "
                                    End Select

                                    objGlobal.SQL.sqlUpdB1(sSQL)
                                    oRs.MoveNext()
                                Next
                            End If
#End Region
                            'Cerramos el documento
                            Dim oDI_COM As EXO_DIAPI.EXO_UDOEntity = Nothing 'Instancia del UDO para Insertar datos
                            oDI_COM = New EXO_DIAPI.EXO_UDOEntity(objGlobal.conexionSAP.refCompañia, "EXO_ADOT") 'UDO de Campos de SAP
                            If oDI_COM.GetByKey(sDocEntryDOT) = True Then
                                If oDI_COM.UDO_Close(sDocEntryDOT) = True Then
                                    objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - Se cierra el documento de asignación de DOT", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    'actualizamos la pantalla.
                                    objGlobal.conexionSAP.SBOApp.ActivateMenuItem("1304")
                                    oForm.Mode = BoFormMode.fm_VIEW_MODE
                                End If
                            End If
                        End If
                    End If
                Case "1"
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        objGlobal.conexionSAP.SBOApp.ActivateMenuItem("1289")
                    End If
            End Select

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Public Overrides Function SBOApp_FormDataEvent(ByRef infoEvento As EXO_Generales.EXO_BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Try
            oForm = SboApp.Forms.Item(infoEvento.FormUID)
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_ADOT"
                        Select Case infoEvento.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                Dim sStatus As String = oForm.DataSources.DBDataSources.Item("@EXO_ADOT").GetValue("Status", 0)
                                If sStatus = "C" Then
                                    oForm.Mode = BoFormMode.fm_VIEW_MODE
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                        End Select

                End Select

            Else
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_ADOT"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                        End Select
                End Select
            End If

            Return MyBase.SBOApp_FormDataEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
#Region "Métodos auxiliares"
    'Public Shared Function CargarComboDOT(ByRef oobjglobal As EXO_Generales.EXO_General, ByRef oForm As SAPbouiCOM.Form, ByVal iRow As Integer) As Boolean
    '    Dim sSQL As String = ""
    '    Dim sArt As String = ""
    '    CargarComboDOT = False

    '    Try
    '        If oForm.Visible = True Then
    '            If CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(iRow).Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
    '                sArt = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(iRow).Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
    '            Else
    '                sArt = ""
    '            End If
    '            sSQL = "Select '-' ""DOT"", 'Definir DOT Nuevo'  "
    '            sSQL &= " UNION ALL "
    '            sSQL &= "Select ""U_EXO_DOT"" ""DOT"", CAST(""U_EXO_CANT"" as VARCHAR) FROM ""@EXO_DOTL"" WHERE ""Code""='" & sArt & "' ORDER BY  ""DOT"" "

    '            oobjglobal.conexionSAP.refSBOApp.cargaCombo(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_2").ValidValues, sSQL)
    '        End If


    '        CargarComboDOT = True

    '    Catch ex As Exception
    '        Throw ex
    '    Finally

    '    End Try
    'End Function
    'Private Function ControlarDOT(ByRef oForm As SAPbouiCOM.Form, ByVal iRow As Integer) As Boolean
    '    Dim sSQL As String = ""
    '    Dim sArt As String = "" : Dim sDes As String = ""
    '    Dim sDOT As String = ""
    '    ControlarDOT = False

    '    Try
    '        If oForm.Visible = True Then
    '            If CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(iRow).Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
    '                sArt = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(iRow).Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
    '                If CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(iRow).Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
    '                    sDOT = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_2").Cells.Item(iRow).Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
    '                End If
    '            End If
    '            If sArt <> "" And sDOT <> "" Then
    '                If sDOT = "-" Then
    '                    'CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(iRow).Specific, SAPbouiCOM.ComboBox).Active = True
    '                    'CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(iRow).Specific, SAPbouiCOM.ComboBox).Item.Click()
    '                    'Sacamos la pantalla para que cree un DOT nuevo
    '                    CargarUDODOT(sArt, sDes)
    '                End If
    '            End If
    '        End If

    '        ControlarDOT = True

    '    Catch ex As Exception
    '        Throw ex
    '    Finally

    '    End Try
    'End Function
    'Private Function CargarUDODOT(ByVal sArticulo As String, ByVal sDescripcion As String) As Boolean
    '    Dim sSQL As String = ""
    '    Dim oRs As SAPbobsCOM.Recordset = Nothing

    '    CargarUDODOT = False

    '    Try
    '        EXO_DOT._sArticulo = sArticulo
    '        EXO_DOT._sDescripcion = sDescripcion

    '        oRs = CType(objGlobal.conexionSAP.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
    '        sSQL = " SELECT ""Code"" FROM ""@EXO_DOT"" WHERE ""Code""='" & sArticulo & "' "
    '        oRs.DoQuery(sSQL)
    '        If oRs.RecordCount > 0 Then
    '            objGlobal.conexionSAP.cargaFormUdoBD_Clave("EXO_DOT", sArticulo)
    '            'objGlobal.conexionSAP.SBOApp.OpenForm(BoFormObjectEnum.fo_UserDefinedObject, "EXO_DOT”, sArticulo.Trim)
    '        Else
    '            objGlobal.conexionSAP.cargaFormUdoBD("EXO_DOT")
    '        End If

    '        CargarUDODOT = True
    '    Catch exCOM As System.Runtime.InteropServices.COMException
    '        Throw exCOM
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '        EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
    '    End Try
    'End Function

    Private Function Comprobar(ByRef oform As SAPbouiCOM.Form) As Boolean
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.conexionSAP.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsDOT As SAPbobsCOM.Recordset = CType(objGlobal.conexionSAP.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim sMensaje As String = ""
        Dim sDocEntry As String = "" : Dim sDot As String = "" : Dim sArticulo As String = ""
        Comprobar = False
        Try
#Region "Comprobar DOT"
            'Comprobaremos todos los DOT a ver si existen
            sDocEntry = oform.DataSources.DBDataSources.Item("@EXO_ADOT").GetValue("DocEntry", 0)
            sSQL = " SELECT ""U_EXO_ART"", ""U_EXO_DOT""  "
            sSQL &= " FROM ""@EXO_ADOTL""  "
            sSQL &= " WHERE ""DocEntry""=" & sDocEntry
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                For i = 0 To oRs.RecordCount - 1
                    sSQL = "Select ""U_EXO_DOT"" From ""@EXO_DOTL"" "
                    sSQL &= " WHERE ""Code""='" & oRs.Fields.Item("U_EXO_ART").Value.ToString & "' "
                    sSQL &= " and ""U_EXO_DOT""='" & oRs.Fields.Item("U_EXO_DOT").Value.ToString & "' "
                    sDot = objGlobal.conexionSAP.refCompañia.SQL.sqlStringB1(sSQL)
                    If sDot = "" Then
                        'No existe, por lo que se advierte si se crea el DOT
                        sArticulo = oRs.Fields.Item("U_EXO_ART").Value.ToString
                        sDot = oRs.Fields.Item("U_EXO_DOT").Value.ToString
                        'If SboApp.MessageBox("DOT: " & sDot & " no existe. ¿Quiere crearlo?", 1, "Sí", "No") = 1 Then
                        '    'Creamos el dot

                        CrearDOT(sArticulo, sDot)
                            'Else
                            '    sMensaje = "Si no se crea el DOT inexistente no s epuede continuar."
                            '    objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            '    Exit Function
                            'End If
                        End If
                    oRs.MoveNext()
                Next
            End If
#End Region
#Region "Comprobar asignación"
            'Recorremos los artículos del documento y vemos si están asignados
            For Each row As DataRow In _dtArt.Rows
                Dim dCantDoc As Double = 0 : dCantDoc = CType(row("Cantidad").ToString, Double)
                sArticulo = row("Articulo").ToString
                sDocEntry = oform.DataSources.DBDataSources.Item("@EXO_ADOT").GetValue("DocEntry", 0)
                Dim sObjType As String = oform.DataSources.DBDataSources.Item("@EXO_ADOT").GetValue("U_EXO_OTYPE", 0)
                sSQL = " SELECT ADOT.""U_EXO_ART"", SUM(ADOT.""U_EXO_CANT"") ""Cantidad"", SUM(DOT.""U_EXO_CANT"") ""CantT"" "
                sSQL &= " FROM ""@EXO_ADOTL"" ADOT  "
                sSQL &= " INNER JOIN ""@EXO_DOTL"" DOT ON DOT.""Code""=ADOT.""U_EXO_ART"" and DOT.""U_EXO_DOT""=ADOT.""U_EXO_DOT"" "
                sSQL &= " WHERE ADOT.""DocEntry""=" & sDocEntry & " And ADOT.""U_EXO_ART""='" & sArticulo & "' "
                sSQL &= " GROUP BY ADOT.""U_EXO_ART"" "
                oRs.DoQuery(sSQL)
                If oRs.RecordCount > 0 Then
                    If oRs.Fields.Item("Cantidad").Value > dCantDoc Then
                        sMensaje = "El artículo " & oRs.Fields.Item("U_EXO_ART").Value.ToString & " tiene asignado en el documento "
                        sMensaje &= " la cantidad de " & oRs.Fields.Item("Cantidad").Value.ToString & " y como max. sólo puede asignarse " & dCantDoc
                        objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        objGlobal.conexionSAP.SBOApp.MessageBox(sMensaje)
                        Exit Function
                    ElseIf oRs.Fields.Item("Cantidad").Value < dCantDoc Then
                        sMensaje = "El artículo " & oRs.Fields.Item("U_EXO_ART").Value.ToString & " tiene asignado en el documento "
                        sMensaje &= " la cantidad de " & oRs.Fields.Item("Cantidad").Value.ToString & " y hay que asignar " & dCantDoc
                        objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        objGlobal.conexionSAP.SBOApp.MessageBox(sMensaje)
                        Exit Function
                    Else
                        Select Case sObjType
                            Case "20"
                            Case "21"
                                sSQL = " Select ADOT.""U_EXO_ART"", ADOT.""U_EXO_DOT"", SUM (ADOT.""U_EXO_CANT"") ""Cantidad"", SUM (DOT.""U_EXO_CANT"") ""CantDOT"" "
                                sSQL &= " FROM ""@EXO_ADOTL"" ADOT  "
                                sSQL &= " INNER JOIN ""@EXO_DOTL"" DOT On DOT.""Code""=ADOT.""U_EXO_ART"" And DOT.""U_EXO_DOT""=ADOT.""U_EXO_DOT"" "
                                sSQL &= " WHERE ADOT.DocEntry=" & sDocEntry & " And ADOT.""U_EXO_ART""='" & sArticulo & "' "
                                sSQL &= " GROUP BY ADOT.""U_EXO_ART"",ADOT.""U_EXO_DOT"" "
                                oRsDOT.DoQuery(sSQL)

                                For i = 0 To oRsDOT.RecordCount - 1
                                    If oRsDOT.Fields.Item("Cantidad").Value > oRsDOT.Fields.Item("CantDOT").Value Then
                                        sMensaje = "El artículo " & oRsDOT.Fields.Item("U_EXO_ART").Value.ToString & " tiene asignado en el documento  con el DOT " & oRsDOT.Fields.Item("U_EXO_DOT").Value.ToString
                                        sMensaje &= " la cantidad de " & oRsDOT.Fields.Item("Cantidad").Value.ToString & " y como max. sólo puede asignarse " & oRsDOT.Fields.Item("CantDOT").Value.ToString
                                        objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        objGlobal.conexionSAP.SBOApp.MessageBox(sMensaje)
                                        Exit Function
                                    End If
                                    oRs.MoveNext()
                                Next
                        End Select

                    End If
                Else
                    sMensaje = "El artículo " & sArticulo & " no tiene asignado en el documento  ningún DOT. "
                    sMensaje &= "Por favor, asigne la cantidad de " & dCantDoc.ToString & " para validar el documento."
                    objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objGlobal.conexionSAP.SBOApp.MessageBox(sMensaje)
                    Exit Function
                End If
            Next
#End Region

            Comprobar = True
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsDOT, Object))
        End Try

    End Function
    Private Sub CrearDOT(ByVal sArticulo As String, ByVal sDot As String)
        Dim sSQL As String = "" : Dim sExiste As String = "" : Dim sMensaje As String = ""
        Dim oDI_COM As EXO_DIAPI.EXO_UDOEntity = Nothing 'Instancia del UDO para Insertar datos
        Try
            oDI_COM = New EXO_DIAPI.EXO_UDOEntity(objGlobal.conexionSAP.refCompañia, "EXO_DOT") 'UDO
#Region "Comprobar año y semana"
            Dim iAnno As Integer = Right(sDot, 2) : Dim iAnnoAct As Integer = Right(Now.Year.ToString("0000"), 2)
            Dim iSemana As Integer = Left(sDot, 2)
            If IsNumeric(iAnno) Then
                If iAnno >= iAnnoAct Then
                    sMensaje = "El Año " & iAnno.ToString & " debe ser inferior al año actual. "
                    objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    objGlobal.conexionSAP.SBOApp.MessageBox(sMensaje)
                    Exit Sub
                End If
            End If
            If IsNumeric(iSemana) Then
                If iSemana > 52 Then
                    sMensaje = "La semana " & iSemana.ToString & " debe ser inferior a 52. "
                    objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    objGlobal.conexionSAP.SBOApp.MessageBox(sMensaje)
                    Exit Sub
                End If
            End If
#End Region
            'Comprobamos que exista para crear o actualizar
            sSQL = "Select ""Code"" From ""@EXO_DOT"" "
            sSQL &= " WHERE ""Code""='" & sArticulo & "' "
            sExiste = objGlobal.conexionSAP.refCompañia.SQL.sqlStringB1(sSQL)
            If sExiste = "" Then
#Region "crear"
                oDI_COM.GetNew()
                oDI_COM.SetValue("Code") = sArticulo
                sSQL = "Select ""ItemName"" From ""OITM"" "
                sSQL &= " WHERE ""ItemCode""='" & sArticulo & "' "
                oDI_COM.SetValue("Name") = objGlobal.conexionSAP.refCompañia.SQL.sqlStringB1(sSQL)
                CrearCamposLíneas(oDI_COM, sDot)
                If oDI_COM.UDO_Add = False Then
                    Throw New Exception("(EXO) - Error al añadir DOT " & oDI_COM.GetLastError)
                End If
#End Region
            Else
#Region "Actualizar"
                oDI_COM.GetByKey(sExiste)
                CrearCamposLíneas(oDI_COM, sDot)
                If oDI_COM.UDO_Update = False Then
                    Throw New Exception("(EXO) - Error al añadir DOT " & oDI_COM.GetLastError)
                End If
#End Region
            End If
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oDI_COM, Object))
        End Try
    End Sub
    Private Sub CrearCamposLíneas(ByRef oDI_COM As EXO_DIAPI.EXO_UDOEntity, ByVal sDOT As String)
        Try
            oDI_COM.GetNewChild("EXO_DOTL")
            oDI_COM.SetValueChild("U_EXO_DOT") = sDOT
            oDI_COM.SetValueChild("U_EXO_ANNO") = Right(sDOT, 2)
            oDI_COM.SetValueChild("U_EXO_SEMANA") = Left(sDOT, 2)
            oDI_COM.SetValueChild("U_EXO_CANT") = 0
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
End Class
