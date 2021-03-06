Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_DOT
    Inherits EXO_Generales.EXO_DLLBase
#Region "Variables públicas"
    Public Shared _sArticulo As String = ""
    Public Shared _sDescripcion As String = ""
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
                contenidoXML = objGlobal.Functions.leerEmbebido(Me.GetType(), "UDO_EXO_DOT.xml")
                objGlobal.conexionSAP.refCompañia.LoadBDFromXML(contenidoXML)
                objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - Validado UDO_EXO_DOT", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

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
            If infoEvento.FormTypeEx = "UDO_FT_EXO_DOT" Then
                If infoEvento.InnerEvent = True Then
                    Select Case infoEvento.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                        Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                            If infoEvento.BeforeAction = False Then
                                If EventHandler_Form_Visible(infoEvento) = False Then
                                    GC.Collect()
                                    Return False
                                End If
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                    End Select
                Else
                    Select Case infoEvento.EventType
                        Case BoEventTypes.et_COMBO_SELECT

                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                            If EventHandler_ItemPressed_After(objGlobal, infoEvento) = False Then
                                Return False
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_CLICK

                        Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                            If infoEvento.BeforeAction = False Then
                                If EventHandler_VALIDATE_after(infoEvento) = False Then
                                    GC.Collect()
                                    Return False
                                End If
                            End If
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
    Private Function EventHandler_ItemPressed_After(ByRef objGlobal As EXO_Generales.EXO_General, ByRef pVal As EXO_Generales.EXO_infoItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.conexionSAP.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.conexionSAP.SBOApp.Forms.Item(pVal.FormUID)
            If pVal.ItemUID = "1" Then
                If oForm.Mode = BoFormMode.fm_ADD_MODE Or oForm.Mode = BoFormMode.fm_UPDATE_MODE Then
                    Dim iRegistros As Integer = oForm.DataSources.DBDataSources.Item("@EXO_DOTL").Size
                    Dim sDOT As String = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_3").Cells.Item(iRegistros).Specific, SAPbouiCOM.EditText).Value
                    If sDOT.Trim = "" Then
                        CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).DeleteRow(iRegistros)
                    End If
                End If
            End If

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
    Private Function EventHandler_Form_Visible(ByRef pVal As EXO_Generales.EXO_infoItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oConds As SAPbouiCOM.Conditions = Nothing
        Dim oCond As SAPbouiCOM.Condition = Nothing

        EventHandler_Form_Visible = False

        Try
            If pVal.ActionSuccess = True Then
                'Recuperar el formulario
                oForm = SboApp.Forms.Item(pVal.FormUID)

                If oForm.Visible = True Then
                    If oForm.Mode = BoFormMode.fm_ADD_MODE Then
                        oForm.DataSources.DBDataSources.Item("@EXO_DOT").SetValue("Code", 0, _sArticulo)
                        oForm.DataSources.DBDataSources.Item("@EXO_DOT").SetValue("Name", 0, _sDescripcion)
                    ElseIf oForm.Mode = BoFormMode.fm_FIND_MODE Then
                        oForm.DataSources.DBDataSources.Item("@EXO_DOT").SetValue("Code", 0, _sArticulo)
                        oForm.DataSources.DBDataSources.Item("@EXO_DOT").SetValue("Name", 0, _sDescripcion)
                    End If
                    oForm.Items.Item("0_U_E").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                End If
            End If

            EventHandler_Form_Visible = True

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
    Private Function EventHandler_VALIDATE_after(ByRef pVal As EXO_Generales.EXO_infoItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oConds As SAPbouiCOM.Conditions = Nothing
        Dim oCond As SAPbouiCOM.Condition = Nothing

        EventHandler_VALIDATE_after = False

        Try
            If pVal.ActionSuccess = True Then
                'Recuperar el formulario
                oForm = SboApp.Forms.Item(pVal.FormUID)
                Dim sMensaje As String = ""
                Dim sAnno As String = "" : Dim sSemana As String = "" : Dim sDOT As String = "" : Dim sAnnoAct As String = Right(Now.Year.ToString("0000"), 2)


                If pVal.ItemUID = "0_U_G" And (pVal.ColUID = "C_0_2" Or pVal.ColUID = "C_0_1" Or pVal.ColUID = "C_0_4") Then
                    sAnno = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value
                    sSemana = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_2").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value

                    If sAnno.Trim.Length = 1 Then
                        sAnno = "0" & sAnno
                    End If
                    If sSemana.Trim.Length = 1 Then
                        sSemana = "0" & sSemana
                    End If
                    sDOT = sSemana & sAnno
                    If sDOT.Length = 4 Then
                        CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_3").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = sDOT
                    Else
                        CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_3").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = ""
                    End If
                    If IsNumeric(sAnno) Then
                        If CInt(sAnno) >= CInt(sAnnoAct) Then
                            sMensaje = "El Año " & sAnno & " debe ser inferior al año actual. "
                            objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            objGlobal.conexionSAP.SBOApp.MessageBox(sMensaje)
                            Exit Function
                        End If
                    End If
                    If IsNumeric(sSemana) Then
                        If CInt(sSemana) > 52 Then
                            sMensaje = "La semana " & sSemana & " debe ser inferior a 52. "
                            objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            objGlobal.conexionSAP.SBOApp.MessageBox(sMensaje)
                            Exit Function
                        End If
                    End If
                End If
                If pVal.ItemUID = "0_U_G" And pVal.ColUID = "C_0_2" And sSemana.Trim <> "" Then
                    Dim iRegistros As Integer = oForm.DataSources.DBDataSources.Item("@EXO_DOTL").Size
                    Dim iRegActivo As Integer = oForm.DataSources.DBDataSources.Item("@EXO_DOTL").Offset + 1
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
End Class
