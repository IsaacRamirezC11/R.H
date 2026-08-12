Option Compare Database
Option Explicit

Private Const TABLA As String = "AHORRO_SOLIDARIO"
Private Const NOMBRE_FORMULARIO As String = "frmNuevoAhorro"


'==========================================================
' CREAR EL FORMULARIO
'==========================================================
Public Sub CrearFormularioNuevoAhorro()

    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Dim fld As DAO.Field

    Dim frm As Access.Form
    Dim ctl As Access.Control
    Dim lbl As Access.Control

    Dim nombreTemporal As String
    Dim nombreControl As String

    Dim x As Long
    Dim y As Long
    Dim i As Long

    Set db = CurrentDb

    '------------------------------------------------------
    ' Comprobar tabla
    '------------------------------------------------------

    Set td = db.TableDefs(TABLA)

    '------------------------------------------------------
    ' Eliminar formulario anterior
    '------------------------------------------------------

    On Error Resume Next

    DoCmd.Close acForm, NOMBRE_FORMULARIO, acSaveNo
    DoCmd.DeleteObject acForm, NOMBRE_FORMULARIO

    On Error GoTo ErrorHandler

    '------------------------------------------------------
    ' Crear formulario temporal
    '------------------------------------------------------

    Set frm = CreateForm

    nombreTemporal = frm.Name

    '------------------------------------------------------
    ' Propiedades
    '------------------------------------------------------

    frm.Caption = "AHORRO SOLIDARIO - Nuevo registro"

    frm.RecordSource = TABLA

    frm.DefaultView = 0
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    frm.DividingLines = False
    frm.AutoCenter = True
    frm.AutoResize = False

    frm.Width = 12500
    frm.ScrollBars = 2

    ' SOLO REGISTRO NUEVO
    frm.DataEntry = True

    '------------------------------------------------------
    ' TÍTULO
    '------------------------------------------------------

    Set lbl = CreateControl( _
        nombreTemporal, acLabel, acDetail, _
        , , 500, 300, 11000, 500)

    lbl.Caption = "NUEVO REGISTRO - AHORRO SOLIDARIO"
    lbl.FontSize = 18
    lbl.FontWeight = 700

    '------------------------------------------------------
    ' SUBTÍTULO
    '------------------------------------------------------

    Set lbl = CreateControl( _
        nombreTemporal, acLabel, acDetail, _
        , , 500, 850, 11000, 350)

    lbl.Caption = "Capture los datos del nuevo registro."
    lbl.FontSize = 10

    '------------------------------------------------------
    ' BOTÓN GUARDAR
    '------------------------------------------------------

    Set ctl = CreateControl( _
        nombreTemporal, acCommandButton, acDetail, _
        , , 8500, 1200, 1500, 500)

    ctl.Name = "cmdGuardar"
    ctl.Caption = "GUARDAR"
    ctl.OnClick = "=GuardarNuevoAhorro()"

    '------------------------------------------------------
    ' BOTÓN LIMPIAR
    '------------------------------------------------------

    Set ctl = CreateControl( _
        nombreTemporal, acCommandButton, acDetail, _
        , , 10100, 1200, 1500, 500)

    ctl.Name = "cmdLimpiar"
    ctl.Caption = "LIMPIAR"
    ctl.OnClick = "=LimpiarNuevoAhorro()"

    '------------------------------------------------------
    ' ESTADO
    '------------------------------------------------------

    Set lbl = CreateControl( _
        nombreTemporal, acLabel, acDetail, _
        , , 500, 1300, 7000, 350)

    lbl.Name = "lblEstado"
    lbl.Caption = "Listo para capturar un nuevo registro."

    '------------------------------------------------------
    ' CAMPOS
    '------------------------------------------------------

    x = 500
    y = 1900
    i = 0

    For Each fld In td.Fields

        If i Mod 2 = 0 Then
            x = 500
        Else
            x = 6100
        End If

        If i > 0 And i Mod 2 = 0 Then
            y = y + 700
        End If

        '----------------------------------------------
        ' ETIQUETA
        '----------------------------------------------

        Set lbl = CreateControl( _
            nombreTemporal, acLabel, acDetail, _
            , , x, y, 2200, 350)

        lbl.Caption = fld.Name
        lbl.FontSize = 8
        lbl.FontWeight = 700

        '----------------------------------------------
        ' CAMPO
        '----------------------------------------------

        nombreControl = "txtCampo" & CStr(i)

        Set ctl = CreateControl( _
            nombreTemporal, acTextBox, acDetail, _
            , , x + 2300, y - 40, 3000, 420)

        ctl.Name = nombreControl

        ctl.ControlSource = "[" & fld.Name & "]"

        ctl.Tag = fld.Name

        ctl.FontSize = 9

        '----------------------------------------------
        ' AUTONUMÉRICO
        '----------------------------------------------

        If (fld.Attributes And dbAutoIncrField) <> 0 Then
            ctl.Locked = True
            ctl.Enabled = False
        End If

        i = i + 1

    Next fld

    frm.Section(acDetail).Height = y + 1000

    '------------------------------------------------------
    ' GUARDAR
    '------------------------------------------------------

    DoCmd.Save acForm, nombreTemporal

    DoCmd.Close acForm, nombreTemporal, acSaveYes

    '------------------------------------------------------
    ' RENOMBRAR
    '------------------------------------------------------

    DoCmd.Rename NOMBRE_FORMULARIO, acForm, nombreTemporal

    '------------------------------------------------------
    ' ABRIR EL FORMULARIO
    '------------------------------------------------------

    DoCmd.OpenForm NOMBRE_FORMULARIO, acNormal

    Exit Sub

ErrorHandler:

    MsgBox _
        "No se pudo crear el formulario." & vbCrLf & vbCrLf & _
        "Error: " & Err.Number & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "AHORRO SOLIDARIO"

End Sub


'==========================================================
' GUARDAR NUEVO REGISTRO
'==========================================================
Public Function GuardarNuevoAhorro() As Boolean

    On Error GoTo ErrorHandler

    Dim frm As Access.Form
    Dim ctl As Access.Control

    Dim rfc As String
    Dim curp As String
    Dim existe As Long

    Set frm = Screen.ActiveForm

    '------------------------------------------------------
    ' Obtener RFC_COMPLETO
    '------------------------------------------------------

    rfc = ValorCampo(frm, "RFC_COMPLETO")

    If Trim(rfc) = "" Then

        MsgBox _
            "RFC_COMPLETO es obligatorio.", _
            vbExclamation, _
            "AHORRO SOLIDARIO"

        EnfocarCampo frm, "RFC_COMPLETO"

        Exit Function

    End If

    '------------------------------------------------------
    ' Obtener CURP
    '------------------------------------------------------

    curp = ValorCampo(frm, "CURP")

    If Trim(curp) = "" Then

        MsgBox _
            "CURP es obligatorio.", _
            vbExclamation, _
            "AHORRO SOLIDARIO"

        EnfocarCampo frm, "CURP"

        Exit Function

    End If

    '------------------------------------------------------
    ' Comprobar RFC duplicado
    '------------------------------------------------------

    existe = DCount( _
        "*", _
        TABLA, _
        "[RFC_COMPLETO]='" & Replace(rfc, "'", "''") & "'")

    If existe > 0 Then

        MsgBox _
            "Ya existe un registro con ese RFC_COMPLETO.", _
            vbExclamation, _
            "AHORRO SOLIDARIO"

        Exit Function

    End If

    '------------------------------------------------------
    ' Guardar
    '------------------------------------------------------

    If frm.Dirty Then
        DoCmd.RunCommand acCmdSaveRecord
    End If

    frm.Controls("lblEstado").Caption = _
        "Registro guardado correctamente."

    MsgBox _
        "REGISTRO GUARDADO CORRECTAMENTE.", _
        vbInformation, _
        "AHORRO SOLIDARIO"

    GuardarNuevoAhorro = True

    Exit Function

ErrorHandler:

    MsgBox _
        "No se pudo guardar el registro." & vbCrLf & vbCrLf & _
        "Error: " & Err.Number & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "AHORRO SOLIDARIO"

End Function


'==========================================================
' LIMPIAR
'==========================================================
Public Function LimpiarNuevoAhorro() As Boolean

    On Error GoTo ErrorHandler

    Dim frm As Access.Form

    Set frm = Screen.ActiveForm

    If frm.Dirty Then
        frm.Undo
    End If

    DoCmd.GoToRecord , , acNewRec

    frm.Controls("lblEstado").Caption = _
        "Listo para capturar un nuevo registro."

    LimpiarNuevoAhorro = True

    Exit Function

ErrorHandler:

    MsgBox _
        "No se pudo limpiar el formulario." & vbCrLf & vbCrLf & _
        "Error: " & Err.Number & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "AHORRO SOLIDARIO"

End Function


'==========================================================
' OBTENER VALOR DE UN CAMPO
'==========================================================
Private Function ValorCampo( _
    ByVal frm As Access.Form, _
    ByVal nombreCampo As String) As String

    Dim ctl As Access.Control

    For Each ctl In frm.Controls

        If ctl.Tag = nombreCampo Then

            ValorCampo = Trim(Nz(ctl.Value, ""))

            Exit Function

        End If

    Next ctl

    ValorCampo = ""

End Function


'==========================================================
' ENFOCAR CAMPO
'==========================================================
Private Sub EnfocarCampo( _
    ByVal frm As Access.Form, _
    ByVal nombreCampo As String)

    Dim ctl As Access.Control

    For Each ctl In frm.Controls

        If ctl.Tag = nombreCampo Then

            ctl.SetFocus

            Exit Sub

        End If

    Next ctl

End Sub