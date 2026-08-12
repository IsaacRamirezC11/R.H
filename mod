Option Compare Database
Option Explicit

Private Const TABLA As String = "AHORRO_SOLIDARIO"
Private Const FORMULARIO As String = "frmAhorroSolidario"


'==========================================================
' PREPARAR FORMULARIO AL ABRIR
'==========================================================

Public Function PrepararBusqueda() As Boolean

    On Error GoTo ErrorHandler

    Dim frm As Access.Form

    Set frm = Screen.ActiveForm

    '------------------------------------------------------
    ' No mostrar ningún registro inicialmente
    '------------------------------------------------------

    frm.Filter = "1=0"
    frm.FilterOn = True

    '------------------------------------------------------
    ' Vaciar cuadro de búsqueda
    '------------------------------------------------------

    If ControlExiste(frm, "txtBusqueda") Then
        frm.Controls("txtBusqueda").Value = Null
    End If

    PrepararBusqueda = True

    Exit Function


ErrorHandler:

    MsgBox _
        "No se pudo preparar el formulario." & vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbCritical, _
        "AHORRO SOLIDARIO"

End Function



'==========================================================
' BUSCAR POR RFC_COMPLETO O CURP
'==========================================================

Public Function BuscarAhorroSolidario() As Boolean

    On Error GoTo ErrorHandler

    Dim frm As Access.Form
    Dim rs As DAO.Recordset

    Dim valor As String
    Dim criterio As String


    '------------------------------------------------------
    ' OBTENER FORMULARIO ACTIVO
    '------------------------------------------------------

    Set frm = Screen.ActiveForm


    '------------------------------------------------------
    ' OBTENER TEXTO DE BÚSQUEDA
    '------------------------------------------------------

    valor = Trim(Nz(frm.Controls("txtBusqueda").Value, ""))


    '------------------------------------------------------
    ' VALIDAR
    '------------------------------------------------------

    If Len(valor) = 0 Then

        MsgBox _
            "Escribe un RFC_COMPLETO o CURP.", _
            vbExclamation, _
            "AHORRO SOLIDARIO"

        frm.Controls("txtBusqueda").SetFocus

        Exit Function

    End If


    '------------------------------------------------------
    ' PROTEGER COMILLAS
    '------------------------------------------------------

    valor = Replace(valor, "'", "''")


    '------------------------------------------------------
    ' OBTENER LOS REGISTROS DEL FORMULARIO
    '------------------------------------------------------

    Set rs = frm.RecordsetClone


    '------------------------------------------------------
    ' CRITERIO DE BÚSQUEDA
    '
    ' Busca exactamente en:
    '
    ' RFC_COMPLETO
    '       O
    ' CURP
    '------------------------------------------------------

    criterio = _
        "[RFC_COMPLETO] = '" & valor & "'" & _
        " OR [CURP] = '" & valor & "'"


    '------------------------------------------------------
    ' BUSCAR
    '------------------------------------------------------

    rs.FindFirst criterio


    '------------------------------------------------------
    ' NO ENCONTRADO
    '------------------------------------------------------

    If rs.NoMatch Then

        rs.Close
        Set rs = Nothing

        ' Volver a ocultar todos los registros

        frm.Filter = "1=0"
        frm.FilterOn = True

        MsgBox _
            "No se encontró ningún registro para:" & _
            vbCrLf & vbCrLf & _
            Replace(valor, "''", "'"), _
            vbInformation, _
            "AHORRO SOLIDARIO"

        frm.Controls("txtBusqueda").SetFocus

        BuscarAhorroSolidario = False

        Exit Function

    End If


    '------------------------------------------------------
    ' REGISTRO ENCONTRADO
    '------------------------------------------------------

    ' Quitar el filtro inicial 1=0

    frm.FilterOn = False


    ' Ir al registro encontrado

    frm.Bookmark = rs.Bookmark


    '------------------------------------------------------
    ' CERRAR RECORDSET
    '------------------------------------------------------

    rs.Close
    Set rs = Nothing


    BuscarAhorroSolidario = True

    Exit Function


ErrorHandler:

    MsgBox _
        "ERROR AL REALIZAR LA BÚSQUEDA." & _
        vbCrLf & vbCrLf & _
        "Número: " & Err.Number & _
        vbCrLf & _
        "Descripción: " & Err.Description, _
        vbCritical, _
        "AHORRO SOLIDARIO"

    BuscarAhorroSolidario = False

End Function



'==========================================================
' LIMPIAR FORMULARIO
'==========================================================

Public Function LimpiarAhorroSolidario() As Boolean

    On Error GoTo ErrorHandler

    Dim frm As Access.Form

    Set frm = Screen.ActiveForm


    '------------------------------------------------------
    ' OCULTAR TODOS LOS REGISTROS
    '------------------------------------------------------

    frm.Filter = "1=0"
    frm.FilterOn = True


    '------------------------------------------------------
    ' LIMPIAR BÚSQUEDA
    '------------------------------------------------------

    If ControlExiste(frm, "txtBusqueda") Then

        frm.Controls("txtBusqueda").Value = Null

    End If


    '------------------------------------------------------
    ' VOLVER AL CUADRO DE BÚSQUEDA
    '------------------------------------------------------

    If ControlExiste(frm, "txtBusqueda") Then

        frm.Controls("txtBusqueda").SetFocus

    End If


    LimpiarAhorroSolidario = True

    Exit Function


ErrorHandler:

    MsgBox _
        "ERROR AL LIMPIAR." & _
        vbCrLf & vbCrLf & _
        "Número: " & Err.Number & _
        vbCrLf & _
        "Descripción: " & Err.Description, _
        vbCritical, _
        "AHORRO SOLIDARIO"

    LimpiarAhorroSolidario = False

End Function



'==========================================================
' COMPROBAR SI EXISTE UN CONTROL
'==========================================================

Private Function ControlExiste( _
    ByVal frm As Access.Form, _
    ByVal nombre As String) As Boolean

    Dim ctl As Access.Control

    ControlExiste = False


    For Each ctl In frm.Controls

        If ctl.Name = nombre Then

            ControlExiste = True

            Exit Function

        End If

    Next ctl

End Function