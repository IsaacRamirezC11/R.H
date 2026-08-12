Option Compare Database
Option Explicit

Private Const FORMULARIO As String = "frmAhorroSolidario"

'==========================================================
' AL ABRIR EL FORMULARIO
'==========================================================

Public Sub PrepararFormulario()

    On Error GoTo ErrorHandler

    Dim frm As Form

    Set frm = Forms(FORMULARIO)

    ' No mostrar ningún registro inicialmente
    frm.Filter = "1=0"
    frm.FilterOn = True

    Exit Sub

ErrorHandler:

    MsgBox "Error al preparar el formulario." & vbCrLf & _
           Err.Number & ": " & Err.Description, _
           vbCritical, "Ahorro Solidario"

End Sub


'==========================================================
' BUSCAR POR RFC_COMPLETO O CURP
'==========================================================

Public Function BuscarAhorroSolidario() As Boolean

    On Error GoTo ErrorHandler

    Dim frm As Form
    Dim valor As String
    Dim filtro As String

    Set frm = Forms(FORMULARIO)

    valor = Trim(Nz(frm!txtBusqueda.Value, ""))

    If valor = "" Then

        MsgBox "Escribe un RFC_COMPLETO o CURP.", _
               vbExclamation, "Ahorro Solidario"

        Exit Function

    End If

    ' Evitar problemas con comillas
    valor = Replace(valor, "'", "''")

    ' Buscar en RFC_COMPLETO o CURP
    filtro = "[RFC_COMPLETO] = '" & valor & "'" & _
             " OR [CURP] = '" & valor & "'"

    ' Aplicar búsqueda
    frm.Filter = filtro
    frm.FilterOn = True

    ' Comprobar si encontró algo
    If frm.RecordsetClone.EOF Then

        frm.Filter = "1=0"
        frm.FilterOn = True

        MsgBox "No se encontró ningún registro.", _
               vbInformation, "Ahorro Solidario"

        BuscarAhorroSolidario = False

    Else

        frm.RecordsetClone.MoveFirst

        MsgBox "Registro encontrado.", _
               vbInformation, "Ahorro Solidario"

        BuscarAhorroSolidario = True

    End If

    Exit Function

ErrorHandler:

    MsgBox "Error al realizar la búsqueda." & vbCrLf & _
           Err.Number & ": " & Err.Description, _
           vbCritical, "Ahorro Solidario"

End Function


'==========================================================
' LIMPIAR
'==========================================================

Public Function LimpiarAhorroSolidario() As Boolean

    On Error GoTo ErrorHandler

    Dim frm As Form

    Set frm = Forms(FORMULARIO)

    ' Borrar búsqueda
    frm!txtBusqueda = Null

    ' Ocultar nuevamente todos los registros
    frm.Filter = "1=0"
    frm.FilterOn = True

    frm!txtBusqueda.SetFocus

    LimpiarAhorroSolidario = True

    Exit Function

ErrorHandler:

    MsgBox "Error al limpiar." & vbCrLf & _
           Err.Number & ": " & Err.Description, _
           vbCritical, "Ahorro Solidario"

End Function