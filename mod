Option Compare Database
Option Explicit

Private Const TABLA As String = "AHORRO_SOLIDARIO"
Private Const FORMULARIO As String = "frmAhorroSolidario"


'==========================================================
' CREAR FORMULARIO
'==========================================================
Public Sub CrearFormularioAhorroSolidario()

    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Dim campo As DAO.Field

    Dim frm As Access.Form
    Dim ctl As Access.Control
    Dim etiqueta As Access.Control

    Dim nombreTemporal As String
    Dim nombreControl As String

    Dim x As Long
    Dim y As Long
    Dim contador As Long

    Set db = CurrentDb

    '------------------------------------------------------
    ' Comprobar que existe la tabla
    '------------------------------------------------------
    Set td = db.TableDefs(TABLA)

    '------------------------------------------------------
    ' Eliminar formulario anterior si existe
    '------------------------------------------------------
    If ExisteFormulario(FORMULARIO) Then

        DoCmd.Close acForm, FORMULARIO, acSaveNo

        DoCmd.DeleteObject acForm, FORMULARIO

    End If

    '------------------------------------------------------
    ' Crear formulario nuevo
    '------------------------------------------------------
    Set frm = CreateForm

    nombreTemporal = frm.Name

    '------------------------------------------------------
    ' Propiedades básicas
    '------------------------------------------------------
    frm.Caption = "AHORRO SOLIDARIO - Consulta de personal"

    frm.RecordSource = TABLA

    frm.DefaultView = 0

    frm.ViewsAllowed = 1

    frm.NavigationButtons = False

    frm.RecordSelectors = False

    frm.DividingLines = False

    frm.AutoCenter = True

    frm.AutoResize = False

    frm.Width = 12000

    frm.ScrollBars = 2

    '------------------------------------------------------
    ' ENCABEZADO
    '------------------------------------------------------

    Set etiqueta = CreateControl( _
        nombreTemporal, _
        acLabel, _
        acDetail, _
        , _
        , _
        500, _
        300, _
        10000, _
        500)

    etiqueta.Caption = "AHORRO SOLIDARIO"

    etiqueta.FontSize = 18

    etiqueta.FontWeight = 700


    Set etiqueta = CreateControl( _
        nombreTemporal, _
        acLabel, _
        acDetail, _
        , _
        , _
        500, _
        850, _
        10000, _
        350)

    etiqueta.Caption = "Consulta por RFC_COMPLETO o CURP"

    etiqueta.FontSize = 10


    '------------------------------------------------------
    ' CAMPO DE BÚSQUEDA
    '------------------------------------------------------

    Set etiqueta = CreateControl( _
        nombreTemporal, _
        acLabel, _
        acDetail, _
        , _
        , _
        500, _
        1350, _
        2500, _
        350)

    etiqueta.Caption = "RFC_COMPLETO / CURP:"

    etiqueta.FontWeight = 700


    Set ctl = CreateControl( _
        nombreTemporal, _
        acTextBox, _
        acDetail, _
        , _
        , _
        3000, _
        1300, _
        5000, _
        450)

    ctl.Name = "txtBusqueda"

    ctl.ControlSource = ""

    ctl.FontSize = 11


    '------------------------------------------------------
    ' BOTÓN BUSCAR
    '------------------------------------------------------

    Set ctl = CreateControl( _
        nombreTemporal, _
        acCommandButton, _
        acDetail, _
        , _
        , _
        8200, _
        1300, _
        1300, _
        450)

    ctl.Name = "cmdBuscar"

    ctl.Caption = "BUSCAR"

    ctl.OnClick = "=BuscarAhorroSolidario()"


    '------------------------------------------------------
    ' BOTÓN LIMPIAR
    '------------------------------------------------------

    Set ctl = CreateControl( _
        nombreTemporal, _
        acCommandButton, _
        acDetail, _
        , _
        , _
        9600, _
        1300, _
        1300, _
        450)

    ctl.Name = "cmdLimpiar"

    ctl.Caption = "LIMPIAR"

    ctl.OnClick = "=LimpiarAhorroSolidario()"


    '------------------------------------------------------
    ' MENSAJE DE ESTADO
    '------------------------------------------------------

    Set etiqueta = CreateControl( _
        nombreTemporal, _
        acLabel, _
        acDetail, _
        , _
        , _
        500, _
        1900, _
        10000, _
        400)

    etiqueta.Name = "lblEstado"

    etiqueta.Caption = "Escribe un RFC_COMPLETO o CURP y presiona BUSCAR."

    etiqueta.FontSize = 9


    '------------------------------------------------------
    ' CREAR TODOS LOS CAMPOS DE LA TABLA
    '------------------------------------------------------

    x = 500
    y = 2400
    contador = 0

    For Each campo In td.Fields

        '----------------------------------------------
        ' Etiqueta
        '----------------------------------------------

        Set etiqueta = CreateControl( _
            nombreTemporal, _
            acLabel, _
            acDetail, _
            , _
            , _
            x, _
            y, _
            2500, _
            350)

        etiqueta.Caption = campo.Name

        etiqueta.FontSize = 8

        etiqueta.FontWeight = 700


        '----------------------------------------------
        ' Caja de datos
        '----------------------------------------------

        nombreControl = "txtDato" & CStr(contador)

        Set ctl = CreateControl( _
            nombreTemporal, _
            acTextBox, _
            acDetail, _
            , _
            , _
            x + 2600, _
            y - 30, _
            2600, _
            400)

        ctl.Name = nombreControl

        ' IMPORTANTE:
        ' Los corchetes permiten trabajar con nombres
        ' de campos especiales.
        ctl.ControlSource = "[" & campo.Name & "]"

        ctl.FontSize = 9

        ctl.Locked = True

        ctl.TabStop = False


        '----------------------------------------------
        ' Siguiente campo
        '----------------------------------------------

        contador = contador + 1

        If contador Mod 2 = 0 Then

            x = 500

            y = y + 650

        Else

            x = 5800

        End If

    Next campo


    '------------------------------------------------------
    ' Tamaño del formulario
    '------------------------------------------------------

    frm.Section(acDetail).Height = y + 1000


    '------------------------------------------------------
    ' INICIALMENTE NO MOSTRAR NINGÚN REGISTRO
    '------------------------------------------------------

    frm.Filter = "1=0"

    frm.FilterOn = True


    '------------------------------------------------------
    ' GUARDAR
    '------------------------------------------------------

    DoCmd.Save acForm, nombreTemporal

    DoCmd.Close acForm, nombreTemporal, acSaveYes


    '------------------------------------------------------
    ' CAMBIAR NOMBRE DEL FORMULARIO
    '------------------------------------------------------

    DoCmd.Rename FORMULARIO, acForm, nombreTemporal


    '------------------------------------------------------
    ' ABRIR FORMULARIO
    '------------------------------------------------------

    DoCmd.OpenForm FORMULARIO, acNormal


    MsgBox _
        "Formulario creado correctamente." & vbCrLf & vbCrLf & _
        "El formulario inicia vacío." & vbCrLf & _
        "Busca utilizando RFC_COMPLETO o CURP.", _
        vbInformation, _
        "AHORRO SOLIDARIO"

    Exit Sub


ErrorHandler:

    MsgBox _
        "NO SE PUDO CREAR EL FORMULARIO." & vbCrLf & vbCrLf & _
        "Número de error: " & Err.Number & vbCrLf & _
        "Descripción: " & Err.Description, _
        vbCritical, _
        "AHORRO SOLIDARIO"

End Sub



'==========================================================
' BUSCAR POR RFC_COMPLETO O CURP
'==========================================================
Public Function BuscarAhorroSolidario() As Boolean

    On Error GoTo ErrorHandler

    Dim frm As Access.Form
    Dim valor As String
    Dim filtro As String

    Set frm = Forms(FORMULARIO)

    valor = Trim(Nz(frm.Controls("txtBusqueda").Value, ""))


    '------------------------------------------------------
    ' No se escribió nada
    '------------------------------------------------------

    If Len(valor) = 0 Then

        frm.Controls("lblEstado").Caption = _
            "Escribe un RFC_COMPLETO o CURP."

        MsgBox _
            "Escribe un RFC_COMPLETO o CURP.", _
            vbExclamation, _
            "AHORRO SOLIDARIO"

        BuscarAhorroSolidario = False

        Exit Function

    End If


    '------------------------------------------------------
    ' Evitar problemas con comillas
    '------------------------------------------------------

    valor = Replace(valor, "'", "''")


    '------------------------------------------------------
    ' Buscar en ambos campos
    '------------------------------------------------------

    filtro = _
        "[RFC_COMPLETO] = '" & valor & "'" & _
        " OR " & _
        "[CURP] = '" & valor & "'"


    '------------------------------------------------------
    ' Aplicar filtro
    '------------------------------------------------------

    frm.Filter = filtro

    frm.FilterOn = True


    '------------------------------------------------------
    ' Comprobar resultado
    '------------------------------------------------------

    If frm.RecordsetClone.EOF Then

        frm.FilterOn = False

        frm.Filter = "1=0"

        frm.FilterOn = True

        frm.Controls("lblEstado").Caption = _
            "No se encontró ningún registro."

        MsgBox _
            "No se encontró ningún registro para:" & _
            vbCrLf & vbCrLf & _
            valor, _
            vbInformation, _
            "AHORRO SOLIDARIO"

    Else

        frm.Controls("lblEstado").Caption = _
            "Registro encontrado."

    End If


    BuscarAhorroSolidario = True

    Exit Function


ErrorHandler:

    MsgBox _
        "ERROR AL BUSCAR." & vbCrLf & vbCrLf & _
        "Número: " & Err.Number & vbCrLf & _
        "Descripción: " & Err.Description, _
        vbCritical, _
        "AHORRO SOLIDARIO"

    BuscarAhorroSolidario = False

End Function



'==========================================================
' LIMPIAR
'==========================================================
Public Function LimpiarAhorroSolidario() As Boolean

    On Error GoTo ErrorHandler

    Dim frm As Access.Form

    Set frm = Forms(FORMULARIO)


    ' Quitar búsqueda
    frm.Controls("txtBusqueda").Value = Null


    ' Volver a dejar el formulario vacío
    frm.Filter = "1=0"

    frm.FilterOn = True


    frm.Controls("lblEstado").Caption = _
        "Escribe un RFC_COMPLETO o CURP y presiona BUSCAR."


    LimpiarAhorroSolidario = True

    Exit Function


ErrorHandler:

    MsgBox _
        "ERROR AL LIMPIAR." & vbCrLf & vbCrLf & _
        "Número: " & Err.Number & vbCrLf & _
        "Descripción: " & Err.Description, _
        vbCritical, _
        "AHORRO SOLIDARIO"

    LimpiarAhorroSolidario = False

End Function



'==========================================================
' COMPROBAR SI EXISTE EL FORMULARIO
'==========================================================
Private Function ExisteFormulario(ByVal nombre As String) As Boolean

    Dim objeto As AccessObject

    ExisteFormulario = False

    For Each objeto In CurrentProject.AllForms

        If objeto.Name = nombre Then

            ExisteFormulario = True

            Exit Function

        End If

    Next objeto

End Function