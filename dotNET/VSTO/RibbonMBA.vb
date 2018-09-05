#Region "Imports"

Imports Microsoft.Office.Tools.Ribbon
Imports System.Collections
Imports System.IO
Imports System.Text.RegularExpressions

#End Region

'Cinta de opciones de Word
Public Class RibbonMBA

#Region "Declaraciones"

    'PathRTFs Procesados
    Private _pathRTF As String

    'Puntero para movimiento entre tablas
    Private _NroTabla As Integer = 0

    'Puntero para movimiento entre shapes
    Private _NroShape As Integer = 0

    'Puntero para movimiento entre inlineshapes
    Private _NroInlineShape As Integer = 0

    'Último patrón de cuadro utilizado
    Dim _UltCuadro As String = ""

    'Enum con tipos de cuadros
    Private Enum TiposCuadro
        UnaCelda = 5
        DosColumnas = 0
        Encabezado = 6
        TresColumnas = 1
        TresColumnasGrande = 4
        CombinadoTresCeldas = 2
        CuatroColumnas = 3
    End Enum

    'Enum con tipos de procesamiento de BO
    Private Enum TipoProcBO
        General = 0
        Específico = 1
    End Enum

    'Constantes
    Const Header As String = "MBA - PrecapturaDoc"

#End Region

#Region "Eventos"

#Region "Formulario"

    'Al cargar la cinta
    private sub ribbonmba_load(byval sender as system.object, byval e as ribbonuieventargs) handles mybase.load
        Try
            'carga parámetros
            _pathrtf = startframe.systemfunctions.files.leertagxml("c:\fenix\operativo\local\startframe.us.loader.exe.config", "pathwordaddin")
            If _pathrtf = String.empty Then
                _pathrtf = "f:\online\pub\rtf-pdf\"
            End If

            'TabPrecapturaDoc.Label = Me.GetType.Assembly.GetName.Version.ToString()

            'pruebas (comentar para poner operativo)
            '_pathrtf = "u:\todo mba\osc\docsprecaptura\"

        Catch ex As exception
            'si da error, toma la configuración por defecto
            '.....para pruebas:
            '_pathrtf = "u:\todo mba\osc\docsprecaptura\"
            '.....para operativo:
            _pathrtf = "f:\online\pub\rtf-pdf\"
        end try
    end sub

#End Region

#Region "Shapes"

    'Convertir InLineShapes en tablas con encabezado BCRA
    Private Sub btnBCRA_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnBCRA.Click
        Try
            With Globals.ThisAddIn.Application.ActiveDocument
                'Validación
                If .InlineShapes.Count = 0 Then
                    MsgBox("No hay imágenes del BCRA con el formato buscado.", MsgBoxStyle.Information, Header)
                    Exit Sub
                End If

                'Recorre los inLineShapes para reemplazarlos
                For i As Integer = .InlineShapes.Count To 1 Step -1
                    Me.ReemplazarImagenBCRA(i, False, Me.chkConfirmar.Checked, False)
                Next

                'Deselecciona todo
                .Range(Start:=0, End:=0).Select()

                'Informa
                MsgBox("Reemplazos de imagen de BCRA finalizado.", MsgBoxStyle.Information, Header)
            End With

        Catch ex As Exception
            MsgBox("Error eliminando los InLineShapes: " & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, Header)
        End Try
    End Sub

    'Convertir todos los Shapes en tablas
    Private Sub btnTabla_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        Try
            With Globals.ThisAddIn.Application.ActiveDocument
                'Recorre los shapes para reemplazarlos
                For i As Integer = .Shapes.Count To 1 Step -1
                    'Agrega una tabla en el lugar del shape
                    Me.ReemplazarShapeXtabla(i)
                Next
                'Deselecciona todo
                .Range(Start:=0, End:=0).Select()
            End With

        Catch ex As Exception
            MsgBox("Error reemplazando los shapes por tablas: " & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, Header)
        End Try
    End Sub

    'Elimina todos los shapes del doc.
    Private Sub btnEliminarShape_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnEliminarShape.Click
        Dim procesar As Boolean = True
        Try
            'Recorre los shapes para eliminarlos
            With Globals.ThisAddIn.Application.ActiveDocument
                'Validación
                If .Shapes.Count = 0 Then
                    MsgBox("No hay shapes para eliminar.", MsgBoxStyle.Information, Header)
                    Exit Sub
                End If

                'Elimina los shapes
                For i As Integer = .Shapes.Count To 1 Step -1
                    procesar = True
                    If Me.chkConfirmar.Checked Then
                        .Shapes.Item(i).Select()
                        With Globals.ThisAddIn.Application.Selection
                            .MoveStart(Unit:=Word.WdUnits.wdLine, Count:=-1)
                            .MoveEnd(Unit:=Word.WdUnits.wdLine, Count:=1)
                            .Move()
                        End With
                        If Startframe.US.Display.MsgBox("¿Confirma el cambio?", "Confirmación del usuario requerida", US.Display.MsgBoxTipos.msgConfirmacion) = Windows.Forms.DialogResult.No Then
                            procesar = False
                        End If
                    End If
                    If procesar Then
                        .Shapes.Item(i).Delete()
                    End If
                Next

                'Deselecciona todo
                .Range(Start:=0, End:=0).Select()

                'Informa
                MsgBox("Shapes eliminados.", MsgBoxStyle.Information, Header)
            End With

        Catch ex As Exception
            MsgBox("Error eliminando los shapes: " & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, Header)
        End Try
    End Sub

    'Reemplaza el shape por el texto de BCRA
    Private Sub btnShapeBCRA_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnShapeBCRA.Click
        With Globals.ThisAddIn.Application.ActiveDocument
            'Validación
            If .Shapes.Count() = 0 OrElse _NroShape = 0 OrElse _NroShape > .Shapes.Count() Then
                MsgBox("Primero seleccione el shape a reemplazar", MsgBoxStyle.Information, Header)
                Exit Sub
            End If

            'Reemplaza por BCRA
            Me.ReemplazarImagenBCRAenShape(_NroShape, True, Me.chkConfirmar.Checked, True)
        End With
    End Sub

    'Elimina el shape pero antes extrae su texto
    Private Sub btnShapeToText_Click(sender As Object, e As RibbonControlEventArgs) Handles btnShapeToText.Click
        With Globals.ThisAddIn.Application.ActiveDocument
            'Validación
            If .Shapes.Count() = 0 OrElse _NroShape = 0 OrElse _NroShape > .Shapes.Count() Then
                MsgBox("Primero seleccione el shape a reemplazar", MsgBoxStyle.Information, Header)
                Exit Sub
            End If

            'Reemplaza por BCRA
            Me.ReemplazarShapeXtexto(_NroShape)
        End With
    End Sub

#End Region

#Region "Formato"

    'Aplica un formato estándar a todo el documento
    Private Sub btnReformatear_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnReformatear.Click
        With Globals.ThisAddIn.Application.ActiveDocument
            'Confirma revisiones
            .Revisions.AcceptAll()

            'Selecciona todo el documento excepto la primer línea
            .Select()
            With .Application.Selection
                .MoveStart(Unit:=Word.WdUnits.wdLine, Count:=1)

                'Fuente estándar
                .Font.Color = Word.WdColor.wdColorBlack
                .Font.Position = 0      'Espaciado entre caracteres (posición)

                'párrafo
                .ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle
                .ParagraphFormat.SpaceAfter = 0
                .ParagraphFormat.LineUnitAfter = 0
                .ParagraphFormat.LineUnitBefore = 0
                .ParagraphFormat.RightIndent = 0

                'página
                '.PageSetup.PaperSize = Word.WdPaperSize.wdPaperA4
                '.PageSetup.LeftMargin = CentimetersToPoints(1)
                Try
                    .PageSetup.RightMargin = CentimetersToPoints(1.5)
                Catch ex As Exception
                    'Informa el error pero continúa con el proceso
                    MsgBox("No se pudo estrablecer el Margen Derecho. " & vbCrLf & _
                           "Configure el mismo manualmente en 1.5 cm. " & vbCrLf & _
                           "El error reportado por Word es: " & vbCrLf & ex.Message)
                End Try
                Try
                    .PageSetup.TopMargin = CentimetersToPoints(2.5)
                Catch ex As Exception
                    'Informa el error pero continúa con el proceso
                    MsgBox("No se pudo estrablecer el Margen Superior. " & vbCrLf & _
                           "Configure el mismo manualmente en 2.5 cm. " & vbCrLf & _
                           "El error reportado por Word es: " & vbCrLf & ex.Message)
                End Try
                '.PageSetup.BottomMargin = CentimetersToPoints(2.5)

                'Reemplazos varios
                With Globals.ThisAddIn.Application.ActiveDocument
                    'Viñetas a texto
                    If .Lists.Count > 0 Then
                        For i As Integer = .Lists.Count To 1 Step -1
                            .Lists(i).ConvertNumbersToText()

                            '...cambia los TABs por ESPACIOS
                            With .Application.Selection.Find
                                .ClearFormatting()
                                .Text = "^t"
                                With .Replacement
                                    .ClearFormatting()
                                    .Text = " "
                                End With
                                .Execute(Format:=True, Replace:=Word.WdReplace.wdReplaceAll)
                            End With
                        Next
                    End If

                    'Resalta el texto en Symbol
                    .Select()
                    With .Application.Selection.Find
                        .ClearFormatting()
                        .Font.Name = "Symbol"
                        With .Replacement
                            .ClearFormatting()
                            .Highlight = 1
                        End With
                        .Execute(Format:=True, Replace:=Word.WdReplace.wdReplaceAll)
                    End With
                    'Tamaño de fuente 11 por 10
                    .Select()
                    With .Application.Selection.Find
                        .ClearFormatting()
                        .Font.Size = 11
                        .Text = ""
                        With .Replacement
                            .ClearFormatting()
                            .Font.Size = 10
                            .Text = ""
                        End With
                        .Execute(Format:=True, Replace:=Word.WdReplace.wdReplaceAll)
                    End With
                    'Tamaño de fuente 12 por 10
                    .Select()
                    With .Application.Selection.Find
                        .ClearFormatting()
                        .Font.Size = 12
                        .Text = ""
                        With .Replacement
                            .ClearFormatting()
                            .Font.Size = 10
                            .Text = ""
                        End With
                        .Execute(Format:=True, Replace:=Word.WdReplace.wdReplaceAll)
                    End With
                    'Tamaño de fuente 14 por 10
                    .Select()
                    With .Application.Selection.Find
                        .ClearFormatting()
                        .Font.Size = 14
                        .Text = ""
                        With .Replacement
                            .ClearFormatting()
                            .Font.Size = 10
                            .Text = ""
                        End With
                        .Execute(Format:=True, Replace:=Word.WdReplace.wdReplaceAll)
                    End With
                    'Tamaño de fuente 16 por 10
                    .Select()
                    With .Application.Selection.Find
                        .ClearFormatting()
                        .Font.Size = 16
                        .Text = ""
                        With .Replacement
                            .ClearFormatting()
                            .Font.Size = 10
                            .Text = ""
                        End With
                        .Execute(Format:=True, Replace:=Word.WdReplace.wdReplaceAll)
                    End With
                    'Tamaño de fuente 18 por 10
                    .Select()
                    With .Application.Selection.Find
                        .ClearFormatting()
                        .Font.Size = 18
                        .Text = ""
                        With .Replacement
                            .ClearFormatting()
                            .Font.Size = 10
                            .Text = ""
                        End With
                        .Execute(Format:=True, Replace:=Word.WdReplace.wdReplaceAll)
                    End With
                    'Caracter N° (?) por Nº (Alt+167)
                    .Select()
                    With .Application.Selection.Find
                        .ClearFormatting()
                        .Text = "N°"
                        With .Replacement
                            .ClearFormatting()
                            .Text = "Nº"
                        End With
                        .Execute(Format:=True, Replace:=Word.WdReplace.wdReplaceAll)
                    End With
                    'Caracter ° (?) por º (Alt+167)
                    .Select()
                    With .Application.Selection.Find
                        .ClearFormatting()
                        .Text = "°"
                        With .Replacement
                            .ClearFormatting()
                            .Text = "º"
                        End With
                        .Execute(Format:=True, Replace:=Word.WdReplace.wdReplaceAll)
                    End With

                    '''Subrayado grueso por normal
                    ''.Select()
                    ''With .Application.Selection.Find
                    ''    .ClearFormatting()
                    ''    .Font.Underline = Word.WdUnderline.wdUnderlineThick
                    ''    .Text = ""
                    ''    With .Replacement
                    ''        .ClearFormatting()
                    ''        .Font.Underline = Word.WdUnderline.wdUnderlineSingle
                    ''        .Text = ""
                    ''    End With
                    ''    .Execute(Format:=True, Replace:=Word.WdReplace.wdReplaceAll)
                    ''End With

                    'Deselecciona todo
                    .Range(Start:=0, End:=0).Select()
                End With
            End With

            'Formato final (después de reemplazos)
            .Select()
            With .Application.Selection
                .MoveStart(Unit:=Word.WdUnits.wdLine, Count:=1)
                'Fuente estándar
                .Font.Name = "Arial"
            End With
            'Resalta el texto en Symbol
            .Select()
            With .Application.Selection.Find
                .ClearFormatting()
                .Font.Name = "Symbol"
                With .Replacement
                    .ClearFormatting()
                    .Highlight = 1
                End With
                .Execute(Format:=True, Replace:=Word.WdReplace.wdReplaceAll)
                .ClearFormatting()
            End With

            'Recorre cada shape del documento para formatear el texto
            For i As Integer = 1 To .Shapes.Count
                With .Shapes(i)
                    '...selecciona el shape
                    .Select()
                    With .Application.Selection
                        Try
                            '...le da formato al texto contenido
                            .Font.Name = "Arial"
                            .Font.Size = 10
                        Catch ex As Exception
                            '...si no tiene texto, da error => ignorar y continuar con el siguiente
                        End Try
                    End With
                End With
            Next

            'Deselecciona todo
            .Range(Start:=0, End:=0).Select()
        End With
    End Sub

    'Transforma las viñetas en texto
    Private Sub btnNumeracion_Click(sender As Object, e As RibbonControlEventArgs) Handles btnNumeracion.Click
        'Reemplazos varios
        With Globals.ThisAddIn.Application.ActiveDocument
            'Viñetas a texto
            If .Lists.Count > 0 Then
                For i As Integer = .Lists.Count To 1 Step -1
                    .Lists(i).ConvertNumbersToText()

                    '...cambia los TABs por ESPACIOS
                    With .Application.Selection.Find
                        .ClearFormatting()
                        .Text = "^t"
                        With .Replacement
                            .ClearFormatting()
                            .Text = " "
                        End With
                        .Execute(Format:=True, Replace:=Word.WdReplace.wdReplaceAll)
                    End With
                Next
            End If
        End With
    End Sub

    'Recorre las tablas del documento y aplica el formato estándar
    Private Sub btnFormatoTablas_Click(ByVal sender As Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnFormatoTablas.Click
        With Globals.ThisAddIn.Application.ActiveDocument

            'Recorre cada shape del documento para extraer las tablas
            'OJO: asume que hay una única tabla por shape
            Dim vecShapesAEliminar As New ArrayList
            For i As Integer = 1 To .Shapes.Count
                With .Shapes(i)
                    '...selecciona el shape
                    .Select()
                    '...busca si hay tablas dentro del mismo
                    If .Application.Selection.Tables.Count > 0 Then
                        '...almacena el número del shape que deberá eliminarse
                        vecShapesAEliminar.Add(i)
                        '...selecciona la tabla y la almacena en el clipboard
                        .Application.Selection.Tables(1).Select()
                        .Application.Selection.Copy()
                        '...selecciona el shape (para insertar la tabla en la posición actual)
                        .Select()
                        '.Application.Selection.MoveStart(Unit:=Word.WdUnits.wdCharacter, Count:=-1)
                        '.Application.Selection.MoveEnd(Unit:=Word.WdUnits.wdCharacter, Count:=1)

                        'Pega la tabla en el documento
                        .Application.Selection.Paste()
                    End If
                End With
            Next

            'Elimina los shapes que contienen sólo tablas (encontrados en el punto anterior)
            If Not vecShapesAEliminar Is Nothing Then
                For i As Integer = vecShapesAEliminar.Count - 1 To 0 Step -1
                    .Shapes.Item(vecShapesAEliminar(i)).Delete()
                Next
                vecShapesAEliminar = Nothing
            End If

            'Recorre cada tabla del documento
            For i As Integer = 1 To .Tables.Count
                With .Tables(i)
                    'Expande el alto de cada ROW
                    .Rows.HeightRule = Word.WdRowHeightRule.wdRowHeightAtLeast

                    'Formatea el párrafo en que se encuentra la tabla
                    '.Select()
                    'FormatearSeleccion(.Application.Selection, False, False)
                    'With .Application.Selection
                    '    .ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                    'End With

                    'Modifica la sangría de cada celda
                    'For r As Single = 1 To .Rows.Count
                    '    For c As Single = 1 To .Columns.Count
                    '        Try
                    '            'Elimina sangría de cada celda
                    '            .Cell(r, c).Range.Select()
                    '            With .Application.Selection
                    '                .ParagraphFormat.LeftIndent = 0
                    '                .ParagraphFormat.RightIndent = 0
                    '                .ParagraphFormat.FirstLineIndent = 0
                    '            End With
                    '        Catch ex As Exception
                    '            'ignora la celda inexistente (span)
                    '        End Try
                    '    Next
                    'Next
                End With
            Next

            'Deselecciona todo
            .Range(Start:=0, End:=0).Select()
            'Informa
            MsgBox("Tablas formateadas.", MsgBoxStyle.Information, Header)
        End With
    End Sub

    'Elimina todos los saltos de sección de todo el documento
    Private Sub btnEliminarSaltosSec_Click(ByVal sender As Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnEliminarSaltosSec.Click
        '1. Controla que no existan más columnas antes de eliminar todas las secciones
        Dim hayColumnas As Boolean = False
        Dim hayPaginasApaisadas As Boolean = False

        'Selecciona todo el documento para realizar la búsqueda
        Globals.ThisAddIn.Application.ActiveDocument.Select()

        'Realiza la búsqueda
        With Globals.ThisAddIn.Application.ActiveDocument.Application.Selection
            'Busca el siguiente salto de sección
            With .Find
                .ClearFormatting()
                .Text = "^n"
                .Forward = True
                .Execute()
            End With

            If .Find.Found = True Then
                hayColumnas = True
            End If
        End With
        'Confirma el proceso
        If hayColumnas Then
            If Startframe.US.Display.MsgBox("Se detectaron secciones con COLUMNAS en el documento." & vbCrLf & "¿Confirma que desea eliminar las secciones de todas formas?", Header, US.Display.MsgBoxTipos.msgConfirmacion) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If
        End If

        '2. Verifica si hay secciones apaisadas
        Globals.ThisAddIn.Application.ActiveDocument.Select()
        With Globals.ThisAddIn.Application.ActiveDocument.Application.Selection
            For i As Integer = 1 To .Sections.Count
                If .Sections(i).PageSetup.SectionStart = Word.WdSectionStart.wdSectionNewPage _
                        AndAlso .Sections(i).PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape Then
                    hayPaginasApaisadas = True
                End If
            Next
        End With
        'Confirma el proceso
        If hayPaginasApaisadas Then
            If Startframe.US.Display.MsgBox("Se detectaron secciones con PÁGINAS APAISADAS." & vbCrLf & "¿Confirma que desea eliminar las secciones de todas formas?", Header, US.Display.MsgBoxTipos.msgConfirmacion) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If
        End If

        '3. Reemplaza secciones
        With Globals.ThisAddIn.Application.ActiveDocument
            .Select()
            With .Application.Selection
                'Reemplaza todos los saltos de sección por blanco
                .Find.Execute(FindText:="^b", ReplaceWith:="", Replace:=Word.WdReplace.wdReplaceAll)
            End With
            'Deselecciona todo
            .Range(Start:=0, End:=0).Select()
        End With

        '4. Avisa de fin de procesamiento
        MsgBox("Secciones eliminadas.", MsgBoxStyle.Information, Header)
    End Sub

    'Elimina el formato del párrafo y luego le aplica el formato estándar
    Private Sub btnBorrarFormato_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnBorrarFormato.Click
        Me.FormatearSeleccion(Globals.ThisAddIn.Application.ActiveDocument.Application.Selection)
    End Sub

    'Alínea el texto seleccionado con formato justificado
    Private Sub btnAlinearJus_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnAlinearJus.Click
        With Globals.ThisAddIn.Application.ActiveDocument.Application.Selection
            .ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
        End With
    End Sub

    'Alínea el texto seleccionado a la izquierda
    Private Sub btnAlinearIzq_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnAlinearIzq.Click
        With Globals.ThisAddIn.Application.ActiveDocument.Application.Selection
            .ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
        End With
    End Sub

    'Convierte las tablas del documento a texto
    Private Sub btnTablasAtexto_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnTablasAtexto.Click
        With Globals.ThisAddIn.Application.ActiveDocument

            'Recorre cada tabla del documento para convertirlas en texto
            For i As Integer = 1 To .Tables.Count
                With .Tables(i)
                    .ConvertToText()
                End With
            Next

            'Deselecciona todo
            .Range(Start:=0, End:=0).Select()
            'Informa
            MsgBox("Tablas convertidas.", MsgBoxStyle.Information, Header)
        End With
    End Sub

    'Convierte el texto seleccionado en una tabla
    Private Sub btnTextoAtabla_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnTextoAtabla.Click
        Try
            With Globals.ThisAddIn.Application.ActiveDocument.Application.Selection
                'Convierte la selección en tabla
                .ConvertToTable()
                'Configura la tabla
                With .Application.Selection.Tables(1)
                    .Borders(Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                    .Borders(Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                    .Borders(Word.WdBorderType.wdBorderLeft).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                    .Borders(Word.WdBorderType.wdBorderRight).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                    .Borders(Word.WdBorderType.wdBorderHorizontal).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                    .Borders(Word.WdBorderType.wdBorderVertical).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                End With
            End With

        Catch ex As Exception
            MsgBox("Error tratando de interpretar el formato de la tabla.", MsgBoxStyle.Exclamation, Header)
        End Try
    End Sub

    'Acepta las revisiones y las marca como novedades
    Private Sub btnMarcarNovedades_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnMarcarNovedades.Click
        With Globals.ThisAddIn.Application.ActiveDocument
            'Recorre cada revisión pendiente
            Dim i As Integer = 1
            Do While i <= .Revisions.Count
                With .Revisions(i)
                    If .Type = Word.WdRevisionType.wdRevisionDelete Then
                        'acepta el texto eliminado
                        .Accept()
                    Else
                        'acepta y marca el texto como novedad
                        .Range.Select()
                        .Accept()
                        With Globals.ThisAddIn.Application.Selection
                            .Font.Color = Word.WdColor.wdColorRed
                            .Font.Italic = True
                        End With
                    End If
                End With
            Loop

            'Deselecciona todo
            .Range(Start:=0, End:=0).Select()
            'Informa
            MsgBox("Revisiones aceptadas y marcadas como novedad.", MsgBoxStyle.Information, Header)
        End With
    End Sub

#End Region

#Region "Patrones"

    'Busca y reemplaza patrones de 1er pág. en el documento
    Private Sub grpPatPag1_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles cboPatPag1.Click
        Dim item As String = Me.cboPatPag1.SelectedItem.Tag

        Try
            Select Case item
                Case "IEF"
                    btnSgteShape_Click(sender, e)
                    If MsgBox("Este es el primer shape del documento." & vbCrLf & "¿Continúa con el patrón IEF?", MsgBoxStyle.Question & MsgBoxStyle.YesNo, Header) = MsgBoxResult.Yes Then
                        Me.PatronIEF()
                    Else
                        'No continúa con el patrón
                        If MsgBox("¿Desea eliminar el shape seleccionado?", MsgBoxStyle.Question & MsgBoxStyle.YesNo, Header) = MsgBoxResult.Yes Then
                            Globals.ThisAddIn.Application.ActiveDocument.Application.Selection.Delete()
                        End If
                    End If
                Case "REF"
                    'Se eliminó este patrón desde el menú
                    Me.PatronREF()
                Case "FIRMAS"
                    Me.FirmasEnTabla(Globals.ThisAddIn.Application.ActiveDocument.Application.Selection)
                Case "ENCABEZADO"
                    With Globals.ThisAddIn.Application.ActiveDocument
                        Dim texto As String
                        texto = .Application.Selection.Text.Trim
                        'texto seleccionado?
                        If texto Is Nothing OrElse Len(texto) < 10 Then
                            MsgBox("Debe seleccionar un texto para aplicar este patrón.", MsgBoxStyle.Exclamation, Header)
                            Return
                        Else
                            'inserta ENTER al final
                            With .Application.Selection
                                .InsertParagraphAfter()
                                .InsertAfter(vbCrLf)
                            End With
                        End If
                    End With
                    'aplica el patrón de cuadro de encabezado
                    Me.PatronCuadro(2, 1, 2, TiposCuadro.Encabezado)
                Case Else
                    MsgBox("Patrón no implementado", MsgBoxStyle.Exclamation, Header)
            End Select

        Catch ex As Exception
            MsgBox("Error procesando el patrón " & item & ":" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, Header)
        End Try
    End Sub

    'Reemplaza la selección por patrones de cuadros
    Private Sub cboPatCuadros_ButtonClick(ByVal sender As Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) _
                                        Handles cboPatCuadros.ButtonClick, btnReemplazarXcuadro.Click
        Dim item As String = e.Control.Id
        Dim bUltPatron As Boolean = False

        If item = "btnReemplazarXcuadro" And _UltCuadro <> "" Then
            'Aplica el último patrón de cuadro utilizado
            bUltPatron = True
            item = _UltCuadro
        Else
            'Almacena el últ. patrón utilizado
            _UltCuadro = item
        End If

        Try
            'Aplica el patrón seleccionado
            Select Case item
                Case "btnCuadro1"
                    Me.btnReemplazarXcuadro.Image = Global.WordAddInPrecaptura.My.Resources.Resources.icoCuadro1
                    PatronCuadro(2, 1, 2, TiposCuadro.DosColumnas)
                Case "btnCuadro2"
                    Me.btnReemplazarXcuadro.Image = Global.WordAddInPrecaptura.My.Resources.Resources.icoCuadro2
                    PatronCuadro(3, 2, 2, TiposCuadro.CombinadoTresCeldas)
                Case "btnCuadro3"
                    Me.btnReemplazarXcuadro.Image = Global.WordAddInPrecaptura.My.Resources.Resources.icoCuadro3
                    PatronCuadro(3, 1, 3, TiposCuadro.TresColumnas)
                Case "btnCuadro4"
                    Me.btnReemplazarXcuadro.Image = Global.WordAddInPrecaptura.My.Resources.Resources.icoCuadro4
                    PatronCuadro(4, 1, 4, TiposCuadro.CuatroColumnas)
                Case "btnCuadro5"
                    Me.btnReemplazarXcuadro.Image = Global.WordAddInPrecaptura.My.Resources.Resources.icoCuadro3b
                    PatronCuadro(0, 1, 3, TiposCuadro.TresColumnasGrande)
                Case "btnCuadro6"
                    Me.btnReemplazarXcuadro.Image = Global.WordAddInPrecaptura.My.Resources.Resources.icoCuadro5
                    PatronCuadro(0, 1, 1, TiposCuadro.UnaCelda)
                Case "btnCuadro7"
                    Me.btnReemplazarXcuadro.Image = Global.WordAddInPrecaptura.My.Resources.Resources.icoCuadro7
                    PatronCuadro(2, 1, 2, TiposCuadro.Encabezado)
                Case Else
                    Me.btnReemplazarXcuadro.Image = Global.WordAddInPrecaptura.My.Resources.Resources.icoAyuda
                    MsgBox("Patrón no implementado", MsgBoxStyle.Exclamation, Header)
            End Select

            If bUltPatron And _UltCuadro <> "" And _UltCuadro <> "btnReemplazarXcuadro" Then
                'Busca la siguiente ocurrencia
                btnSgteColumna_Click(sender, e)
            End If

        Catch ex As Exception
            MsgBox("Error procesando el patrón " & item & ":" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, Header)
        End Try
    End Sub

    'Procesa párrafos en columnas convirtiendo el texto
    Private Sub cboPatCol_ButtonClick(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles cboPatCol.ButtonClick
        Dim item As String = e.Control.Id

        Try
            Select Case item
                Case "btnColConv1"
                    FormatearColumnas(Globals.ThisAddIn.Application.ActiveDocument.Application.Selection)
                Case "btnCol2Arriba"
                    PatronCol2(Globals.ThisAddIn.Application.ActiveDocument.Application.Selection)
                Case "btnColOracion2Col"
                    PatronCol1(Globals.ThisAddIn.Application.ActiveDocument.Application.Selection)
                Case Else
                    MsgBox("Patrón no implementado", MsgBoxStyle.Exclamation, Header)
            End Select

        Catch ex As Exception
            MsgBox("Error procesando el patrón " & item & ":" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, Header)
        End Try

    End Sub

#End Region

#Region "Ir A"

    'Siguiente tabla
    Private Sub btnSgteTabla_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnSgteTabla.Click
        With Globals.ThisAddIn.Application.ActiveDocument
            'Validación
            If .Tables.Count() = 0 Then
                MsgBox("No existen tablas en este documento", MsgBoxStyle.Information, Header)
                Exit Sub
            End If

            'Mueve el puntero a la siguiente tabla
            If chkFordward.Checked Then
                _NroTabla += 1
                If _NroTabla > .Tables.Count() Then
                    MsgBox("No existe tabla siguiente", MsgBoxStyle.Information, Header)
                    _NroTabla = .Tables.Count()
                    chkFordward.Checked = False
                End If
            Else
                _NroTabla -= 1
                If _NroTabla < 1 Then
                    MsgBox("No existe tabla anterior", MsgBoxStyle.Information, Header)
                    _NroTabla = 1
                    chkFordward.Checked = True
                End If
            End If

            'Se posiciona en la tabla en cuestión
            Try
                .Tables(_NroTabla).Range.Select()
            Catch ex As Exception
                'No encuentra dicha tabla
            End Try
        End With
    End Sub

    'Siguiente shape
    Private Sub btnSgteShape_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnSgteShape.Click
        With Globals.ThisAddIn.Application.ActiveDocument
            'Validación
            If .Shapes.Count() = 0 Then
                MsgBox("No existen shapes en este documento", MsgBoxStyle.Information, Header)
                Exit Sub
            End If

            'Mueve el puntero al siguiente shape
            _NroShape += 1
            If _NroShape > .Shapes.Count() Then
                _NroShape = 1
            End If


            'Se posiciona en el shape en cuestión
            Try
                .Shapes(_NroShape).Select()
                .Shapes(_NroShape).Application.Selection.Cut()
                Globals.ThisAddIn.Application.ActiveDocument.Undo()

            Catch ex As Exception
                'No encuentra dicha tabla
            End Try
        End With
    End Sub

    'Siguiente InlineShape
    Private Sub btnSgteInlineShape_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnSgteInlineShape.Click
        With Globals.ThisAddIn.Application.ActiveDocument
            'Validación
            If .InlineShapes.Count() = 0 Then
                MsgBox("No existen inlineshapes en este documento", MsgBoxStyle.Information, Header)
                Exit Sub
            End If

            'Mueve el puntero al siguiente shape
            _NroInlineShape += 1
            If _NroInlineShape > .InlineShapes.Count() Then
                _NroInlineShape = 1
            End If


            'Se posiciona en el shape en cuestión
            Try
                .InlineShapes(_NroInlineShape).Select()
                '.InlineShapes(_NroInlineShape).Application.Selection.Cut()
                'Globals.ThisAddIn.Application.ActiveDocument.Undo()

            Catch ex As Exception
                'No encuentra dicha tabla
            End Try
        End With
    End Sub

    'Siguiente columna
    Private Sub btnSgteColumna_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnSgteColumna.Click
        'Almacena las coord. de selección actual
        Dim inicio As Integer = Globals.ThisAddIn.Application.ActiveDocument.Application.Selection.Start
        Dim fin As Integer = Globals.ThisAddIn.Application.ActiveDocument.Application.Selection.End

        'Selecciona todo el documento para realizar la búsqueda
        Globals.ThisAddIn.Application.ActiveDocument.Select()

        'Realiza la búsqueda
        With Globals.ThisAddIn.Application.ActiveDocument.Application.Selection
            If chkFordward.Checked Then
                '=> el inicio de la búsqueda es el fin de la selección anterior
                .Start = fin
            Else
                '=> el final de la búsqueda es el inicio de la selección anterior
                .End = inicio
            End If

            'Busca el siguiente salto de sección
            With .Find
                .ClearFormatting()
                .Text = "^n"
                If chkFordward.Checked Then
                    .Forward = True
                Else
                    .Forward = False
                End If
                .Execute()
            End With

            If .Find.Found = False Then
                Globals.ThisAddIn.Application.ActiveDocument.Application.ActiveDocument.Range(Start:=0, End:=0).Select()
                MsgBox("No existen más ocurrencias", MsgBoxStyle.Information, Header)
            Else
                'Mueve la selección para que se vea
                .MoveStart(Unit:=Word.WdUnits.wdSection, Count:=-1)
                .MoveStart(Unit:=Word.WdUnits.wdCharacter, Count:=-1)
                .MoveEnd(Unit:=Word.WdUnits.wdSection, Count:=1)
            End If
        End With
    End Sub

    'Abre un diálogo de búsqueda de "- " y reemplazo por "" a partir de la posición actual
    Private Sub btnBuscarGuiones_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnBuscarGuiones.Click
        'Configura la búsqueda
        With Globals.ThisAddIn.Application.ActiveDocument.Application.Selection.Find
            .ClearFormatting()
            .Replacement.ClearFormatting()
            .Text = "- "
            .Replacement.Text = ""
            .Forward = True
        End With

        'Define el diálogo a mostrar
        Dim missing As Object = Type.Missing
        Dim dialog As Word.Dialog = Globals.ThisAddIn.Application.Dialogs(Word.WdWordDialog.wdDialogEditReplace)
        'Info de diálogos disponibles
        'http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdworddialog(office.11).aspx

        'Muestra el diálogo
        Try
            dialog.Display(missing)
        Catch ex As Exception
            'No muestra errores por fallo en la búsqueda
        End Try
    End Sub

    'Abre un diálogo de búsqueda de subrayado grueso y reemplazo por subrayado simple a partir de la posición actual
    Private Sub btnBuscarSubrayado_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnBuscarSubrayado.Click
        'Configura la búsqueda
        With Globals.ThisAddIn.Application.ActiveDocument.Application.Selection.Find
            .ClearFormatting()
            .Replacement.ClearFormatting()
            .Text = ""
            .Forward = True
            .Font.Underline = Word.WdUnderline.wdUnderlineThick
            With .Replacement
                .ClearFormatting()
                .Font.Underline = Word.WdUnderline.wdUnderlineSingle
                .Text = ""
            End With
        End With

        'Define el diálogo a mostrar
        Dim missing As Object = Type.Missing
        Dim dialog As Word.Dialog = Globals.ThisAddIn.Application.Dialogs(Word.WdWordDialog.wdDialogEditReplace)

        'Muestra el diálogo
        Try
            dialog.Display(missing)
        Catch ex As Exception
            'No muestra errores por fallo en la búsqueda
        End Try
    End Sub

#End Region

#Region "Información"

    'Info del documento
    Private Sub btnInfoDoc_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnInfoDoc.Click
        Dim msg As String
        With Globals.ThisAddIn.Application.ActiveDocument
            msg = String.Format("Este documento tiene:{0}{1} Tablas,{0}{2} Secciones,{0}{3} Shapes,{0}{4} InLineShapes",
                vbCrLf & vbTab,
                .Tables.Count().ToString.Trim,
                .Sections.Count.ToString.Trim,
                .Shapes.Count(),
                .InlineShapes.Count())
        End With

        MsgBox(msg, MsgBoxStyle.Information, Header)
    End Sub

    'Preview en la Tedit
    Private Sub btnTedit_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnTedit.Click
        Dim nombreSinExt, nombre, path As String
        Dim sArchivoOriginal, sArchivoAdelgazado, sArchivoFinal As String
        Dim fi As FileInfo
        Dim nKBorig, nKBadel, nKBtedit As Integer
        Dim msg As New StringBuilder

        'Alerta
        Dim mensaje As New AvisoAlerta
        mensaje.lblMensaje.Text = "Grabando..."
        mensaje.Show()

        Try
            'Nombre del archivo
            With Globals.ThisAddIn.Application.ActiveDocument
                '...con extensión, sin path (Ej.: A_4187.rtf)
                nombre = .Name
                '...sólo la ruta (ej.: c:\docs\)
                path = .FullName.Replace(nombre, "")
                '...sólo el nombre (sin ext. ni ruta ni sufijos de procesamiento previo)
                nombreSinExt = nombre.Substring(0, nombre.LastIndexOf("."))
                nombreSinExt = nombreSinExt.Replace("_precaptura", "")
                nombreSinExt = nombreSinExt.Replace("_adelgazado", "")

                'Graba el RTF procesado y lo cierra
                sArchivoOriginal = path & nombreSinExt & "_precaptura.RTF"
                .SaveAs(FileName:=CObj(sArchivoOriginal), FileFormat:=Word.WdSaveFormat.wdFormatRTF)
                .Close()
            End With

            'Adelgaza el documento
            mensaje.lblMensaje.Text = "Adelgazando..."
            mensaje.Refresh()
            sArchivoFinal = _pathRTF & nombreSinExt & ".RTF"
            Me.AdelgazarRtf(sArchivoOriginal, sArchivoFinal)

            'Info de los docs. procesados
            fi = New FileInfo(sArchivoOriginal)
            nKBorig = CInt(fi.Length / 1024)

            sArchivoAdelgazado = sArchivoOriginal.Substring(0, sArchivoOriginal.LastIndexOf(".")) & "_adelgazado.rtf"
            fi = New FileInfo(sArchivoAdelgazado)
            nKBadel = CInt(fi.Length / 1024)

            fi = New FileInfo(sArchivoFinal)
            nKBtedit = CInt(fi.Length / 1024)

            mensaje.lblMensaje.Text = "Verificando..."
            mensaje.Refresh()

            'Informa del resultado
            mensaje.Close()
            msg.AppendFormat("Proceso Finalizado.{0}", vbCrLf)
            msg.AppendFormat("1. Archivo original ({2} KB): {1}{0}", vbCrLf, sArchivoOriginal, nKBorig)
            msg.AppendFormat("2. Archivo adelgazado ({2} KB): {1}{0}", vbCrLf, sArchivoAdelgazado, nKBadel)
            msg.AppendFormat("3. Archivo final ({2} KB): {1}{0}", vbCrLf, sArchivoFinal, nKBtedit)
            msg.AppendLine()
            msg.AppendFormat("Ahora se mostrará el resultado en la TEedit.{0}Luego deberá capturar el archivo final en el Manager.", vbCrLf)
            StartFrame.US.Display.MsgBox(msg.ToString, Header, US.Display.MsgBoxTipos.msgInformacion)

            'Muestra el resultado en la TEedit
            Dim frm As New frmTedit
            frm.ArchivoRTF = sArchivoFinal
            frm.ShowDialog()

        Catch ex As Exception
            MsgBox("Error al grabar el archivo: " & ex.Message, MsgBoxStyle.Exclamation, Header)

        Finally
            fi = Nothing
            msg = Nothing
            mensaje.Close()
            mensaje.Dispose()
            mensaje = Nothing
        End Try
    End Sub

#End Region

#Region "Boletín Oficial"

    'Aplica un formato específico de BO a todo el documento
    Private Sub btnBOformatear_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnBOformatear.Click
        With Globals.ThisAddIn.Application.ActiveDocument
            'Confirma revisiones
            .Revisions.AcceptAll()

            'Selecciona todo el documento excepto la primer línea
            .Select()
            With .Application.Selection
                'OV: anulado temporalmente por pedido de MAT
                '.MoveStart(Unit:=Word.WdUnits.wdLine, Count:=1)

                'Fuente estándar
                .Font.Color = Word.WdColor.wdColorBlack
                .Font.Size = 10
                .Font.Name = "Arial"

                'párrafo
                .ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle
                .ParagraphFormat.SpaceAfter = 0
                .ParagraphFormat.LineUnitAfter = 0
                .ParagraphFormat.LineUnitBefore = 0
                .ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle
                .ParagraphFormat.LeftIndent = 0
                .ParagraphFormat.RightIndent = 0
            End With

            'Formato puntual del BO
            Me.FormatearBO(TipoProcBO.Específico)

            'Deselecciona todo
            .Range(Start:=0, End:=0).Select()
        End With
    End Sub

    'Aplica un formato estándar genérico a todo el documento
    Private Sub btnBOformatearGen_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnBOformatearGen.Click
        With Globals.ThisAddIn.Application.ActiveDocument
            'Confirma revisiones
            .Revisions.AcceptAll()

            'Selecciona todo el documento excepto la primer línea
            .Select()
            With .Application.Selection
                'OV: anulado temporalmente por pedido de MAT
                '.MoveStart(Unit:=Word.WdUnits.wdLine, Count:=1)

                'Fuente estándar
                .Font.Color = Word.WdColor.wdColorBlack
                .Font.Size = 10
                .Font.Name = "Arial"

                'párrafo
                .ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle
                .ParagraphFormat.SpaceAfter = 0
                .ParagraphFormat.LineUnitAfter = 0
                .ParagraphFormat.LineUnitBefore = 0
                .ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle
                .ParagraphFormat.LeftIndent = 0
                .ParagraphFormat.RightIndent = 0
            End With

            'Formato puntual del BO
            Me.FormatearBO(TipoProcBO.General)

            'Deselecciona todo
            .Range(Start:=0, End:=0).Select()
        End With
    End Sub

#End Region

#Region "Check List"

    'Elimina los Enters de todas las tablas del documento
    'OSC: en desuso 2016-10
    Private Sub btnEliminaEntersTabla_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        With Globals.ThisAddIn.Application.ActiveDocument
            'Recorre cada tabla del documento
            For i As Integer = 1 To .Tables.Count
                With .Tables(i)
                    .Select()
                    With .Application.Selection
                        'Elimina objetos en el texto
                        '...Enter
                        .Find.Execute(FindText:="^p", ReplaceWith:="", Replace:=Word.WdReplace.wdReplaceAll)
                    End With
                    .Select()
                    With .Application.Selection
                        'Elimina objetos en el texto
                        '...Shift+Enter
                        .Find.Execute(FindText:="^l", ReplaceWith:="", Replace:=Word.WdReplace.wdReplaceAll)
                    End With
                End With
            Next

        End With
    End Sub

    'Siguiente shape
    Private Sub btnSgteShape2_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        With Globals.ThisAddIn.Application.ActiveDocument
            'Validación
            If .Shapes.Count() = 0 Then
                MsgBox("No existen shapes en este documento", MsgBoxStyle.Information, Header)
                Exit Sub
            End If

            'Mueve el puntero al siguiente shape
            _NroShape += 1
            If _NroShape > .Shapes.Count() Then
                MsgBox("No existen shapes en este documento", MsgBoxStyle.Information, Header)
                Exit Sub
            End If

            'Se posiciona en el shape en cuestión
            Try
                .Shapes(_NroShape).Select()
                .Shapes(_NroShape).Application.Selection.Cut()
                Globals.ThisAddIn.Application.ActiveDocument.Undo()

            Catch ex As Exception
                'No encuentra dicha tabla
            End Try
        End With
    End Sub

    'Elimina los Enters del texto seleccionado
    Private Sub btnEliminaEntersTexto_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        With Globals.ThisAddIn.Application.ActiveDocument.Application.Selection
            'Elimina objetos en el texto
            '...Enter
            .Find.Execute(FindText:="^p", ReplaceWith:="", Replace:=Word.WdReplace.wdReplaceAll)
            '...Shift+Enter
            .Find.Execute(FindText:="^l", ReplaceWith:="", Replace:=Word.WdReplace.wdReplaceAll)
        End With
    End Sub

    'Reemplaza los Enters del texto seleccionado por espacio
    Private Sub btnReemplazaEntersPorEspacioTextoSeleccionado_Click(sender As Object, e As RibbonControlEventArgs)
        With Globals.ThisAddIn.Application.ActiveDocument.Application.Selection
            'Reemplaza objetos en el texto
            '...Enter
            .Find.Execute(FindText:="^p", ReplaceWith:=" ", Replace:=Word.WdReplace.wdReplaceAll)
            '...Shift+Enter
            .Find.Execute(FindText:="^l", ReplaceWith:=" ", Replace:=Word.WdReplace.wdReplaceAll)
        End With
    End Sub

    'Convierte la selección en una tabla de una fila y n columnas
    Private Sub btnTextoAtabla2_Click(sender As Object, e As RibbonControlEventArgs)
        Try
            With Globals.ThisAddIn.Application.ActiveDocument.Application.Selection
                'No selecciona el último ENTER
                .MoveEnd(Unit:=Word.WdUnits.wdCharacter, Count:=-1)
                'Reemplaza ENTERS por TABs
                .Find.Execute(FindText:="^p", ReplaceWith:="^t", Replace:=Word.WdReplace.wdReplaceAll)
                'Convierte la selección en tabla
                .ConvertToTable()
                'Configura la tabla
                With .Application.Selection.Tables(1)
                    .Borders(Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                    .Borders(Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                    .Borders(Word.WdBorderType.wdBorderLeft).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                    .Borders(Word.WdBorderType.wdBorderRight).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                    .Borders(Word.WdBorderType.wdBorderHorizontal).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                    .Borders(Word.WdBorderType.wdBorderVertical).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                End With
            End With

        Catch ex As Exception
            MsgBox("Error tratando de interpretar el formato de la tabla.", MsgBoxStyle.Exclamation, Header)
        End Try
    End Sub

    'Corrige el archivo para generar el .SAN
    Private Sub btnCorregirParaSan_Click(sender As Object, e As RibbonControlEventArgs)

        With Globals.ThisAddIn.Application.ActiveDocument

            'Busca si hay shapes
            If .Shapes.Count() > 0 Then
                If US.Display.MsgBox("Existen SHAPES en este documento." & vbCrLf & "¿Está seguro que desea generar el .SAN?", "Confirmación Requerida", US.Display.MsgBoxTipos.msgConfirmacion) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If
            End If

            'Busca un salto de sección de columna
            '...Selecciona todo el documento para realizar la búsqueda
            .Select()
            '...Realiza la búsqueda
            With .Application.Selection
                With .Find
                    .ClearFormatting()
                    .Text = "^n"
                    .Forward = True
                    .Execute()

                    If .Found = True Then
                        If StartFrame.US.Display.MsgBox("Se detectaron secciones con COLUMNAS en el documento." & vbCrLf & "¿Está seguro que desea generar el .SAN?", Header, US.Display.MsgBoxTipos.msgConfirmacion) = Windows.Forms.DialogResult.No Then
                            Globals.ThisAddIn.Application.ActiveDocument.Range(Start:=0, End:=0).Select()
                            Exit Sub
                        End If
                    End If
                End With
            End With
            '...Deselecciona todo
            .Range(Start:=0, End:=0).Select()

            Dim _tabla, _fila, _columna As Integer
            Dim _filaActual As Integer = 0
            Dim textoCelda As String

            'variables para los parrafos que estan fuera de las tablas
            Dim _begin As Integer = 0
            Dim _end As Integer = 0
            Dim rngParagraphs As Word.Range

            'intento obtener el tipo y numero de la norma del encabezado
            Dim tipo As String = String.Empty
            Dim numero As String = String.Empty
            Dim rg As Word.Range
            rg = .Sentences(1)
            rg.Select()
            For s As Integer = 2 To 5 ' tiene que estar dentro de los primeros 5 parrafos
                rg.MoveEnd(Unit:=Word.WdUnits.wdSentence, Count:=1)
                rg.Select()
                If .Application.Selection.Text.ToUpper.IndexOf("COMUNICACIÓN") <> -1 Then
                    rg = .Sentences(s)
                    rg.Select()
                    Dim textoSeleccion As String = .Application.Selection.Text
                    Dim match As Match = Regex.Match(textoSeleccion, "COMUNICACI.N\s+.(?<TIPO>[A-Z]).\s+(?<NRO>[0-9]+)", RegexOptions.Compiled Or RegexOptions.IgnoreCase Or RegexOptions.Multiline)
                    tipo = match.Groups("TIPO").Value.Trim
                    numero = match.Groups("NRO").Value.Trim
                    Exit For
                End If
            Next

            Dim pathDOC As String = .Path
            Dim pathSAN As String = .Path.ToLower.Replace("\docs", "\texto")
            If Not Directory.Exists(pathSAN) Then pathSAN = pathDOC 'si no existe el SAN guardo la salida en el directio del DOC

            Dim nombreSAN As String = String.Empty
            If tipo.Length > 0 AndAlso numero.Length > 0 Then
                nombreSAN = If(tipo.ToUpper = "C", "0C", "09") & numero & ".SAN"
            Else
                nombreSAN = .Name.Substring(0, .Name.LastIndexOf(".")) & ".SAN"
            End If

            Dim archivoSalida As String = Path.Combine(pathSAN, nombreSAN)
            If File.Exists(archivoSalida) Then File.Delete(archivoSalida)

            Dim sw As StreamWriter = New StreamWriter(archivoSalida, True, System.Text.Encoding.Default)

            Try

                'verifico si hay alguna tabla con encabezado, si hay paso por el proceso de tablas, si no lo mando derecho al proceso de comunicaciones sin tablas.
                Dim esComunicacionSimple As Boolean = True
                If .Tables.Count > 0 Then
                    For _tabla = 1 To .Tables.Count
                        If .Tables(_tabla).Cell(1, 1).Range.Text.Trim.Length >= 6 AndAlso .Tables(_tabla).Cell(1, 1).Range.Text.Trim.ToLower.Substring(0, 6) = "oficio" Then
                            'Es el header
                            esComunicacionSimple = False
                            Exit For
                        End If
                    Next _tabla
                End If

                If Not esComunicacionSimple Then

                    '*****************************
                    '* comunicaciones con tablas *
                    '*****************************

                    Dim aTablaCompleta(20, 0) As String
                    Dim aTieneBordeInferior(20, 0) As Boolean

                    Dim nColOficio As Integer = 1
                    Dim nColExpte As Integer = 3
                    Dim nColJurisdiccion As Integer = 0
                    Dim nColMonto As Integer = 0

                    For _tabla = 1 To .Tables.Count

                        Dim cantColumnas As Integer = .Tables(_tabla).Columns.Count

                        ReDim Preserve aTablaCompleta(20, 0)
                        ReDim Preserve aTieneBordeInferior(20, 0)

                        If .Tables(_tabla).Cell(1, 1).Range.Text.Trim.Length >= 6 AndAlso .Tables(_tabla).Cell(1, 1).Range.Text.Trim.ToLower.Substring(0, 6) = "oficio" Then
                            'Es el header

                            nColJurisdiccion = 0
                            nColMonto = 0
                            For _columna = 1 To cantColumnas
                                If .Tables(_tabla).Cell(1, _columna).Range.Text.Replace(ChrW(7), "").Trim.Length = 0 Then
                                    MsgBox("La tabla " & _tabla & " tiene el encabezado roto." & vbCrLf & "Corríjalo y vuelva a ejecutar el proceso.", vbCritical Or vbOKOnly)
                                    Exit Sub
                                End If
                                'Busca columnas clave
                                If .Tables(_tabla).Cell(1, _columna).Range.Text.Trim.Length >= 5 AndAlso .Tables(_tabla).Cell(1, _columna).Range.Text.Trim.ToLower.Substring(0, 5) = "monto" Then
                                    nColMonto = _columna
                                ElseIf .Tables(_tabla).Cell(1, _columna).Range.Text.Trim.Length >= 5 AndAlso .Tables(_tabla).Cell(1, _columna).Range.Text.Trim.ToLower.Substring(0, 5) = "juris" Then
                                    nColJurisdiccion = _columna
                                End If
                            Next

                            If _tabla = 1 Then
                                _begin = 1
                            Else
                                _begin = .Tables(_tabla - 1).Range.End
                            End If
                            _end = .Tables(_tabla).Range.Start

                            rngParagraphs = .Range(_begin, _end)
                            rngParagraphs.Select()

                            sw.WriteLine()
                            sw.WriteLine("******************************************************************************************************")
                            sw.WriteLine()
                            sw.Write(.Application.Selection.Text.Replace(vbCr, vbCrLf))
                            sw.WriteLine()
                            sw.WriteLine()

                        Else
                            Dim _stop As Boolean = False
                        End If

                        For _fila = 1 To .Tables(_tabla).Rows.Count

                            ReDim Preserve aTablaCompleta(20, aTablaCompleta.GetUpperBound(1) + 1)
                            ReDim Preserve aTieneBordeInferior(20, aTieneBordeInferior.GetUpperBound(1) + 1)

                            For _columna = 1 To cantColumnas
                                Try
                                    textoCelda = .Tables(_tabla).Cell(_fila, _columna).Range.Text

                                    If _columna = nColOficio Then       'oficio
                                        textoCelda = Trim(Regex.Replace(textoCelda.Replace(ChrW(7), " "), "\s+", "", RegexOptions.Compiled Or RegexOptions.Multiline)).Replace("- ", "")
                                    ElseIf _columna = nColExpte Then    'expediente
                                        textoCelda = Trim(Regex.Replace(textoCelda.Replace(ChrW(7), " "), "\s+", "", RegexOptions.Compiled Or RegexOptions.Multiline)).Replace("-", "")
                                    ElseIf _columna = nColMonto Then    'monto
                                        textoCelda = Trim(Regex.Replace(textoCelda.Replace(ChrW(7), ""), "\s+", "", RegexOptions.Compiled Or RegexOptions.Multiline))
                                    Else
                                        textoCelda = Regex.Replace(textoCelda, "\s+-\s+", "#-#")
                                        textoCelda = Trim(Regex.Replace(textoCelda.Replace(ChrW(7), " "), "\s+", " ", RegexOptions.Compiled Or RegexOptions.Multiline)).Replace("- ", "")
                                        textoCelda = Trim(textoCelda.Replace("#-#", " - "))
                                    End If

                                    aTablaCompleta(_columna, _fila) = textoCelda
                                    aTieneBordeInferior(_columna, _fila) = (.Tables(_tabla).Cell(_fila, _columna).Borders(-3).LineStyle <> 0)
                                Catch ex As Exception
                                    textoCelda = ""
                                    aTablaCompleta(_columna, _fila) = ""
                                    aTieneBordeInferior(_columna, _fila) = False
                                End Try
                            Next _columna

                        Next _fila

                        'como se han dado casos de tablas sin borde inferior en la ultima fila, lo fuerzo
                        For _columna = 1 To cantColumnas
                            aTieneBordeInferior(_columna, .Tables(_tabla).Rows.Count) = True
                        Next _columna

                        'acomodo el texto y fusiono lo que corresponda (teniendo en cuenta los bordes inferiores)
                        Dim _filaAux As Integer
                        For _fila = 1 To aTablaCompleta.GetUpperBound(1)
                            For _columna = 1 To cantColumnas
                                If aTablaCompleta(_columna, _fila) <> "" Then
                                    _filaAux = _fila
                                    Do While True

                                        Call ReemplazoCaracteresParaSAN(aTablaCompleta(_columna, _filaAux))

                                        If _fila <> _filaAux AndAlso aTablaCompleta(_columna, _filaAux) <> "" Then
                                            aTablaCompleta(_columna, _fila) &= " " & aTablaCompleta(_columna, _filaAux)
                                            aTablaCompleta(_columna, _fila) = Regex.Replace(aTablaCompleta(_columna, _fila), "([^ ])- ", "$1").Trim
                                            'inicializo
                                            aTablaCompleta(_columna, _filaAux) = ""
                                        End If

                                        If aTieneBordeInferior(_columna, _filaAux) Then Exit Do

                                        _filaAux += 1
                                    Loop

                                End If

                                'reemplazos especificos una vez que termine de concatenar
                                If _columna = nColMonto Then
                                    aTablaCompleta(_columna, _fila) = aTablaCompleta(_columna, _fila).Replace("+", " + ")
                                End If

                                'palabras puntuales con espacios en el medio
                                aTablaCompleta(_columna, _fila) = Regex.Replace(aTablaCompleta(_columna, _fila), "a\s*p\s*r\s*e\s*m\s*i\s*o", "APREMIO", RegexOptions.Compiled Or RegexOptions.IgnoreCase)
                                aTablaCompleta(_columna, _fila) = Regex.Replace(aTablaCompleta(_columna, _fila), "[uú]\s*n\s*i\s*c\s*a", "UNICA", RegexOptions.Compiled Or RegexOptions.IgnoreCase)
                                aTablaCompleta(_columna, _fila) = Regex.Replace(aTablaCompleta(_columna, _fila), "e\s*x\s*c\s*e\s*p\s*t\s*u\s*a\s*r", "EXCEPTUAR", RegexOptions.Compiled Or RegexOptions.IgnoreCase)
                                aTablaCompleta(_columna, _fila) = Regex.Replace(aTablaCompleta(_columna, _fila), "t\s*r\s*a\s*n\s*s\s*f\s*e\s*r\s*i\s*r", "TRANSFERIR", RegexOptions.Compiled Or RegexOptions.IgnoreCase)
                                aTablaCompleta(_columna, _fila) = Regex.Replace(aTablaCompleta(_columna, _fila), "e\s*j\s*e\s*c\s*u\s*c\s*i\s*[oó]\s*n", "EJECUCION", RegexOptions.Compiled Or RegexOptions.IgnoreCase)

                            Next _columna
                        Next _fila

                        ' vuelco los datos en el archivo de texto
                        judicEscribirSan(sw, aTablaCompleta, aTieneBordeInferior, cantColumnas)

                    Next _tabla

                    'ultimos parrafos

                    _begin = .Tables(.Tables.Count).Range.End
                    _end = .Range.End

                    rngParagraphs = .Range(_begin, _end)
                    rngParagraphs.Select()

                    sw.WriteLine()
                    sw.WriteLine("******************************************************************************************************")
                    sw.WriteLine()
                    sw.Write(.Application.Selection.Text.Replace(vbCr, vbCrLf))
                    sw.WriteLine()
                    sw.WriteLine()

                Else

                    '***************************
                    '* comunicaciones clásicas *
                    '***************************

                    _begin = 1
                    _end = .Range.End
                    rngParagraphs = .Range(_begin, _end)
                    rngParagraphs.Select()

                    Dim texto As String = .Application.Selection.Text.Replace(vbCr, vbCrLf)

                    Call ReemplazoCaracteresParaSAN(texto)

                    sw.Write(texto)

                End If

            Catch ex As Exception
                MsgBox("Se produjo un error al generar el .SAN, revise el texto generado!!!" & vbCrLf & "ERROR:  " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical)
                Exit Sub
            Finally
                sw.Flush()
                sw.Close()
            End Try

            MsgBox("Documento guardado en:" & vbCrLf & "'" & archivoSalida & "'", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)

        End With

    End Sub

    'Graba el .SAN
    Private Sub judicEscribirSan(ByRef sw As StreamWriter, ByVal aTablaCompleta(,) As String, ByVal aTieneBordeInferior(,) As Boolean, ByVal cantColumnas As Integer)

        Dim _fila, _columna As Integer
        Dim auxTexto As String

        'recorro el array nuevamente y escribo solo las filas que al menos tiene un dato
        For _fila = 1 To aTablaCompleta.GetUpperBound(1)
            Dim imprimir As Boolean = False
            For _columna = 1 To cantColumnas
                If aTablaCompleta(_columna, _fila).Length > 0 Then
                    imprimir = True
                    Exit For
                End If
            Next _columna
            If imprimir Then
                For _columna = 1 To cantColumnas
                    auxTexto = aTablaCompleta(_columna, _fila)
                    Call ReemplazoCaracteresParaSAN(auxTexto)
                    sw.WriteLine(auxTexto)
                    sw.Flush()
                Next _columna
            End If
        Next _fila

        sw.Flush()

    End Sub

    Private Sub ReemplazoCaracteresParaSAN(ByRef texto As String)

        texto = texto.Replace("–", "-")

        texto = texto.Replace(ChrW(147), """") '“
        texto = texto.Replace(ChrW(148), """") '”
        texto = texto.Replace(Chr(145), "'") '‘
        texto = texto.Replace(Chr(146), "'") '’

        texto = texto.Replace("…", "...")
        texto = texto.Replace(Chr(&H85), "...") '…
        texto = texto.Replace(ChrW(&H2026), "...") '…

        texto = texto.Replace(ChrW(&H201C), """") '“
        texto = texto.Replace(ChrW(&H201D), """") '”
        texto = texto.Replace(ChrW(&H2018), """") '‘
        texto = texto.Replace(ChrW(&H2019), """") '’

        texto = texto.Replace(ChrW(&HBC), "1/4") '¼
        texto = texto.Replace(ChrW(&HBD), "1/2") '½
        texto = texto.Replace(ChrW(&HBE), "3/4") '¾

    End Sub

#End Region

#End Region

#Region "Métodos"

#Region "Patrones"

    'Reemplaza el patrón IEF de la 1er página
    Private Sub PatronIEF()
        Dim rg As Word.Range
        Dim texto As String

        With Globals.ThisAddIn.Application.ActiveDocument
            'Validación de ejecución previa
            .InlineShapes.Item(1).Select()
            If .Application.Selection.Start > 500 Then
                If US.Display.MsgBox("Al parecer ya ejecutó este patrón en el presente documento. ¿Está seguro que desea volver a aplicarlo?", "Confirmación Requerida", US.Display.MsgBoxTipos.msgConfirmacion) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If
            End If

            '--------------------------------------------
            '1. Sangría primera oración (cab.página)
            '--------------------------------------------
            .Paragraphs(1).LeftIndent = CentimetersToPoints(8)

            '--------------------------------------------
            '2. Reemplaza el primer inLineShape
            '--------------------------------------------
            Me.ReemplazarImagenBCRA(1, False, False, False)

            '--------------------------------------------
            '3. Agrega una tabla para el encabezado
            '--------------------------------------------
            Me.ReemplazarShapeXtabla(1)
            '...configura la tabla
            With .Tables(1)
                Try
                    .Rows.Height = CentimetersToPoints(0.83)
                    .Columns(1).Width = CentimetersToPoints(14.37)
                    .Columns(2).Width = CentimetersToPoints(2.75)
                    .Rows(1).Alignment = Word.WdRowAlignment.wdAlignRowCenter
                Catch ex As Exception
                    'No pudo dar formato a la tabla porque vino con otra estructura
                End Try
            End With
            '...elimina el shape
            .Shapes(1).Delete()
            '...agrega el texto
            For s As Integer = 1 To 20
                rg = .Sentences(s)
                rg.Select()
                If .Application.Selection.Text.IndexOf("COMUNICA") <> -1 Then
                    'Encontró el texto buscado
                    texto = .Application.Selection.Text.Trim
                    'Elimina la oración
                    .Application.Selection.Delete()
                    'Agrega el texto a la tabla
                    rg = .Tables(1).Cell(1, 1).Range
                    rg.Text = texto.Substring(0, texto.Length - 10)
                    rg = .Tables(1).Cell(1, 2).Range
                    rg.Text = Microsoft.VisualBasic.Right(texto, 10)
                    'Cancela la búsqueda
                    Exit For
                End If
            Next


            '--------------------------------------------
            '4. Elimina el primer shape (subrayado)
            '--------------------------------------------
            Try
                .Shapes.Item(1).Delete()
            Catch ex As Exception
                'no siempre está este shape
            End Try

            '--------------------------------------------
            '5. Coloca las firmas del doc. en una tabla
            '--------------------------------------------
            Me.FirmasEnTabla()

        End With

    End Sub

    'Reemplaza el texto seleccionado por un patrón con un cuadro
    'Si cantParrafos=0 ==> asume que no se puede identificar (toma todos los de la selección, sin controlarlos)
    Private Sub PatronCuadro(ByVal cantParrafos As Integer,
                             ByVal cantFilas As Integer,
                             ByVal cantColumnas As Integer,
                             ByVal tipoCuadro As TiposCuadro)


        'Reemplaza los TABS por ENTERS para diferenciar los párrafos en la selección
        With Globals.ThisAddIn.Application.ActiveDocument.Application.Selection
            .Find.Execute(FindText:="^t", ReplaceWith:="*t", Replace:=Word.WdReplace.wdReplaceAll)
            .Find.Execute(FindText:="*t", ReplaceWith:="^p", Replace:=Word.WdReplace.wdReplaceAll)
        End With

        'Formatea a una columna
        Me.FormatearColumnas(Globals.ThisAddIn.Application.ActiveDocument.Application.Selection)

        'Variables
        Dim oraciones As New ArrayList()

        With Globals.ThisAddIn.Application.ActiveDocument.Application.Selection
            'Recorre los párrafos seleccionados y verifica que tenga lo necesario para el patrón elegido
            oraciones.Clear()
            For i As Integer = 1 To .Paragraphs.Count
                'Verifica si el párrafo tiene contenido
                If .Paragraphs(i).Range.Text.Trim <> String.Empty Then
                    'Almacena el mismo en una matriz
                    oraciones.Add(.Paragraphs(i).Range.Text.Trim)
                End If
            Next

            'Si hay menos de los párrafos necesarios, amplía la selección
            Do While cantParrafos <> 0 And oraciones.Count < cantParrafos
                For i As Integer = oraciones.Count To cantParrafos
                    'Agrega un párrafo
                    .MoveEnd(Unit:=Word.WdUnits.wdParagraph, Count:=1)
                    If .Paragraphs(.Paragraphs.Count).Range.Text.Trim <> String.Empty Then
                        'Almacena el último párrafo en la matriz (si no está vacío)
                        oraciones.Add(.Paragraphs(.Paragraphs.Count).Range.Text.Trim)
                    End If
                Next
            Loop

            'Si la cantidad de párrafos no coincide con los esperados, informa
            If cantParrafos <> 0 And cantParrafos <> oraciones.Count Then
                MsgBox("La cantidad de párrafos detectados no coincide con los esperados.", MsgBoxStyle.Exclamation, Header)
                Exit Sub
            End If

            'Crea una tabla con el formato deseado
            InsertarTablaCuadro(Globals.ThisAddIn.Application.ActiveDocument.Application.Selection,
                cantFilas, cantColumnas, oraciones, tipoCuadro)
        End With
    End Sub

    'Transforma el párrafo en una columna y unifica el rdo. en un único párrafo
    Private Sub PatronCol1(ByVal seleccion As Microsoft.Office.Interop.Word.Selection)
        'Formatea a una columna
        Me.FormatearColumnas(seleccion)

        'Unifica los párrafos
        Dim oraciones As New ArrayList()
        Dim texto As String = ""

        With seleccion
            'Recorre los párrafos seleccionados
            For i As Integer = 1 To .Paragraphs.Count
                'Verifica si el párrafo tiene contenido
                If .Paragraphs(i).Range.Text.Trim <> String.Empty Then
                    'Almacena el mismo en una matriz
                    oraciones.Add(.Paragraphs(i).Range.Text.Trim)
                End If
            Next

            'Reacomoda las oraciones
            For i As Integer = oraciones.Count - 1 To 0 Step -1
                texto &= oraciones(i)
                'quita el último guión si existe o agrega un espacio
                If Right(texto, 1) = "-" Then
                    texto = texto.Substring(0, texto.Length - 1)
                Else
                    texto &= " "
                End If
            Next
            texto = texto.Trim
            'Agrega un ENTER final
            texto &= vbCrLf

            'Modifica la selección por las oraciones reacomodadas
            .Delete()
            .InsertBefore(texto)
            .Font.Underline = Word.WdUnderline.wdUnderlineNone

            'Modifica el margen izq. del primer renglón
            .ParagraphFormat.FirstLineIndent = CentimetersToPoints(2.5)
        End With
    End Sub

    'Transforma el párrafo en una columna y coloca el texto de la columna 2 adelante
    Private Sub PatronCol2(ByVal seleccion As Microsoft.Office.Interop.Word.Selection)
        'Formatea a una columna
        Me.FormatearColumnas(seleccion)

        'Unifica los párrafos
        Dim oraciones As New ArrayList()
        Dim texto As String = ""

        With seleccion
            'Recorre los párrafos seleccionados
            For i As Integer = 1 To .Paragraphs.Count
                'Verifica si el párrafo tiene contenido
                If .Paragraphs(i).Range.Text.Trim <> String.Empty Then
                    'Almacena el mismo en una matriz
                    oraciones.Add(.Paragraphs(i).Range.Text.Trim)
                End If
            Next

            'Reacomoda las oraciones (la última primero)
            texto = vbCrLf
            texto &= oraciones(oraciones.Count - 1)
            texto &= vbCrLf

            'Elimina el texto de la última oración y lo coloca al principio
            For i As Integer = .Sentences.Count To 1 Step -1
                If .Sentences(i).Text.Trim <> String.Empty Then
                    .Sentences(i).Delete()
                    Exit For
                End If
            Next
            .InsertBefore(texto)
        End With
    End Sub

    'Busca el párrafo que comienza con "Ref.:   " para darle el formato adecuado
    Private Sub PatronREF()
        Dim rg As Word.Range

        With Globals.ThisAddIn.Application.ActiveDocument
            'busca el inicio de las referencias
            For s As Integer = 1 To .Sentences.Count
                rg = .Sentences(s)
                rg.Select()
                If .Application.Selection.Text.IndexOf("Ref.: ") <> -1 Then
                    'Encontró el inicio de las referencias

                    '==> Formatea el texto seleccionado (1º oración)
                    Me.FormatearSeleccion(.Application.Selection, True, False, Word.WdParagraphAlignment.wdAlignParagraphJustify)

                    '==> Formatea el siguiente renglón (2º oración)
                    rg = .Sentences(s + 2)
                    rg.End = rg.Start + 1
                    rg.Select()
                    Me.FormatearSeleccion(.Application.Selection, True, False, Word.WdParagraphAlignment.wdAlignParagraphJustify)

                    '==> Formatea el texto en negrita (3º oración)
                    rg = .Sentences(s + 3)
                    rg.End = rg.Start + 1
                    rg.Select()
                    Me.FormatearSeleccion(.Application.Selection, True, False, Word.WdParagraphAlignment.wdAlignParagraphJustify)

                    'Termina el bucle pq encontró el texto deseado
                    Exit For
                End If
                'Condición de corte (overflow)
                If s > 100 Then
                    MsgBox("No se encontró el patrón solicitado.", MsgBoxStyle.Exclamation, Header)
                    Exit Sub
                End If
            Next
        End With

    End Sub

    'Formatea el Boletín Oficial
    Private Sub FormatearBO(ByVal tipoProceso As TipoProcBO)
        'Realiza las búsquedas y reemplazos en el texto seleccionado
        With Globals.ThisAddIn.Application.ActiveDocument
            Dim rg As Word.Range
            Dim OracionInicio As Integer = 0
            Dim OracionFinal As Integer = 0

            If tipoProceso = TipoProcBO.Específico Then
                'Realiza un primer formateo específico del encabezado
                FormatearBO_encabezado(OracionInicio, OracionFinal)
            End If

            '----------------------------------------------------------
            'Selecciona todo el documento y realiza ajustes generales
            .Select()

            '...Reemplaza Enter de salto de línea por Enter de salto de párrafo
            ReemplazarEnSeleccion(.Application.Selection, "^l", "^p")
            '...Cambia ".^p" por ".^p^p"
            ReemplazarEnSeleccion(.Application.Selection, ".^p", ".^p^p")
            '...Cambia ";^p" por ";^p^p"
            ReemplazarEnSeleccion(.Application.Selection, ";^p", ";^p^p")
            '...Cambia ":^p" por ":^p^p"
            ReemplazarEnSeleccion(.Application.Selection, ":^p", ":^p^p")
            '...Elimina los Tabs
            ReemplazarEnSeleccion(.Application.Selection, "^t", "")

            '...Reemplazo avanzado de ENTERS
            ReemplazarEnSeleccion(.Application.Selection, ".^p^p", "#punto_enter#")
            ReemplazarEnSeleccion(.Application.Selection, ":^p^p", "#dospunto_enter#")
            ReemplazarEnSeleccion(.Application.Selection, ";^p^p", "#puntoycoma_enter#")
            ReemplazarEnSeleccion(.Application.Selection, "^p^p", "#dosenters#")

            ReemplazarEnSeleccion(.Application.Selection, "^pArtículo", "#articulo_enter#")
            ReemplazarEnSeleccion(.Application.Selection, "^pArt.", "#art_enter#")
            ReemplazarEnSeleccion(.Application.Selection, "^pSección", "#seccion_enter#")
            ReemplazarEnSeleccion(.Application.Selection, "^pSecc.", "#secc_enter#")
            ReemplazarEnSeleccion(.Application.Selection, "^pCapítulo", "#capitulo_enter#")
            ReemplazarEnSeleccion(.Application.Selection, "^pCap.", "#cap_enter#")
            ReemplazarEnSeleccion(.Application.Selection, "^pAnexo", "#anexo_enter#")

            ReemplazarEnSeleccion(.Application.Selection, "^pI.", "#enterI#")
            ReemplazarEnSeleccion(.Application.Selection, "^pII.", "#enterII#")
            ReemplazarEnSeleccion(.Application.Selection, "^pIII.", "#enterIII#")
            ReemplazarEnSeleccion(.Application.Selection, "^pIV.", "#enterIV#")
            ReemplazarEnSeleccion(.Application.Selection, "^pV.", "#enterV#")
            ReemplazarEnSeleccion(.Application.Selection, "^pVI.", "#enterVI#")
            ReemplazarEnSeleccion(.Application.Selection, "^pVII.", "#enterVII#")
            ReemplazarEnSeleccion(.Application.Selection, "^pVIII.", "#enterVIII#")
            ReemplazarEnSeleccion(.Application.Selection, "^pIX.", "#enterIX#")
            ReemplazarEnSeleccion(.Application.Selection, "^pX.", "#enterX#")
            ReemplazarEnSeleccion(.Application.Selection, "^p1.", "#enter1#")
            ReemplazarEnSeleccion(.Application.Selection, "^p2.", "#enter2#")
            ReemplazarEnSeleccion(.Application.Selection, "^p3.", "#enter3#")
            ReemplazarEnSeleccion(.Application.Selection, "^p4.", "#enter4#")
            ReemplazarEnSeleccion(.Application.Selection, "^p5.", "#enter5#")
            ReemplazarEnSeleccion(.Application.Selection, "^p6.", "#enter6#")
            ReemplazarEnSeleccion(.Application.Selection, "^p7.", "#enter7#")
            ReemplazarEnSeleccion(.Application.Selection, "^p8.", "#enter8#")
            ReemplazarEnSeleccion(.Application.Selection, "^p9.", "#enter9#")
            ReemplazarEnSeleccion(.Application.Selection, "^p10.", "#enter10#")

            ReemplazarEnSeleccion(.Application.Selection, "-^p", "")
            ReemplazarEnSeleccion(.Application.Selection, "^p", " ")

            ReemplazarEnSeleccion(.Application.Selection, "#enterI#", "^p^pI.")
            ReemplazarEnSeleccion(.Application.Selection, "#enterII#", "^p^pII.")
            ReemplazarEnSeleccion(.Application.Selection, "#enterIII#", "^p^pIII.")
            ReemplazarEnSeleccion(.Application.Selection, "#enterIV#", "^p^pIV.")
            ReemplazarEnSeleccion(.Application.Selection, "#enterV#", "^p^pV.")
            ReemplazarEnSeleccion(.Application.Selection, "#enterVI#", "^p^pVI.")
            ReemplazarEnSeleccion(.Application.Selection, "#enterVII#", "^p^pVII.")
            ReemplazarEnSeleccion(.Application.Selection, "#enterVIII#", "^p^pVIII.")
            ReemplazarEnSeleccion(.Application.Selection, "#enterIX#", "^p^pIX.")
            ReemplazarEnSeleccion(.Application.Selection, "#enterX#", "^p^pX.")
            ReemplazarEnSeleccion(.Application.Selection, "#enter1#", "^p^p1.")
            ReemplazarEnSeleccion(.Application.Selection, "#enter2#", "^p^p2.")
            ReemplazarEnSeleccion(.Application.Selection, "#enter3#", "^p^p3.")
            ReemplazarEnSeleccion(.Application.Selection, "#enter4#", "^p^p4.")
            ReemplazarEnSeleccion(.Application.Selection, "#enter5#", "^p^p5.")
            ReemplazarEnSeleccion(.Application.Selection, "#enter6#", "^p^p6.")
            ReemplazarEnSeleccion(.Application.Selection, "#enter7#", "^p^p7.")
            ReemplazarEnSeleccion(.Application.Selection, "#enter8#", "^p^p8.")
            ReemplazarEnSeleccion(.Application.Selection, "#enter9#", "^p^p9.")
            ReemplazarEnSeleccion(.Application.Selection, "#enter10#", "^p^p10.")

            ReemplazarEnSeleccion(.Application.Selection, "#anexo_enter#", "^p^pAnexo")
            ReemplazarEnSeleccion(.Application.Selection, "#cap_enter#", "^p^pCap.")
            ReemplazarEnSeleccion(.Application.Selection, "#capitulo_enter#", "^p^pCapítulo")
            ReemplazarEnSeleccion(.Application.Selection, "#secc_enter#", "^p^pSecc.")
            ReemplazarEnSeleccion(.Application.Selection, "#seccion_enter#", "^p^pSección")
            ReemplazarEnSeleccion(.Application.Selection, "#art_enter#", "^p^pArt.")
            ReemplazarEnSeleccion(.Application.Selection, "#articulo_enter#", "^p^pArtículo")

            ReemplazarEnSeleccion(.Application.Selection, "#punto_enter#", ".^p^p")
            ReemplazarEnSeleccion(.Application.Selection, "#dospunto_enter#", ":^p^p")
            ReemplazarEnSeleccion(.Application.Selection, "#puntoycoma_enter#", ";^p^p")
            ReemplazarEnSeleccion(.Application.Selection, "#dosenters#", "^p^p")

            With .Application.Selection
                '...Formato del párrafo
                .Font.Name = "Arial"
                .Font.Size = 10
                With .ParagraphFormat
                    .Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                    .LeftIndent = 0
                    .RightIndent = 0
                    .FirstLineIndent = 0
                End With
            End With

            '...Agrega un título
            Dim titulo As New StringBuilder
            titulo.AppendFormat("Fecha de Publicación en el Boletín Oficial Nº {0}: {1}",
                                nu_publicacion.Text.Trim, Format(Now, "d/M/yyyy"))
            With .Application.Selection
                .InsertParagraphBefore()
                .InsertBefore(vbCrLf)
                .InsertBefore(titulo.ToString)

                Dim inicio As Integer = .Start
                .End = inicio + titulo.ToString.Length
                .Font.Bold = True
                .Font.Underline = Word.WdUnderline.wdUnderlineSingle
                .ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                .ParagraphFormat.FirstLineIndent = 0
            End With


            '----------------------------------------------------------
            If tipoProceso = TipoProcBO.Específico Then
                'Realiza un primer formateo específico del final
                FormatearBO_final(OracionInicio, OracionFinal)
            End If


            '----------------------------------------------------------
            If tipoProceso = TipoProcBO.Específico Then
                'Selecciona para ajustar sangría (desde "RESUELVE:..." hasta el final del documento)
                .Range(Start:=0, End:=0).Select()

                '...busca el inicio del texto a seleccionar
                rg = .Sentences(OracionFinal)
                rg.Select()
                For s As Integer = OracionFinal + 1 To .Sentences.Count
                    rg.MoveEnd(Unit:=Word.WdUnits.wdSentence, Count:=1)
                    rg.Select()
                Next
            Else
                'Selecciona todo el documento
                .Select()
            End If

            With .Application.Selection
                .ParagraphFormat.LeftIndent = 0
                .ParagraphFormat.FirstLineIndent = CentimetersToPoints(0.5)
            End With


        End With

    End Sub

    'Formatea el encabezado del Boletín Oficial
    Private Sub FormatearBO_encabezado(ByRef OracionInicio As Integer, ByRef OracionFinal As Integer)
        'Realiza las búsquedas y reemplazos en el texto seleccionado
        With Globals.ThisAddIn.Application.ActiveDocument
            Dim rg As Word.Range
            OracionInicio = 0
            OracionFinal = 0
            '----------------------------------------------------------
            'Formatea el encabezado(del inicio hasta "considerando")
            .Range(Start:=0, End:=0).Select()

            '...busca el final del encabezado
            rg = .Sentences(1)
            rg.Select()
            For s As Integer = 2 To .Sentences.Count
                rg.MoveEnd(Unit:=Word.WdUnits.wdSentence, Count:=1)
                rg.Select()
                If .Application.Selection.Text.ToUpper.IndexOf("CONSIDERANDO:") <> -1 Then
                    'encontró la última línea del encabezado
                    Exit For
                End If
            Next

            '...en la selección, realiza ciertos reemplazos separando oraciones
            ReemplazarEnSeleccion(.Application.Selection, "^pResolución", "^p^pResolución", True)
            ReemplazarEnSeleccion(.Application.Selection, "^pDecisión", "^p^pDecisión", True)
            ReemplazarEnSeleccion(.Application.Selection, "^pInstrucción", "^p^pInstrucción", True)
            ReemplazarEnSeleccion(.Application.Selection, "^pAcordada", "^p^pAcordada", False)
            ReemplazarEnSeleccion(.Application.Selection, "^pDisposición", "^p^pDisposición", True)
            ReemplazarEnSeleccion(.Application.Selection, "^pDictamen", "^p^pDictamen", False)
            ReemplazarEnSeleccion(.Application.Selection, "^pNota Externa", "^p^pNota Externa", False)
            ReemplazarEnSeleccion(.Application.Selection, "^pComunicado", "^p^pComunicado", False)
            ReemplazarEnSeleccion(.Application.Selection, "^pDecreto", "^p^pDecreto", False)

            ReemplazarEnSeleccion(.Application.Selection, "0^p", "0^p^p")
            ReemplazarEnSeleccion(.Application.Selection, "1^p", "1^p^p")
            ReemplazarEnSeleccion(.Application.Selection, "2^p", "2^p^p")
            ReemplazarEnSeleccion(.Application.Selection, "3^p", "3^p^p")
            ReemplazarEnSeleccion(.Application.Selection, "4^p", "4^p^p")
            ReemplazarEnSeleccion(.Application.Selection, "5^p", "5^p^p")
            ReemplazarEnSeleccion(.Application.Selection, "6^p", "6^p^p")
            ReemplazarEnSeleccion(.Application.Selection, "7^p", "7^p^p")
            ReemplazarEnSeleccion(.Application.Selection, "8^p", "8^p^p")
            ReemplazarEnSeleccion(.Application.Selection, "9^p", "9^p^p")

            ReemplazarEnSeleccion(.Application.Selection, "^pConsiderando", "^p^pConsiderando")


            '----------------------------------------------------------
            'Busca una sección clave del texto para realizar reemplazos especiales
            '(desde "POR ELLO" hasta "RESUELVE...")
            .Range(Start:=0, End:=0).Select()

            '...busca el inicio del texto a seleccionar
            rg = .Sentences(1)
            rg.Select()
            For s As Integer = 2 To .Sentences.Count
                rg.MoveEnd(Unit:=Word.WdUnits.wdSentence, Count:=1)
                rg.Select()
                If .Application.Selection.Text.ToUpper.IndexOf("POR ELLO,") <> -1 Then
                    'encontró el inicio real del texto
                    OracionInicio = s
                    Exit For
                End If
            Next
            If OracionInicio <> 0 Then
                rg = .Sentences(OracionInicio)
                rg.Select()
                For s As Integer = OracionInicio + 1 To .Sentences.Count
                    rg.MoveEnd(Unit:=Word.WdUnits.wdSentence, Count:=1)
                    rg.Select()
                    With .Application.Selection.Text.ToUpper
                        If .IndexOf("DECIDE:") <> -1 Or .IndexOf("DECRETA:") <> -1 _
                            Or .IndexOf("RESUELVE:") <> -1 Or .IndexOf("INSTRUYE:") <> -1 _
                            Or .IndexOf("DICTAMINA:") <> -1 Or .IndexOf("ACORDARON:") <> -1 _
                            Or .IndexOf("ACUERDAN:") <> -1 Or .IndexOf("DISPONE:") <> -1 Then
                            'encontró el final del texto
                            Exit For
                        End If
                    End With
                Next
            End If

            '...Separa oraciones
            ReemplazarEnSeleccion(.Application.Selection, "^p", "^p^p")
        End With
    End Sub

    'Formatea el final del Boletín Oficial
    Private Sub FormatearBO_final(ByRef OracionInicio As Integer, ByRef OracionFinal As Integer)
        'Realiza las búsquedas y reemplazos en el texto seleccionado
        With Globals.ThisAddIn.Application.ActiveDocument
            Dim rg As Word.Range
            OracionInicio = 0
            OracionFinal = 0

            '----------------------------------------------------------
            'Selecciona para ajustar sangría (párrafo "VISTO...")
            .Range(Start:=0, End:=0).Select()

            '...busca el párrafo deseado
            rg = .Sentences(1)
            rg.Select()
            For s As Integer = 2 To .Sentences.Count
                rg.MoveEnd(Unit:=Word.WdUnits.wdSentence, Count:=1)
                rg.Select()
                If .Application.Selection.Text.ToUpper.IndexOf("VISTO") <> -1 Then
                    'encontró el párrafo deseado => lo selecciona
                    rg = .Sentences(s)
                    rg.Select()
                    Exit For
                End If
            Next
            With .Application.Selection
                .ParagraphFormat.LeftIndent = CentimetersToPoints(0.5)
                .ParagraphFormat.FirstLineIndent = CentimetersToPoints(0.5) * -1
            End With


            '----------------------------------------------------------
            'Selecciona para ajustar sangría (desde "CONSIDERANDO" hasta antes de "POR ELLO,")
            .Range(Start:=0, End:=0).Select()

            '...busca el inicio del texto a seleccionar
            OracionInicio = 0
            rg = .Sentences(1)
            rg.Select()
            For s As Integer = 2 To .Sentences.Count
                rg.MoveEnd(Unit:=Word.WdUnits.wdSentence, Count:=1)
                rg.Select()
                If .Application.Selection.Text.ToUpper.IndexOf("CONSIDERANDO:") <> -1 Then
                    'encontró el inicio real del texto
                    OracionInicio = s
                    Exit For
                End If
            Next
            If OracionInicio <> 0 Then
                rg = .Sentences(OracionInicio)
                rg.Select()
                For s As Integer = OracionInicio + 1 To .Sentences.Count
                    rg.MoveEnd(Unit:=Word.WdUnits.wdSentence, Count:=1)
                    rg.Select()
                    If .Application.Selection.Text.ToUpper.IndexOf("POR ELLO,:") <> -1 Then
                        'encontró el final del texto
                        rg.MoveEnd(Unit:=Word.WdUnits.wdSentence, Count:=-1)
                        rg.Select()
                        Exit For
                    End If
                Next
            End If
            With .Application.Selection
                .ParagraphFormat.LeftIndent = CentimetersToPoints(0.5)
            End With


            '----------------------------------------------------------
            'Selecciona para ajustar sangría (desde después de "POR ELLO," hasta "  RESUELVE:...")
            .Range(Start:=0, End:=0).Select()

            '...busca el inicio del texto a seleccionar
            OracionInicio = 0
            rg = .Sentences(1)
            rg.Select()
            For s As Integer = 2 To .Sentences.Count
                rg.MoveEnd(Unit:=Word.WdUnits.wdSentence, Count:=1)
                rg.Select()
                If .Application.Selection.Text.ToUpper.IndexOf("POR ELLO,") <> -1 Then
                    'encontró el inicio real del texto
                    OracionInicio = s + 1
                    Exit For
                End If
            Next
            If OracionInicio <> 0 Then
                rg = .Sentences(OracionInicio)
                rg.Select()
                For s As Integer = OracionInicio + 1 To .Sentences.Count
                    rg.MoveEnd(Unit:=Word.WdUnits.wdSentence, Count:=1)
                    rg.Select()
                    With .Application.Selection.Text.ToUpper
                        If .IndexOf("DECIDE:") <> -1 Or .IndexOf("DECRETA:") <> -1 _
                            Or .IndexOf("RESUELVE:") <> -1 Or .IndexOf("INSTRUYE:") <> -1 _
                            Or .IndexOf("DICTAMINA:") <> -1 Or .IndexOf("ACORDARON:") <> -1 _
                            Or .IndexOf("ACUERDAN:") <> -1 Or .IndexOf("DISPONE:") <> -1 Then
                            'encontró el final del texto
                            OracionFinal = s + 1
                            rg.Select()
                            Exit For
                        End If
                    End With
                Next
            End If
            With .Application.Selection
                .ParagraphFormat.FirstLineIndent = 0
                .ParagraphFormat.LeftIndent = 0
            End With
            'Cambia "^p^p" por "^p"
            ReemplazarEnSeleccion(.Application.Selection, "^p^p", "^p")
        End With
    End Sub

#End Region

#Region "Adelgazar"

    'Adelgaza el documento en formato RTF
    Private Sub AdelgazarRtf(ByVal sArchivoRtf As String, ByVal sArchRtfFinal As String)

        Dim sRtf As String                      'Archivo RTF para trabajarlo en memoria
        Dim oMbaTern As Mba.MbaTern = Nothing   'TEedit

        Try
            '******************
            '      RTF !!!
            '******************
            'levanto el rtf seleccionado a memoria
            sRtf = Archivos.LeerArchivoTexto(sArchivoRtf)

            'elimino los tags molestos
            EliminoTagsConValores(sRtf, "expnd")
            EliminoTagsConValores(sRtf, "expndtw")

            'grabo el resultado a disco
            sArchivoRtf = sArchivoRtf.Substring(0, sArchivoRtf.LastIndexOf(".")) & "_adelgazado.rtf"
            Archivos.GrabarArchivoTexto(sArchivoRtf, sRtf)


            '******************
            '     TEDIT !!!
            '******************
            'creo el objeto teedit
            oMbaTern = New MbaTern

            'levanto el RTF en la TEdit
            oMbaTern.MbaLoadRtf(sArchivoRtf)

            'grabo el RTF a disco
            oMbaTern.MbaSaveRtf(sArchRtfFinal, False, True)

            'cierro el objeto teedit
            oMbaTern = Nothing

        Catch ex As Exception
            Throw ex

        Finally
            If Not oMbaTern Is Nothing Then
                oMbaTern.Dispose()
                oMbaTern = Nothing
            End If

        End Try

    End Sub

    'Elimina TAGs del RTF
    Private Sub EliminoTagsConValores(ByRef sRtf As String, ByVal sTag As String)

        ReemplazoPatron(sRtf, "(\D)\\" & sTag & "[-]?[\d]+\s", "$1")
        ReemplazoPatron(sRtf, "(\d)\\" & sTag & "[-]?[\d]+\s(\D)", "$1$2")
        ReemplazoPatron(sRtf, "(\\[a-z]+[\d]+)\\" & sTag & "[-]?[\d]+\s(\d)", "$1 $2")
        ReemplazoPatron(sRtf, "(\d)\\" & sTag & "[-]?[\d]+\s", "$1")
        ReemplazoPatron(sRtf, "\\" & sTag & "[-]?[\d]+", "")

    End Sub

    'Ejecuta una expresión regular de reemplazo
    Private Sub ReemplazoPatron(ByRef sRtf As String, ByVal pattern As String, ByVal reemplazo As String)

        Dim _regEx As Regex

        _regEx = New Regex(pattern, RegexOptions.IgnoreCase Or RegexOptions.Multiline Or RegexOptions.Compiled)
        Do While _regEx.IsMatch(sRtf)
            sRtf = _regEx.Replace(sRtf, reemplazo)
        Loop
        _regEx = Nothing

    End Sub

#End Region

#Region "Auxiliares"

    'Devuelve los points equivalentes a los cm pasados por parámetro
    Private Function CentimetersToPoints(ByVal centimetros As Single) As Single
        '1 cm = 28.35 points
        Return centimetros * 28.35
    End Function

    'Determina la cantidad de filas y columnas del shape
    Private Sub DeterminarFormatoShape(ByVal oShape As Word.Shape, ByRef lineas As Integer, ByRef columnas As Integer)
        lineas = -1 : columnas = -1
        Try
            'Recorre cada línea del shape
            For l As Integer = 1 To oShape.GroupItems.Count
                With oShape.GroupItems(l)
                    If .Height > 0 And .Width <= 1 Then
                        'línea vertical
                        columnas += 1
                    ElseIf .Height <= 1 And .Width > 0 Then
                        'línea horizontal
                        lineas += 1
                    Else
                        'Es una línea diagonal => la ignora
                    End If
                End With
            Next

            'Verifica tamaño mínimo
            If lineas < 1 Then lineas = 1
            If columnas < 1 Then columnas = 1

        Catch ex As Exception
            'Formato estándar de tabla
            lineas = 1
            columnas = 1
        End Try
    End Sub

    'Aplica un estandar de formato a la selección pasada por parámetro
    Private Sub FormatearSeleccion(ByRef seleccion As Word.Selection, Optional ByVal MantenerMargen As Boolean = True,
                                   Optional ByVal MantenerAlineacion As Boolean = True,
                                   Optional ByVal Alineacion As Word.WdParagraphAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft)
        Dim margenIzq As Single = 0

        With seleccion
            If MantenerMargen Then
                margenIzq = .Paragraphs.LeftIndent
            End If
            If MantenerAlineacion Then
                Alineacion = .ParagraphFormat.Alignment
            End If

            'borra el formato actual
            .ClearParagraphAllFormatting()
            .ParagraphFormat.TabStops.ClearAll()

            'aplica un nuevo formato
            .ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle
            .ParagraphFormat.SpaceAfter = 0
            .ParagraphFormat.LeftIndent = margenIzq
            .ParagraphFormat.LineUnitAfter = 0
            .ParagraphFormat.LineUnitBefore = 0
            .ParagraphFormat.Alignment = Alineacion
            .Font.Size = 10
            .Font.Position = 0

            'elimina objetos en el texto
            .Find.Execute(FindText:="^t", ReplaceWith:=" ", Replace:=Word.WdReplace.wdReplaceAll)
        End With
    End Sub

    'Aplica formato de una columna a la selección pasada por parámetro
    Private Sub FormatearColumnas(ByRef seleccion As Word.Selection)
        With seleccion
            'elimina los saltos de columna
            .Find.Execute(FindText:="^n", ReplaceWith:=" ", Replace:=Word.WdReplace.wdReplaceAll)
            'elimina los saltos de sección
            .Find.Execute(FindText:="^b", ReplaceWith:=" ", Replace:=Word.WdReplace.wdReplaceAll)
            'aplica formato de párrafo estándar
            'FormatearSeleccion(seleccion)
        End With
    End Sub

    'Reemplaza en la selección el textoBuscado por el textoNuevo
    Private Sub ReemplazarEnSeleccion(ByRef seleccion As Word.Selection,
                                      ByVal textoBuscado As String, ByVal textoNuevo As String,
                                      Optional ByVal ignorarAcentos As Boolean = False)
        With seleccion
            With .Find
                .ClearFormatting()
                .Text = textoBuscado
                .MatchCase = False      'ignora may/min
                .MatchWholeWord = True  'palabras completas
                .Forward = True
                .Replacement.ClearFormatting()
                .Replacement.Text = textoNuevo
                .Execute(Format:=False, Replace:=Word.WdReplace.wdReplaceAll)

                '...nueva búsqueda sin acentos (si la búsqueda con acentos no dio resultados)
                If ignorarAcentos And Not .Found Then
                    textoBuscado = textoBuscado.Replace("á", "a")
                    textoBuscado = textoBuscado.Replace("é", "e")
                    textoBuscado = textoBuscado.Replace("í", "i")
                    textoBuscado = textoBuscado.Replace("ó", "o")
                    textoBuscado = textoBuscado.Replace("ú", "u")

                    .ClearFormatting()
                    .Text = textoBuscado
                    .MatchCase = False      'ignora may/min
                    .MatchWholeWord = True  'palabras completas
                    .Forward = True
                    .Replacement.ClearFormatting()
                    .Replacement.Text = textoNuevo
                    .Execute(Format:=False, Replace:=Word.WdReplace.wdReplaceAll)
                End If
            End With
        End With
    End Sub

#End Region

#Region "Shapes"

    'Reemplaza el InLineShape indicado por parámetro por el texto del BCRA
    Private Sub ReemplazarImagenBCRA(ByVal InLineShapeIndex As Integer,
                                     Optional ByVal TextoDentroDeTabla As Boolean = False,
                                     Optional ByVal confirmarCambio As Boolean = False,
                                     Optional ByVal agregarEnter As Boolean = True)
        Dim rg As Word.Range
        Dim inicio, fin As Integer

        With Globals.ThisAddIn.Application.ActiveDocument
            'Selecciona el shape
            .InlineShapes.Item(InLineShapeIndex).Select()
            'Pausa para confirmación del usuario
            If confirmarCambio Then
                If StartFrame.US.Display.MsgBox("¿Confirma el cambio?", "Confirmación del usuario requerida", US.Display.MsgBoxTipos.msgConfirmacion) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If
            End If

            If Not TextoDentroDeTabla Then
                'Agrega el texto puro
                With .Application.Selection
                    If agregarEnter Then
                        .InsertBefore(vbCrLf & "BANCO CENTRAL" & vbCrLf)
                    Else
                        .InsertBefore("BANCO CENTRAL" & vbCrLf)
                    End If
                    .InsertAfter("DE LA REPUBLICA ARGENTINA")
                End With
                FormatearSeleccion(.Application.Selection, False)
                With .Application.Selection
                    .ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                    .ParagraphFormat.LeftIndent = CentimetersToPoints(6)
                    inicio = .Start
                    fin = .End
                End With

                'Elimina el shape
                .InlineShapes.Item(InLineShapeIndex).Select()
                .Application.Selection.Delete()

                'Agrega un salto de línea
                If agregarEnter Then
                    .Range(Start:=CObj(inicio), End:=CObj(inicio)).Select()
                    .Application.Selection.InsertBreak(Type:=Word.WdBreakType.wdPageBreak)
                End If
            Else
                FormatearSeleccion(.Application.Selection, False)
                With .Application.Selection
                    .ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                    .ParagraphFormat.LeftIndent = CentimetersToPoints(6)
                End With
                'Agrega una tabla
                .Tables.Add(Range:= .Application.Selection.Range,
                               NumRows:=1, NumColumns:=1,
                               DefaultTableBehavior:=Word.WdDefaultTableBehavior.wdWord9TableBehavior,
                               AutoFitBehavior:=Word.WdAutoFitBehavior.wdAutoFitFixed)
                'Configura la tabla
                With .Application.Selection.Tables(1)
                    .Borders(Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleNone
                    .Borders(Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleNone
                    .Borders(Word.WdBorderType.wdBorderLeft).LineStyle = Word.WdLineStyle.wdLineStyleNone
                    .Borders(Word.WdBorderType.wdBorderRight).LineStyle = Word.WdLineStyle.wdLineStyleNone
                    .Rows.Height = CentimetersToPoints(0.5)
                    .Columns.Item(1).Width = CentimetersToPoints(6)
                    .Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter
                    '...celda 1
                    .Cell(1, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    rg = .Cell(1, 1).Range
                    rg.Text = "BANCO CENTRAL" & vbCrLf & "DE LA REPUBLICA ARGENTINA"
                    rg.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                End With
            End If
        End With

    End Sub

    'Reemplaza el InLineShape indicado por parámetro por el texto del BCRA
    Private Sub ReemplazarImagenBCRAenShape(ByVal ShapeIndex As Integer,
                                     Optional ByVal TextoDentroDeTabla As Boolean = False,
                                     Optional ByVal confirmarCambio As Boolean = False,
                                     Optional ByVal agregarEnter As Boolean = True)
        Dim rg As Word.Range
        Dim inicio, fin As Integer

        With Globals.ThisAddIn.Application.ActiveDocument
            'Selecciona el shape
            .Shapes.Item(ShapeIndex).Select()

            'Pone el shape en línea con el texto y almacena su posición
            Try
                .Shapes.Item(ShapeIndex).ConvertToFrame()
            Catch ex As Exception
                'si falla la conversión a frame, igualmente trata de reemplazar el shape por el texto correspondiente
            End Try
            inicio = .Application.Selection.Start
            fin = .Application.Selection.End

            'Pausa para confirmación del usuario
            If confirmarCambio Then
                If StartFrame.US.Display.MsgBox("¿Confirma el cambio?", "Confirmación del usuario requerida", US.Display.MsgBoxTipos.msgConfirmacion) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If
            End If

            'Elimina el shape
            .Application.Selection.Delete()

            'Ajusta la posición del puntero para insertar el texto
            .Range(Start:=CObj(inicio), End:=CObj(inicio)).Select()
            '.Application.Selection.MoveDown(Unit:=Word.WdUnits.wdLine, Count:=4)

            If Not TextoDentroDeTabla Then
                'Agrega el texto puro
                With .Application.Selection
                    If agregarEnter Then
                        .InsertBefore(vbCrLf & "BANCO CENTRAL" & vbCrLf)
                    Else
                        .InsertBefore("BANCO CENTRAL" & vbCrLf)
                    End If
                    .InsertAfter("DE LA REPUBLICA ARGENTINA")
                End With
                FormatearSeleccion(.Application.Selection, False)
                With .Application.Selection
                    .ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                    .ParagraphFormat.LeftIndent = CentimetersToPoints(6)
                    inicio = .Start
                    fin = .End
                End With

                'Agrega un salto de línea
                If agregarEnter Then
                    .Range(Start:=CObj(inicio), End:=CObj(inicio)).Select()
                    .Application.Selection.InsertBreak(Type:=Word.WdBreakType.wdPageBreak)
                End If
            Else
                'FormatearSeleccion(.Application.Selection, False)
                With .Application.Selection
                    .ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                    .ParagraphFormat.LeftIndent = CentimetersToPoints(6)
                End With
                'Agrega una tabla
                .Tables.Add(Range:= .Application.Selection.Range,
                               NumRows:=1, NumColumns:=1,
                               DefaultTableBehavior:=Word.WdDefaultTableBehavior.wdWord9TableBehavior,
                               AutoFitBehavior:=Word.WdAutoFitBehavior.wdAutoFitFixed)
                'Configura la tabla
                With .Application.Selection.Tables(1)
                    .Borders(Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleNone
                    .Borders(Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleNone
                    .Borders(Word.WdBorderType.wdBorderLeft).LineStyle = Word.WdLineStyle.wdLineStyleNone
                    .Borders(Word.WdBorderType.wdBorderRight).LineStyle = Word.WdLineStyle.wdLineStyleNone
                    .Rows.Height = CentimetersToPoints(0.5)
                    .Columns.Item(1).Width = CentimetersToPoints(6)
                    .Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter
                    '...celda 1
                    .Cell(1, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    rg = .Cell(1, 1).Range
                    rg.Text = "BANCO CENTRAL" & vbCrLf & "DE LA REPUBLICA ARGENTINA"
                    rg.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                End With
            End If
        End With

    End Sub

    'Reemplaza el Shape (de una columna) por una tabla con dos columnas
    Private Sub ReemplazarShapeXtabla2C(ByVal ShapeIndex As Integer)
        Dim lineas, columnas As Integer
        lineas = 1 : columnas = 2

        With Globals.ThisAddIn.Application.ActiveDocument
            'Selecciona el shape
            .Shapes.Item(ShapeIndex).Select()
            Dim texto As String = .Shapes.Item(ShapeIndex).AlternativeText()
            .Shapes.Item(ShapeIndex).Delete()

            'Agrega una tabla en el lugar del shape
            .Application.Selection.MoveUp()
            .Tables.Add(Range:= .Application.Selection.Range,
                           NumRows:=lineas, NumColumns:=columnas,
                           DefaultTableBehavior:=Word.WdDefaultTableBehavior.wdWord9TableBehavior,
                           AutoFitBehavior:=Word.WdAutoFitBehavior.wdAutoFitFixed)
            '...configura la tabla
            With .Application.Selection.Tables(1)
                'bordes
                .Borders(Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                .Borders(Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                .Borders(Word.WdBorderType.wdBorderLeft).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                .Borders(Word.WdBorderType.wdBorderRight).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                'formato de la tabla
                .Rows.Height = CentimetersToPoints(0.83)
                .Columns(1).Width = CentimetersToPoints(14.37)
                .Columns(2).Width = CentimetersToPoints(2.75)
                .Rows(1).Alignment = Word.WdRowAlignment.wdAlignRowCenter
                'celdas
                For r As Integer = 1 To .Rows.Count
                    .Rows(r).Alignment = Word.WdRowAlignment.wdAlignRowCenter
                    Dim rg As Microsoft.Office.Interop.Word.Range
                    rg = .Cell(r, 1).Range
                    rg.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    rg.Text = texto
                    rg = Nothing
                    For c As Integer = 1 To .Columns.Count
                        .Cell(r, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        .Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    Next
                Next
                'Formato
                .Select()
                Me.FormatearSeleccion(Globals.ThisAddIn.Application.ActiveDocument.Application.Selection)
            End With
        End With
    End Sub

    'Reemplaza el Shape por el texto que contiene, eliminando dicho shape
    Private Sub ReemplazarShapeXtexto(ByVal ShapeIndex As Integer)
        With Globals.ThisAddIn.Application.ActiveDocument
            'Selecciona el shape, extrae el texto y elimina el cuadro
            .Shapes.Item(ShapeIndex).Select()
            Dim texto As String = .Shapes.Item(ShapeIndex).AlternativeText()
            texto = texto.Replace("Cuadro de texto:", "")
            texto = texto.Replace("Cuadro de texto :", "")
            texto = texto.Replace("Cuadro de texto", "")
            texto = texto.Trim()
            .Shapes.Item(ShapeIndex).Delete()

            'Pega el texto en el documento
            .Application.Selection.MoveDown(Unit:=Word.WdUnits.wdLine, Count:=2)
            .Application.Selection.InsertAfter(texto)

            'Formato
            Me.FormatearSeleccion(Globals.ThisAddIn.Application.ActiveDocument.Application.Selection)
        End With
    End Sub

    'Reemplaza el Shape indicado por parámetro por una tabla con la misma estructura
    Private Sub ReemplazarShapeXtabla(ByVal ShapeIndex As Integer)
        Dim lineas, columnas As Integer

        With Globals.ThisAddIn.Application.ActiveDocument
            'Selecciona el shape
            .Shapes.Item(ShapeIndex).Select()
            .Application.Selection.MoveUp()
            'Determina las filas y columnas del shape
            DeterminarFormatoShape(.Shapes(ShapeIndex), lineas, columnas)
            'Agrega una tabla en el lugar del shape
            .Application.Selection.MoveDown(Unit:=Word.WdUnits.wdLine, Count:=4)
            .Tables.Add(Range:= .Application.Selection.Range,
                           NumRows:=lineas, NumColumns:=columnas,
                           DefaultTableBehavior:=Word.WdDefaultTableBehavior.wdWord9TableBehavior,
                           AutoFitBehavior:=Word.WdAutoFitBehavior.wdAutoFitFixed)
            '...configura la tabla
            With .Application.Selection.Tables(1)
                'bordes
                .Borders(Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                .Borders(Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                .Borders(Word.WdBorderType.wdBorderLeft).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                .Borders(Word.WdBorderType.wdBorderRight).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                'filas
                .Rows.Height = CentimetersToPoints(0.5)
                .Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter
                'celdas
                For r As Integer = 1 To .Rows.Count
                    .Rows(r).Alignment = Word.WdRowAlignment.wdAlignRowCenter
                    For c As Integer = 1 To .Columns.Count
                        .Cell(r, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        .Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    Next
                Next
                'Formato
                .Select()
                Me.FormatearSeleccion(Globals.ThisAddIn.Application.ActiveDocument.Application.Selection)
            End With
        End With
    End Sub

#End Region

#Region "Tablas"

    'Crea una tabla con el formato indicado en la selección pasada por parámetro
    Private Sub InsertarTablaCuadro(ByRef seleccion As Word.Selection, ByVal cantFilas As Integer,
                                    ByVal cantColumnas As Integer, ByVal textoCeldas As ArrayList,
                                    ByVal tipoCuadro As TiposCuadro)
        Dim rg As Word.Range

        With seleccion
            'Crea una tabla con el formato deseado
            .Tables.Add(Range:= .Range,
                           NumRows:=cantFilas, NumColumns:=cantColumnas,
                           DefaultTableBehavior:=Word.WdDefaultTableBehavior.wdWord9TableBehavior,
                           AutoFitBehavior:=Word.WdAutoFitBehavior.wdAutoFitFixed)
            'Configura la tabla
            With .Tables(1)
                '...bordes
                .Borders(Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                .Borders(Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                .Borders(Word.WdBorderType.wdBorderLeft).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                .Borders(Word.WdBorderType.wdBorderRight).LineStyle = Word.WdLineStyle.wdLineStyleSingle

                '...tamaño/alineación
                .Rows.Height = CentimetersToPoints(0.5)
                .Rows.Alignment = Word.WdRowAlignment.wdAlignRowLeft
                Select Case tipoCuadro
                    Case TiposCuadro.UnaCelda
                        .Columns.Item(1).Width = CentimetersToPoints(10)
                    Case TiposCuadro.DosColumnas
                        .Columns.Item(1).Width = CentimetersToPoints(3)
                        .Columns.Item(2).Width = CentimetersToPoints(14)
                    Case TiposCuadro.Encabezado
                        .Columns.Item(1).Width = CentimetersToPoints(14)
                        .Columns.Item(2).Width = CentimetersToPoints(3)
                        .Rows.Item(1).Height = CentimetersToPoints(1)
                        .Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter
                    Case TiposCuadro.CombinadoTresCeldas
                        .Columns.Item(1).Width = CentimetersToPoints(3)
                        .Columns.Item(2).Width = CentimetersToPoints(14)
                    Case TiposCuadro.TresColumnas
                        .Columns.Item(1).Width = CentimetersToPoints(3)
                        .Columns.Item(2).Width = CentimetersToPoints(11)
                        .Columns.Item(3).Width = CentimetersToPoints(3)
                    Case TiposCuadro.TresColumnasGrande
                        .Columns.Item(1).Width = CentimetersToPoints(3)
                        .Columns.Item(2).Width = CentimetersToPoints(11)
                        .Columns.Item(3).Width = CentimetersToPoints(3)
                    Case TiposCuadro.CuatroColumnas
                        .Columns.Item(1).Width = CentimetersToPoints(3)
                        .Columns.Item(2).Width = CentimetersToPoints(8)
                        .Columns.Item(3).Width = CentimetersToPoints(3)
                        .Columns.Item(4).Width = CentimetersToPoints(3)
                End Select

                '...celdas
                Dim t As Integer = 0
                '......filas
                For r As Integer = 1 To .Rows.Count
                    If tipoCuadro = TiposCuadro.Encabezado Then
                        .Rows(r).Alignment = Word.WdRowAlignment.wdAlignRowCenter
                    Else
                        .Rows(r).Alignment = Word.WdRowAlignment.wdAlignRowLeft
                    End If
                    '......columnas
                    For c As Integer = 1 To .Columns.Count
                        .Cell(r, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        rg = .Cell(r, c).Range
                        If tipoCuadro = TiposCuadro.CuatroColumnas Or tipoCuadro = TiposCuadro.Encabezado Then
                            rg.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                        Else
                            rg.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                        End If

                        'texto
                        If tipoCuadro = TiposCuadro.CombinadoTresCeldas And r = 2 And c = 1 Then
                            'no lleva texto esta celda
                        ElseIf tipoCuadro = TiposCuadro.TresColumnasGrande And r = 1 And c = 2 Then
                            'agrega todo el texto intermedio en esta celda (la del medio) hasta que encuentre el texto "Código"
                            For i As Integer = t To textoCeldas.Count - 1
                                If textoCeldas(i).ToString.IndexOf("Código") = -1 Then
                                    rg.Text &= textoCeldas(t)
                                    t += 1
                                Else
                                    Exit For
                                End If
                            Next
                        ElseIf tipoCuadro = TiposCuadro.TresColumnasGrande And r = 1 And c = 3 Then
                            'agrega todo el texto final en esta celda (la ultima)
                            For i As Integer = t To textoCeldas.Count - 1
                                rg.Text &= textoCeldas(t)
                                t += 1
                            Next
                        ElseIf tipoCuadro = TiposCuadro.UnaCelda Then
                            'agrega todo el texto en la única celda
                            For i As Integer = t To textoCeldas.Count - 1
                                rg.Text &= textoCeldas(t)
                                t += 1
                            Next
                        Else
                            rg.Text = textoCeldas(t)
                            t += 1
                        End If
                    Next
                Next

                'Celdas combinadas
                If tipoCuadro = TiposCuadro.CombinadoTresCeldas Then
                    .Columns(1).Cells.Merge()
                End If

                'Selecciona la tabla al finalizar
                .Select()
            End With

            'Formato del párrafo donde está la tabla
            If tipoCuadro <> TiposCuadro.Encabezado Then
                .ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            End If
            .ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle
            .ParagraphFormat.SpaceAfter = 0
            .ParagraphFormat.LeftIndent = 0
            .ParagraphFormat.LineUnitAfter = 0
            .ParagraphFormat.LineUnitBefore = 0
            .Font.Size = 10
        End With
    End Sub

    'Reemplaza las firmas del documento por una tabla
    Private Sub FirmasEnTabla(Optional ByVal seleccion As Microsoft.Office.Interop.Word.Selection = Nothing)
        Dim texto As String
        Dim rg As Word.Range

        With Globals.ThisAddIn.Application.ActiveDocument
            If Not seleccion Is Nothing Then
                'mandó las firmas por parámetro => no las busca
                texto = seleccion.Text.Trim
            Else
                'busca las firmas en el documento
                texto = ""
                For s As Integer = 1 To .Sentences.Count
                    rg = .Sentences(s)
                    rg.Select()
                    If .Application.Selection.Text.IndexOf("Saludamos a Ud") <> -1 Then
                        'Encontró el texto buscado
                        For w As Integer = 1 To 100
                            'Avanza por palabras hasta llegar al texto deseado
                            rg.MoveEnd(Unit:=Word.WdUnits.wdWord, Count:=1)
                            rg.Select()
                            If .Application.Selection.Text.Trim.IndexOf("REPUBLICA ARGENTINA") <> -1 Then
                                'Encontró el texto previo a las firmas
                                ' => Mueve el inicio de la selección al final del texto encontrado
                                rg.Start = rg.End
                                'Avanza por oraciones buscando los datos de cada firma
                                For o As Integer = 1 To 100
                                    rg.MoveEnd(Unit:=Word.WdUnits.wdSentence, Count:=1)
                                    rg.Select()
                                    If .Application.Selection.Text.Trim.IndexOf("ANEXO") <> -1 Then
                                        'Encontró el texto posterior al final de las firmas
                                        ' => Retrocede el final una oración
                                        rg.MoveEnd(Unit:=Word.WdUnits.wdSentence, Count:=-1)
                                        rg.Select()
                                        'Firmas
                                        texto = .Application.Selection.Text.Trim
                                        'Termina el bucle pq encontró el texto deseado
                                        Exit For
                                    End If
                                Next
                                'Termina el bucle pq encontró el texto deseado
                                Exit For
                            End If
                        Next
                        'Termina el bucle pq encontró el texto deseado
                        Exit For
                    End If
                Next
            End If

            If texto = "" Then
                MsgBox("No se pudieron detectar las firmas", MsgBoxStyle.Exclamation, Header)
            Else
                'Avanza la selección hasta llegar al primer caracter alfabético
                Dim char1 As Char
                char1 = CType(.Application.Selection.Text.Substring(0, 1), Char)
                Do Until Char.IsLetter(char1)
                    .Application.Selection.MoveStart(Unit:=Word.WdUnits.wdWord, Count:=1)
                    char1 = CType(.Application.Selection.Text.Substring(0, 1), Char)
                Loop

                'Separa los componentes de cada firma
                Dim n1, n2, c1, c2 As String
                Dim f2 As String() = Nothing
                Dim f1 As String() = Nothing
                n1 = "" : n2 = "" : c1 = "" : c2 = ""

                '...componentes primarios
                texto = texto.Replace(vbTab, vbCr)  'Reemplaza TAB por ENTER
                f1 = texto.Split(vbCr)              'Split por ENTER
                Dim sb As New StringBuilder()
                sb.Append("Se detectaron las siguientes firmas:" & vbCrLf)
                For i As Integer = 0 To f1.Length - 1
                    sb.AppendFormat("{0}.- {1}{2}", i + 1, f1(i), vbCrLf)
                Next
                sb.Append(vbCrLf & "¿Es correcto?")
                If US.Display.MsgBox(sb.ToString, "Confirmación del Usuario Requerida", US.Display.MsgBoxTipos.msgConfirmacion) = Windows.Forms.DialogResult.Yes Then
                    Try
                        '...nombres
                        n1 = f1(0)
                        n2 = f1(1)
                        If f1(0).IndexOf(vbTab) <> -1 Then
                            f2 = f1(0).Split(vbTab) 'Split nombres por TAB
                            n1 = f2(0)
                            n2 = f2(1)
                        End If

                        '...cargos
                        c1 = f1(2)
                        c2 = f1(3)
                        If f1.Length > 4 Then
                            c2 &= " " & f1(4)
                        End If

                    Catch ex As Exception
                        MsgBox("No se pudieron descomponer las firmas en sus partes individuales", MsgBoxStyle.Exclamation, Header)
                        n1 = "" : n2 = "" : c1 = "" : c2 = ""
                    End Try

                    'Elimina las firmas actualmente seleccionadas
                    .Application.Selection.Delete()

                    'Crea la tabla para insertar las firmas
                    .Tables.Add(Range:=.Application.Selection.Range, _
                                   NumRows:=2, NumColumns:=2, _
                                   DefaultTableBehavior:=Word.WdDefaultTableBehavior.wdWord9TableBehavior, _
                                   AutoFitBehavior:=Word.WdAutoFitBehavior.wdAutoFitFixed)
                    '...configura la tabla
                    With .Application.Selection.Tables(1)
                        .Borders(Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleNone
                        .Borders(Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleNone
                        .Borders(Word.WdBorderType.wdBorderLeft).LineStyle = Word.WdLineStyle.wdLineStyleNone
                        .Borders(Word.WdBorderType.wdBorderRight).LineStyle = Word.WdLineStyle.wdLineStyleNone
                        .Rows.Height = CentimetersToPoints(0.5)
                        .Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter

                        '...alineación
                        For r As Integer = 1 To .Rows.Count
                            .Rows(r).Alignment = Word.WdRowAlignment.wdAlignRowCenter
                            For c As Integer = 1 To .Columns.Count
                                .Cell(r, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                .Cell(r, c).Borders(Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleNone
                                .Cell(r, c).Borders(Word.WdBorderType.wdBorderRight).LineStyle = Word.WdLineStyle.wdLineStyleNone
                                rg = .Cell(r, c).Range
                                rg.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                Select Case r.ToString.Trim & "-" & c.ToString.Trim
                                    Case "1-1"
                                        rg.Text = n1
                                    Case "1-2"
                                        rg.Text = n2
                                    Case "2-1"
                                        rg.Text = c1
                                    Case "2-2"
                                        rg.Text = c2
                                End Select
                            Next
                        Next
                        'Elimina caracteres raros en las firmas
                        .Select()
                        .Application.Selection.Find.Execute(FindText:="^n", ReplaceWith:="", Replace:=Word.WdReplace.wdReplaceAll)

                        'Vuelve a seleccionar el área completa
                        .Range.Select()

                        'Tamaño de fuente 10
                        Me.FormatearSeleccion(Globals.ThisAddIn.Application.ActiveDocument.Application.Selection)

                    End With
                    'Inserta una separación adicional
                    .Application.Selection.Start = .Application.Selection.End
                    .Application.Selection.MoveEnd(Unit:=Word.WdUnits.wdLine, Count:=1)
                    .Application.Selection.InsertBefore(vbCrLf)
                    .Application.Selection.InsertBefore(vbCrLf)
                    .Application.ActiveDocument.Range(Start:=0, End:=0).Select()
                End If
            End If
        End With

    End Sub

#End Region

#End Region

End Class


'Clase auxiliar para manejo de archivos del adelgazador
Public Class Archivos

    Public Shared Function CopiarEstructura(ByVal fileNameOrigen As String, _
                                            ByVal fileNameDestino As String, _
                                            ByRef sError As String, _
                                            Optional ByVal bBorrarArchivoOrigen As Boolean = False, _
                                            Optional ByVal bPisarArchivoDestinoSiExiste As Boolean = False, _
                                            Optional ByVal bCrearPathDestino As Boolean = False) As Boolean

        Try

            'Copia un archivo 
            'y verifica la copia con el tamaño de los archivos

            If Not File.Exists(fileNameOrigen) Then
                'no existe el archivo origen
                sError = "No existe el archivo origen:" & vbCrLf & "'" & fileNameOrigen & "'"
                Return False
            End If

            If bCrearPathDestino Then
                Dim sPathDestino As String
                sPathDestino = fileNameDestino.Substring(fileNameDestino.LastIndexOf("\"))
                CrearDirectorio(sPathDestino)
            End If

            File.Copy(fileNameOrigen, fileNameDestino, bPisarArchivoDestinoSiExiste)

            If FileLen(fileNameOrigen) <> FileLen(fileNameDestino) Then
                sError = "Error al copiar el archivo origen:" & vbCrLf & "'" & fileNameOrigen & "'"
                Return False
            End If

            If bBorrarArchivoOrigen Then
                File.Delete(fileNameOrigen)
            End If

            Return True

        Catch ex As Exception
            sError = ex.Message
            Return False
        End Try

    End Function

    Public Shared Function CrearDirectorio(ByVal ruta As String) As Boolean

        Try

            'si no existe la carpeta destino la creo
            If Not Directory.Exists(ruta) Then
                Directory.CreateDirectory(ruta)
            End If

            Return True

        Catch ex As Exception
            'si no pudo crear el directorio
            Return False
        End Try

    End Function

    Public Shared Function GrabarArchivoTexto(ByVal sArchivo As String, _
                                              ByVal sContenido As String, _
                                              Optional ByVal sError As String = Nothing) As Boolean

        Try

            'grabo el html en un temporal
            Dim sw As StreamWriter = File.CreateText(sArchivo)
            sw.WriteLine(sContenido)
            sw.Close()
            sw = Nothing

            Return True

        Catch ex As Exception
            sError = ex.Message
            Return False
        End Try

    End Function

    Public Shared Function LeerArchivoTexto(ByVal sArchivo As String) As String

        'Try

        Dim sTexto As String

        'levanto el html a memoria
        Dim sr As StreamReader
        sr = File.OpenText(sArchivo)
        sTexto = sr.ReadToEnd()
        sr.Close()
        sr = Nothing

        Return sTexto

        'Catch ex As Exception
        '    Return ""
        'End Try

    End Function

End Class
