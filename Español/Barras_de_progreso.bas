Attribute VB_Name = "Barras_de_progreso"
' Declaración de variables globales
Public colorBarraTotal As Long
Public colorBarraProgreso As Long
Public bordeBarraTotal As Long
Public tamañoBordeBarra As Single
Public anchoBarra As Single
Public largoBarraTotal As Single
Public largoBarraProgreso As Single
Public alturaBarra As Single
Public posicionYBarra As Single
Public excluirPrimera As Boolean
Public excluirUltima As Boolean

Sub InicializarVariables_Barra()
    ' Inicialización de las variables globales para las barras de progreso
    colorBarraTotal = RGB(202, 202, 202) ' Color para la barra que representa el total de diapositivas (gris claro)
    bordeBarraTotal = RGB(0, 0, 0) ' Color del borde de la barra total (negro)
    tamañoBordeBarra = 0.75 ' Grosor del borde de la barra total en puntos
    
    colorBarraProgreso = RGB(0, 0, 0) ' Color de la barra que representa el progreso actual (verde)
    
    anchoBarra = 10 ' Ancho de ambas barras
    largoBarraTotal = 500 ' Largo de la barra total
    alturaBarra = ActivePresentation.PageSetup.SlideHeight - 100 ' Posición vertical de las barras
    posicionYBarra = ActivePresentation.PageSetup.SlideHeight - 50 ' Altura en Y ajustable para las barras
    
    ' Variables para exclusión de la primera o última diapositiva
    excluirPrimera = False
    excluirUltima = False
End Sub

Sub DibujarBarrasProgreso()
    Dim sld As Slide
    Dim totalSlides As Integer
    Dim currentSlide As Integer
    Dim xPos As Single
    Dim yPos As Single
    Dim largoProgreso As Single
    Dim i As Integer

    ' Asegurarse de que las variables estén inicializadas
    Call InicializarVariables_Barra

    totalSlides = ActivePresentation.Slides.Count
    
    ' Ajustar el total de diapositivas si se excluye la primera o última
    If excluirPrimera Then totalSlides = totalSlides - 1
    If excluirUltima Then totalSlides = totalSlides - 1

    ' Calcular la posición inicial en X para la barra
    xPos = (ActivePresentation.PageSetup.SlideWidth - largoBarraTotal) / 2
    yPos = posicionYBarra ' Posición en Y ajustable

    ' Borrar barras anteriores
    Call BorrarBarras_TodasLasDiapositivas

    ' Dibujar barras en todas las diapositivas
    For Each sld In ActivePresentation.Slides
        currentSlide = sld.SlideIndex
        
        ' Saltar la primera o última diapositiva si están excluidas
        If excluirPrimera And currentSlide = 1 Then GoTo Siguiente
        If excluirUltima And currentSlide = ActivePresentation.Slides.Count Then GoTo Siguiente

        ' Dibujar barra total (Barra1)
        With sld.Shapes.AddShape(msoShapeRectangle, xPos, yPos, largoBarraTotal, anchoBarra)
            .Fill.ForeColor.RGB = colorBarraTotal
            .Line.ForeColor.RGB = bordeBarraTotal
            .Line.Weight = tamañoBordeBarra ' Grosor del borde en puntos
            .Name = "BarraTotal"
        End With

        ' Calcular el largo de la barra de progreso según el número de diapositivas avanzadas
        largoProgreso = (currentSlide / totalSlides) * largoBarraTotal

        ' Dibujar barra de progreso (Barra2)
        With sld.Shapes.AddShape(msoShapeRectangle, xPos, yPos, largoProgreso, anchoBarra)
            .Fill.ForeColor.RGB = colorBarraProgreso
            .Line.Visible = msoFalse ' Sin borde para la barra de progreso
            .Name = "BarraProgreso"
        End With

Siguiente:
    Next sld
End Sub

Sub BorrarBarras_TodasLasDiapositivas()
    Dim sld As Slide
    Dim shp As Shape
    ' Limpiar barras anteriores generadas por esta macro en todas las diapositivas
    For Each sld In ActivePresentation.Slides
        For i = sld.Shapes.Count To 1 Step -1
            Set shp = sld.Shapes(i)
            If shp.Name = "BarraTotal" Or shp.Name = "BarraProgreso" Then
                shp.Delete
            End If
        Next i
    Next sld
End Sub

Sub BorrarBarras_DiapositivaActual()
    ' Elimina las barras de la diapositiva actual
    Dim sld As Slide
    Dim shp As Shape
    Set sld = ActiveWindow.View.Slide

    For i = sld.Shapes.Count To 1 Step -1
        Set shp = sld.Shapes(i)
        If shp.Name = "BarraTotal" Or shp.Name = "BarraProgreso" Then
            shp.Delete
        End If
    Next i
End Sub



