
Module Mac
    
    Sub Tset()
        Dim myrange As Microsoft.Office.Interop.Excel.Range
        'Dim MyBook As Microsoft.Office.Interop.Excel.Workbook
        ' Dim MySheet As Microsoft.Office.Interop.Excel.Worksheet
        'MyBook()
        'MySheet = MyBook.ActiveSheet
        'Dim app As New Application
        'Dim wb = app.Workbooks.Add
        'Dim ws As Worksheet = wb.Worksheets(1)
        'Dim r = ws.Range("A1")
        'r.Value = "Hello!"
        'app.Visible = True
        'Microsoft.Office.Interop.Excel.Application()
        myrange = Globals.ThisAddIn.Application.Range("A1")
        myrange.Value = 7
    End Sub
    Sub ReduceGraph()

        'debug parameters
        ' PerIncrChart = 0
        'PerIncrChart()
        '*******************
        Dim WorkChart As Excel.Chart
        WorkChart = Globals.ThisAddIn.Application.ActiveChart
        If Not WorkChart Is Nothing Then
            Dim X
            Dim Y
            Y = WorkChart.SeriesCollection(1).Values
            X = WorkChart.SeriesCollection(1).XValues

            WorkChart.Axes(xlCategory).MaximumScaleIsAuto = False
            WorkChart.Axes(xlCategory).MinimumScaleIsAuto = False

            Dim PerIncreaseChart As Long
            If PerIncrChart = 0 Then
                PerIncreaseChart = 10
            Else
                PerIncreaseChart = PerIncrChart
            End If

            Dim xdatamax As Double
            Dim xdatamin As Double
            xdatamax = UBound(X)
            xdatamin = LBound(X)

            Dim xmax As Double
            Dim xmin As Double

            xmax = WorkChart.Axes(xlCategory).MaximumScale
            xmin = WorkChart.Axes(xlCategory).MinimumScale

            If xmin > X(xdatamin) And xmax < X(xdatamax) Then
                WorkChart.Axes(xlCategory).MinimumScale = xmin - ((xmax - xmin) / 100) * PerIncreaseChart
                WorkChart.Axes(xlCategory).MaximumScale = xmax + ((xmax - xmin) / 100) * PerIncreaseChart
            Else
                WorkChart.Axes(xlCategory).MinimumScale = X(xdatamin)
                WorkChart.Axes(xlCategory).MaximumScale = X(xdatamax)
            End If

        End If
    End Sub
    Sub IncreaseGraph()

        If TypeName(Globals.ThisAddIn.Application.Selection) = "Point" Then

            'Присвоение активной диаграммы переменной WorkChart
            Dim WorkChart As Excel.Chart
            WorkChart = Globals.ThisAddIn.Application.ActiveChart

            'Вычисление номера выделенной точки
            Dim PointNumber As Long
            Dim SeriesNumber As Long
            Call PointValue(PointNumber, SeriesNumber)

            'Запись минимального и максимального значения по оси Х диаграммы
            Dim xmax As Double
            Dim xmin As Double
            xmax = WorkChart.Axes(xlCategory).MaximumScale
            xmin = WorkChart.Axes(xlCategory).MinimumScale

            'Запись массивов данных по осям в переменные Х и Y
            Dim X
            Dim Y
            Y = WorkChart.SeriesCollection(SeriesNumber).Values
            X = WorkChart.SeriesCollection(SeriesNumber).XValues

            Dim PerIncreaseChart As Long
            If PerIncrChart = 0 Then
                PerIncreaseChart = 10
            Else
                PerIncreaseChart = Natural(PerIncrChart)
            End If

            'вычисление попадает ли точка в 1-процентную область от центра
            'если да границы уменьшаются на заданный процент в переменной PerIncreaseChart
            'если нет границы выравниваются чтоб было да
            If X(PointNumber) > xmin + (xmax - xmin) / 2 + (xmax - xmin) / 100 Then
                WorkChart.Axes(xlCategory).MinimumScale = 2 * X(PointNumber) - xmax
            End If

            If X(PointNumber) < xmin + (xmax - xmin) / 2 - (xmax - xmin) / 100 Then
                WorkChart.Axes(xlCategory).MaximumScale = 2 * X(PointNumber) - xmin
            End If

            If X(PointNumber) > xmin + (xmax - xmin) / 2 - (xmax - xmin) / 100 And X(PointNumber) < xmin + (xmax - xmin) / 2 + (xmax - xmin) / 100 Then
                WorkChart.Axes(xlCategory).MinimumScale = xmin + ((X(PointNumber) - xmin) / 100) * PerIncreaseChart
                WorkChart.axes(xlCategory).MaximumScale = xmax - ((xmax - X(PointNumber)) / 100) * PerIncreaseChart
            End If
        Else
            With Globals.ThisAddIn.Application
                If .ActiveChart Is Nothing Then
                Else
                    Dim xmax As Double
                    Dim xmin As Double

                    Dim PerIncreaseChart As Long
                    If PerIncrChart = 0 Then
                        PerIncreaseChart = 10
                    Else
                        PerIncreaseChart = Natural(PerIncrChart)
                    End If

                    xmax = .ActiveChart.Axes(xlCategory).MaximumScale
                    xmin = .ActiveChart.Axes(xlCategory).MinimumScale
                    .ActiveChart.Axes(xlCategory).MinimumScale = xmin + ((xmax - xmin) / 100) * PerIncreaseChart
                    .ActiveChart.Axes(xlCategory).MaximumScale = xmax - ((xmax - xmin) / 100) * PerIncreaseChart
                End If
            End With
        End If
    End Sub
    Sub PointValue(ByRef PointNumber, Optional ByRef SeriesNumber = 1)
        PointNumber = 0
        'проверка что выделена точка на диаграмме
        If TypeName(Globals.ThisAddIn.Application.Selection) = "Point" Then
            'присвоение переменной WorkPoint веделенной точки
            Dim WorkPoint As Excel.Point
            WorkPoint = Globals.ThisAddIn.Application.Selection
            'вычисление номера точки и номера ряда по имени (работает если рядов <10)
            PointNumber = Val(Right(WorkPoint.Name, Len(WorkPoint.Name) - 3))
            SeriesNumber = Val(Mid(WorkPoint.Name, 2, 1))
            '
            WorkPoint = Nothing
        End If
    End Sub
    Function Natural(WorkNumber) As Long
        If WorkNumber < 1 Then
            Natural = 1
        Else : Natural = WorkNumber
        End If
    End Function
    Function Positive(WorkNumber) As Long
        If WorkNumber < 0 Then
            Positive = 0
        Else : Positive = WorkNumber
        End If
    End Function
    Sub PointStart()
        With Globals.ThisAddIn.Application
            If .ActiveChart Is Nothing Then 'Проверка что выделен график
            Else
                pointnumber = 10
                Dim X

                Call PointValue(pointnumber) 'присваетвает переменной номер выделенной на графике точки
                'FormatAxis1 = ActiveChart.Axes(xlCategory).TickLabels.NumberFormat
                If pointnumber > 0 Then

                    X = .ActiveChart.SeriesCollection(1).XValues 'массив Х присваевает значения массива оси х диаграммы

                    .ActiveChart.Axes(xlCategory).MinimumScale = X(pointnumber) 'левая граница диаграммы уст на выбранное значение
                End If
            End If
        End With
    End Sub
    Sub PointEnd() 'аналогично функции PointStart
        With Globals.ThisAddIn.Application
            If .ActiveChart Is Nothing Then
            Else
                pointnumber = 20
                Dim X
                Call PointValue(pointnumber)

                If pointnumber > 0 Then

                    X = .ActiveChart.SeriesCollection(1).XValues
                    .ActiveChart.Axes(xlCategory).MaximumScale = X(pointnumber)
                End If
            End If
        End With
    End Sub
    Sub YmaxIncreaseSub(Type As String)
        With Globals.ThisAddIn.Application
            'If TypeName(Globals.ThisAddIn.Application.Selection) = "Point" Then
            'Else
            If .ActiveChart Is Nothing Then
            Else
                Dim ymax As Double
                Dim ymin As Double
                ymax = .ActiveChart.Axes(xlValue).MaximumScale
                ymin = .ActiveChart.Axes(xlValue).MinimumScale
                .ActiveChart.Axes(xlValue).MaximumScaleIsAuto = False
                .ActiveChart.Axes(xlValue).MinimumScaleIsAuto = False
                Select Case Type
                    Case "%"
                        .ActiveChart.Axes(xlValue).MaximumScale = ymax - ((ymax - ymin) / 100) * YmaxValue
                        '.ActiveChart.Axes(xlValue).MinimumScale = ymin + ((ymax - ymin) / 100) * YmaxValue
                    Case "_"
                        .ActiveChart.Axes(xlValue).MaximumScale = ymax - YmaxValue
                End Select
            End If
            'End If
        End With

    End Sub
    Sub YmaxReduceSub(Type As String)
        With Globals.ThisAddIn.Application
            'If TypeName(Globals.ThisAddIn.Application.Selection) = "Point" Then
            'Else
            If .ActiveChart Is Nothing Then
            Else
                Dim ymax As Double
                Dim ymin As Double
                ymax = .ActiveChart.Axes(xlValue).MaximumScale
                ymin = .ActiveChart.Axes(xlValue).MinimumScale
                .ActiveChart.Axes(xlValue).MaximumScaleIsAuto = False
                .ActiveChart.Axes(xlValue).MinimumScaleIsAuto = False
                Select Case Type
                    Case "%"
                        .ActiveChart.Axes(xlValue).MaximumScale = ymax + ((ymax - ymin) / 100) * YmaxValue
                        '.ActiveChart.Axes(xlValue).MinimumScale = ymin + ((ymax - ymin) / 100) * YmaxValue
                    Case "_"
                        .ActiveChart.Axes(xlValue).MaximumScale = ymax + YmaxValue
                End Select
            End If
            'End If
        End With

    End Sub
    Sub YminIncreaseSub(Type As String)
        With Globals.ThisAddIn.Application
            'If TypeName(Globals.ThisAddIn.Application.Selection) = "Point" Then
            'Else
            If .ActiveChart Is Nothing Then
            Else
                Dim ymax As Double
                Dim ymin As Double
                ymax = .ActiveChart.Axes(xlValue).MaximumScale
                ymin = .ActiveChart.Axes(xlValue).MinimumScale
                .ActiveChart.Axes(xlValue).MaximumScaleIsAuto = False
                .ActiveChart.Axes(xlValue).MinimumScaleIsAuto = False
                Select Case Type
                    Case "%"
                        .ActiveChart.Axes(xlValue).MinimumScale = ymin - ((ymax - ymin) / 100) * YminValue
                        '.ActiveChart.Axes(xlValue).MinimumScale = ymin + ((ymax - ymin) / 100) * YmaxValue
                    Case "_"
                        .ActiveChart.Axes(xlValue).MinimumScale = ymin - YminValue
                End Select
            End If
            'End If
        End With

    End Sub
    Sub YminReduceSub(Type As String)
        With Globals.ThisAddIn.Application
            'If TypeName(Globals.ThisAddIn.Application.Selection) = "Point" Then
            'Else
            If .ActiveChart Is Nothing Then
            Else
                Dim ymax As Double
                Dim ymin As Double
                ymax = .ActiveChart.Axes(xlValue).MaximumScale
                ymin = .ActiveChart.Axes(xlValue).MinimumScale
                .ActiveChart.Axes(xlValue).MaximumScaleIsAuto = False
                .ActiveChart.Axes(xlValue).MinimumScaleIsAuto = False
                Select Case Type
                    Case "%"
                        .ActiveChart.Axes(xlValue).MinimumScale = ymin + ((ymax - ymin) / 100) * YminValue
                        '.ActiveChart.Axes(xlValue).MinimumScale = ymin + ((ymax - ymin) / 100) * YmaxValue
                    Case "_"
                        .ActiveChart.Axes(xlValue).MinimumScale = ymin + YminValue
                End Select
            End If
            'End If
        End With

    End Sub
    Sub YIncreaseSub(Type As String)
        With Globals.ThisAddIn.Application
            If TypeName(.Selection) = "Point" Then
                Dim PointNumber As Long
                Dim SeriesNumber As Long
                Call PointValue(PointNumber, SeriesNumber)

                Dim ymax As Double
                Dim ymin As Double
                ymax = .ActiveChart.Axes(xlValue).MaximumScale
                ymin = .ActiveChart.Axes(xlValue).MinimumScale
                .ActiveChart.Axes(xlValue).MaximumScaleIsAuto = False
                .ActiveChart.Axes(xlValue).MinimumScaleIsAuto = False

                'Dim X
                Dim Y
                Y = .ActiveChart.SeriesCollection(SeriesNumber).Values
                'X = .ActiveChart.SeriesCollection(SeriesNumber).XValues

                If Y(PointNumber) > ymin + (ymax - ymin) / 2 + (ymax - ymin) / 100 Then
                    .ActiveChart.Axes(xlValue).MinimumScale = 2 * Y(PointNumber) - ymax
                End If

                If Y(PointNumber) < ymin + (ymax - ymin) / 2 - (ymax - ymin) / 100 Then
                    .ActiveChart.Axes(xlValue).MaximumScale = 2 * Y(PointNumber) - ymin
                End If

                 Select Type
                    Case "%"
                        If Y(PointNumber) > ymin + (ymax - ymin) / 2 - (ymax - ymin) / 100 And Y(PointNumber) < ymin + (ymax - ymin) / 2 + (ymax - ymin) / 100 Then
                            .ActiveChart.Axes(xlValue).MinimumScale = ymin + ((Y(PointNumber) - ymin) / 100) * YValue
                            .ActiveChart.Axes(xlValue).MaximumScale = ymax - ((ymax - Y(PointNumber)) / 100) * YValue
                        End If
                    Case "_"
                        If Y(PointNumber) > ymin + (ymax - ymin) / 2 - (ymax - ymin) / 100 And Y(PointNumber) < ymin + (ymax - ymin) / 2 + (ymax - ymin) / 100 Then
                            .ActiveChart.Axes(xlValue).MinimumScale = ymin + YValue
                            .ActiveChart.Axes(xlValue).MaximumScale = ymax - YValue
                        End If
                End Select

            Else
                If .ActiveChart Is Nothing Then
                Else
                    Dim ymax As Double
                    Dim ymin As Double
                    ymax = .ActiveChart.Axes(xlValue).MaximumScale
                    ymin = .ActiveChart.Axes(xlValue).MinimumScale
                    .ActiveChart.Axes(xlValue).MaximumScaleIsAuto = False
                    .ActiveChart.Axes(xlValue).MinimumScaleIsAuto = False
                    Select Case Type
                        Case "%"
                            .ActiveChart.Axes(xlValue).MaximumScale = ymax - ((ymax - ymin) / 100) * YValue
                            .ActiveChart.Axes(xlValue).MinimumScale = ymin + ((ymax - ymin) / 100) * YValue
                        Case "_"
                            .ActiveChart.Axes(xlValue).MaximumScale = ymax - YValue
                            .ActiveChart.Axes(xlValue).MinimumScale = ymin + YValue
                    End Select
                End If
            End If
        End With

    End Sub
    Sub YReduceSub(Type As String)
        With Globals.ThisAddIn.Application
            'If TypeName(.Selection) = "Point" Then

            'Else
            If .ActiveChart Is Nothing Then
            Else
                Dim ymax As Double
                Dim ymin As Double
                ymax = .ActiveChart.Axes(xlValue).MaximumScale
                ymin = .ActiveChart.Axes(xlValue).MinimumScale
                .ActiveChart.Axes(xlValue).MaximumScaleIsAuto = False
                .ActiveChart.Axes(xlValue).MinimumScaleIsAuto = False
                Select Case Type
                    Case "%"
                        .ActiveChart.Axes(xlValue).MaximumScale = ymax + ((ymax - ymin) / 100) * YValue
                        .ActiveChart.Axes(xlValue).MinimumScale = ymin - ((ymax - ymin) / 100) * YValue
                    Case "_"
                        .ActiveChart.Axes(xlValue).MaximumScale = ymax + YValue
                        .ActiveChart.Axes(xlValue).MinimumScale = ymin - YValue
                End Select
            End If
            'End If
        End With
    End Sub
    Sub XAxesRightSub(Type As String)
        With Globals.ThisAddIn.Application
            If .ActiveChart Is Nothing Then
            Else
                Dim xmax As Double
                Dim xmin As Double
                xmax = .ActiveChart.Axes(xlCategory).MaximumScale
                xmin = .ActiveChart.Axes(xlCategory).MinimumScale

                Select Type
                    Case "%"
                        .ActiveChart.Axes(xlCategory).MinimumScale = xmin + ((xmax - xmin) / 100) * XAxesMoveValue
                        .ActiveChart.Axes(xlCategory).MaximumScale = xmax + ((xmax - xmin) / 100) * XAxesMoveValue
                    Case "_"
                        .ActiveChart.Axes(xlCategory).MinimumScale = xmin + XAxesMoveValue
                        .ActiveChart.Axes(xlCategory).MaximumScale = xmax + XAxesMoveValue
                End Select
            End If
        End With
    End Sub

    Sub XAxesLeftSub(Type As String)
        With Globals.ThisAddIn.Application
            If .ActiveChart Is Nothing Then
            Else
                Dim xmax As Double
                Dim xmin As Double
                xmax = .ActiveChart.Axes(xlCategory).MaximumScale
                xmin = .ActiveChart.Axes(xlCategory).MinimumScale

                Select Case Type
                    Case "%"
                        .ActiveChart.Axes(xlCategory).MinimumScale = xmin - ((xmax - xmin) / 100) * XAxesMoveValue
                        .ActiveChart.Axes(xlCategory).MaximumScale = xmax - ((xmax - xmin) / 100) * XAxesMoveValue
                    Case "_"
                        .ActiveChart.Axes(xlCategory).MinimumScale = xmin - XAxesMoveValue
                        .ActiveChart.Axes(xlCategory).MaximumScale = xmax - XAxesMoveValue
                End Select
            End If
        End With
    End Sub

    Sub XminIncreaseSub(Type As String)
        With Globals.ThisAddIn.Application
            If .ActiveChart Is Nothing Then
            Else
                Dim xmax As Double
                Dim xmin As Double
                xmax = .ActiveChart.Axes(xlCategory).MaximumScale
                xmin = .ActiveChart.Axes(xlCategory).MinimumScale

                Select Case Type
                    Case "%"
                        .ActiveChart.Axes(xlCategory).MinimumScale = xmin + ((xmax - xmin) / 100) * XMoveValue
                        '.ActiveChart.Axes(xlCategory).MaximumScale = xmax - ((xmax - xmin) / 100) * XMoveValue
                    Case "_"
                        .ActiveChart.Axes(xlCategory).MinimumScale = xmin + XMoveValue
                        '.ActiveChart.Axes(xlCategory).MaximumScale = xmax - XMoveValue
                End Select
            End If
        End With
    End Sub

    Sub XmaxIncreaseSub(Type As String)
        With Globals.ThisAddIn.Application
            If .ActiveChart Is Nothing Then
            Else
                Dim xmax As Double
                Dim xmin As Double
                xmax = .ActiveChart.Axes(xlCategory).MaximumScale
                xmin = .ActiveChart.Axes(xlCategory).MinimumScale

                Select Case Type
                    Case "%"
                        '.ActiveChart.Axes(xlCategory).MinimumScale = xmin + ((xmax - xmin) / 100) * XMoveValue
                        .ActiveChart.Axes(xlCategory).MaximumScale = xmax + ((xmax - xmin) / 100) * XMoveValue
                    Case "_"
                        '.ActiveChart.Axes(xlCategory).MinimumScale = xmin + XMoveValue
                        .ActiveChart.Axes(xlCategory).MaximumScale = xmax + XMoveValue
                End Select
            End If
        End With
    End Sub

    Sub XmaxReduceSub(Type As String)
        With Globals.ThisAddIn.Application
            If .ActiveChart Is Nothing Then
            Else
                Dim xmax As Double
                Dim xmin As Double
                xmax = .ActiveChart.Axes(xlCategory).MaximumScale
                xmin = .ActiveChart.Axes(xlCategory).MinimumScale

                Select Case Type
                    Case "%"
                        .ActiveChart.Axes(xlCategory).MaximumScale = xmax - ((xmax - xmin) / 100) * XMoveValue
                        '.ActiveChart.Axes(xlCategory).MaximumScale = xmax - ((xmax - xmin) / 100) * XMoveValue
                    Case "_"
                        .ActiveChart.Axes(xlCategory).MaximumScale = xmax - XMoveValue
                        '.ActiveChart.Axes(xlCategory).MaximumScale = xmax - XMoveValue
                End Select
            End If
        End With
    End Sub

    Sub XminReduceSub(Type As String)
        With Globals.ThisAddIn.Application
            If .ActiveChart Is Nothing Then
            Else
                Dim xmax As Double
                Dim xmin As Double
                xmax = .ActiveChart.Axes(xlCategory).MaximumScale
                xmin = .ActiveChart.Axes(xlCategory).MinimumScale

                Select Case Type
                    Case "%"
                        '.ActiveChart.Axes(xlCategory).MinimumScale = xmin - ((xmax - xmin) / 100) * XMoveValue
                        .ActiveChart.Axes(xlCategory).MinimumScale = xmin - ((xmax - xmin) / 100) * XMoveValue
                    Case "_"
                        '.ActiveChart.Axes(xlCategory).MinimumScale = xmin - XMoveValue
                        .ActiveChart.Axes(xlCategory).MinimumScale = xmin - XMoveValue
                End Select
            End If
        End With
    End Sub
    Sub SubLine_Sub(Optional div As Double = 5)
        With Globals.ThisAddIn.Application
            .ActiveChart.Axes(xlCategory).MinorUnit = .ActiveChart.Axes(xlCategory).MajorUnit / div
        End With
    End Sub
End Module
