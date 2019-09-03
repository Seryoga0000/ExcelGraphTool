Imports Microsoft.Office.Tools.Ribbon

Public Class Lenta

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles IncreaseGraphButton.Click
        PerIncrChart = Val(PercentIncrease.Text)
        Call IncreaseGraph()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles ReduceGraphButton.Click
        PerIncrChart = Val(PercentIncrease.Text)
        Call ReduceGraph()
    End Sub

    Private Sub EditBox1_TextChanged(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles PercentIncrease.TextChanged
        PerIncrChart = Val(PercentIncrease.Text)
        If PerIncrChart > 99 Or PerIncrChart < 1 Then
            PerIncrChart = 10
            PercentIncrease.Text = PerIncrChart
        End If
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles Button3.Click
        On Error Resume Next
        Call PointStart()
        EBStartPoint.Text = pointnumber
    End Sub

    Private Sub EBStartPoint_TextChanged(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles EBStartPoint.TextChanged

    End Sub

    Private Sub Button4_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles Button4.Click
        On Error Resume Next
        Call PointEnd()
        EBEndPoint.Text = pointnumber
    End Sub

    Private Sub YmaxIncrease_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles YmaxIncrease.Click
        YmaxValue = Val(EBYmaxValue.Text)

        'EBYminValue.Text = EBYmaxValue.Text
        If Right(EBYmaxValue.Text, 1) = "%" Then
            Call YmaxIncreaseSub("%")
        Else
            Call YmaxIncreaseSub("_")
        End If
       

    End Sub

    Private Sub EBYmaxValue_TextChanged(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles EBYmaxValue.TextChanged
        If Right(EBYmaxValue.Text, 1) = "%" Then
            If Val(EBYmaxValue.Text) > 0 And Val(EBYmaxValue.Text) < 100 Then
            Else
                EBYmaxValue.Text = "10%"
            End If
        End If

    End Sub

    Private Sub YmaxReduce_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles YmaxReduce.Click
        YmaxValue = Val(EBYmaxValue.Text)

        'EBYminValue.Text = EBYmaxValue.Text
        If Right(EBYmaxValue.Text, 1) = "%" Then
            Call YmaxReduceSub("%")
        Else
            Call YmaxReduceSub("_")
        End If


    End Sub

    Private Sub YminIncrease_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        YminValue = Val(EBYminValue.Text)

        'EBYminValue.Text = EBYmaxValue.Text
        If Right(EBYminValue.Text, 1) = "%" Then
            Call YminIncreaseSub("%")
        Else
            Call YminIncreaseSub("_")
        End If

    End Sub

    Private Sub YminReduce_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles YminReduce.Click
        YminValue = Val(EBYminValue.Text)

        'EBYminValue.Text = EBYmaxValue.Text
        If Right(EBYminValue.Text, 1) = "%" Then
            Call YminReduceSub("%")
        Else
            Call YminReduceSub("_")
        End If


    End Sub


    Private Sub EBYminValue_TextChanged(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles EBYminValue.TextChanged
        If Right(EBYminValue.Text, 1) = "%" Then
            If Val(EBYminValue.Text) > 0 And Val(EBYminValue.Text) < 100 Then
            Else
                EBYminValue.Text = "10%"
            End If
        End If
    End Sub

    Private Sub YIncrease_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles YIncrease.Click
        YValue = Val(EBYValue.Text)

        'EBYminValue.Text = EBYmaxValue.Text
        If Right(EBYValue.Text, 1) = "%" Then
            Call YIncreaseSub("%")
        Else
            Call YIncreaseSub("_")
        End If

    End Sub

    Private Sub YReduce_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles YReduce.Click
        YValue = Val(EBYValue.Text)

        'EBYminValue.Text = EBYmaxValue.Text
        If Right(EBYValue.Text, 1) = "%" Then
            Call YReduceSub("%")
        Else
            Call YReduceSub("_")
        End If
    End Sub

    Private Sub EBYValue_TextChanged(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles EBYValue.TextChanged
        If Right(EBYValue.Text, 1) = "%" Then
            If Val(EBYValue.Text) > 0 And Val(EBYValue.Text) < 100 Then
            Else
                EBYValue.Text = "10%"
            End If
        End If
    End Sub

    Private Sub EBXAxesMoveValue_TextChanged(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles EBXAxesMoveValue.TextChanged
        If Right(EBXAxesMoveValue.Text, 1) = "%" Then
            If Val(EBXAxesMoveValue.Text) > 0 And Val(EBXAxesMoveValue.Text) < 100 Then
            Else
                EBXAxesMoveValue.Text = "10%"
            End If
        End If
    End Sub

    Private Sub BXAxesRight_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles BXAxesRight.Click
        XAxesMoveValue = Val(EBXAxesMoveValue.Text)
        If DDTime.Items(0).Tag = 0 Then
            MultTime = 1
        Else
            MultTime = DDTime.Tag * (1 / 86400)
        End If
        If Right(EBXAxesMoveValue.Text, 1) = "%" Then
            Call XAxesRightSub("%")
        Else
            Call XAxesRightSub("_")
        End If

    End Sub


    Private Sub BXAxesLeft_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles BXAxesLeft.Click
        XAxesMoveValue = Val(EBXAxesMoveValue.Text)
        If DDTime.Items(0).Tag = 0 Then
            MultTime = 1
        Else
            MultTime = DDTime.Tag * (1 / 86400)
        End If
        If Right(EBXAxesMoveValue.Text, 1) = "%" Then
            Call XAxesLeftSub("%")
        Else
            Call XAxesLeftSub("_")
        End If
    End Sub

    Private Sub EBXMoveValue_TextChanged(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles EBXMoveValue.TextChanged
        If Right(EBXMoveValue.Text, 1) = "%" Then
            If Val(EBXMoveValue.Text) > 0 And Val(EBXMoveValue.Text) < 100 Then
            Else
                EBXMoveValue.Text = "10%"
            End If
        End If
    End Sub

    Private Sub XminIncrease_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles XminIncrease.Click
        XMoveValue = Val(EBXMoveValue.Text)
        If DDTime.Items(0).Tag = 0 Then
            MultTime = 1
        Else
            MultTime = DDTime.Tag * (1 / 86400)
        End If
        If Right(EBXAxesMoveValue.Text, 1) = "%" Then
            Call XminIncreaseSub("%")
        Else
            Call XminIncreaseSub("_")
        End If
    End Sub

   
    Private Sub XmaxIncrease_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles XmaxIncrease.Click
        XMoveValue = Val(EBXMoveValue.Text)
        If DDTime.Items(0).Tag = 0 Then
            MultTime = 1
        Else
            MultTime = DDTime.Tag * (1 / 86400)
        End If
        If Right(EBXAxesMoveValue.Text, 1) = "%" Then
            Call XmaxIncreaseSub("%")
        Else
            Call XmaxIncreaseSub("_")
        End If
    End Sub

    Private Sub XminReduce_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles XminReduce.Click
        XMoveValue = Val(EBXMoveValue.Text)
        If DDTime.Items(0).Tag = 0 Then
            MultTime = 1
        Else
            MultTime = DDTime.Tag * (1 / 86400)
        End If
        If Right(EBXAxesMoveValue.Text, 1) = "%" Then
            Call XminReduceSub("%")
        Else
            Call XminReduceSub("_")
        End If
    End Sub

    Private Sub XmaxReduce_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles XmaxReduce.Click
        XMoveValue = Val(EBXMoveValue.Text)
        If DDTime.Items(0).Tag = 0 Then
            MultTime = 1
        Else
            MultTime = DDTime.Tag * (1 / 86400)
        End If
        If Right(EBXAxesMoveValue.Text, 1) = "%" Then
            Call XmaxReduceSub("%")
        Else
            Call XmaxReduceSub("_")
        End If
    End Sub

    Private Sub DDTime_SelectionChanged(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles DDTime.SelectionChanged
        If DDTime.Items(0).Tag = 0 Then
            MultTime = 1
        Else
            MultTime = DDTime.Tag * (1 / 86400)
        End If
    End Sub

    Private Sub DDSubLines_SelectionChanged(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles DDSubLines.SelectionChanged
        Call SubLine_Sub(DDSubLines.Tag)
    End Sub
End Class
