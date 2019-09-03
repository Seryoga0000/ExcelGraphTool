Imports Microsoft.Office.Tools.Ribbon

Public Class Lenta

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles IncreaseGraphButton.Click
        Call BuckUpSet()
        PerIncrChart = Val(PercentIncrease.Text)
        Call IncreaseGraph()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles ReduceGraphButton.Click
        Call BuckUpSet()
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
        Call BuckUpSet()
        'On Error Resume Next
        Call PointStart()
        'EBStartPoint.Text = pointnumber
    End Sub

    Private Sub EBStartPoint_TextChanged(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)

    End Sub

    Private Sub Button4_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles Button4.Click
        Call BuckUpSet()
        'On Error Resume Next
        Call PointEnd()
        'EBEndPoint.Text = pointnumber
    End Sub

    Private Sub YmaxIncrease_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles YmaxIncrease.Click
        Call BuckUpSet()
        YmaxValue = Val(EBYValue.Text)

        'EBYminValue.Text = EBYmaxValue.Text
        If Right(EBYValue.Text, 1) = "%" Then
            Call YmaxIncreaseSub("%")
        Else
            Call YmaxIncreaseSub("_")
        End If
       

    End Sub

    Private Sub EBYmaxValue_TextChanged(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        'If Right(EBYmaxValue.Text, 1) = "%" Then
        '    If Val(EBYmaxValue.Text) > 0 And Val(EBYmaxValue.Text) < 100 Then
        '    Else
        '        EBYmaxValue.Text = "10%"
        '    End If
        'End If

    End Sub

    Private Sub YmaxReduce_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        'YmaxValue = Val(EBYmaxValue.Text)

        ''EBYminValue.Text = EBYmaxValue.Text
        'If Right(EBYmaxValue.Text, 1) = "%" Then
        '    Call YmaxReduceSub("%")
        'Else
        '    Call YmaxReduceSub("_")
        'End If


    End Sub

    Private Sub YminIncrease_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        YminValue = Val(EBYValue.Text)

        'EBYminValue.Text = EBYmaxValue.Text
        If Right(EBYValue.Text, 1) = "%" Then
            Call YminIncreaseSub("%")
        Else
            Call YminIncreaseSub("_")
        End If

    End Sub

    Private Sub YminReduce_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        'YminValue = Val(EBYminValue.Text)

        ''EBYminValue.Text = EBYmaxValue.Text
        'If Right(EBYminValue.Text, 1) = "%" Then
        '    Call YminReduceSub("%")
        'Else
        '    Call YminReduceSub("_")
        'End If


    End Sub


    Private Sub EBYminValue_TextChanged(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        'If Right(EBYminValue.Text, 1) = "%" Then
        '    If Val(EBYminValue.Text) > 0 And Val(EBYminValue.Text) < 100 Then
        '    Else
        '        EBYminValue.Text = "10%"
        '    End If
        'End If
    End Sub

    Private Sub YIncrease_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles YIncrease.Click
        Call BuckUpSet()
        YValue = Val(EBYValue.Text)

        'EBYminValue.Text = EBYmaxValue.Text
        If Right(EBYValue.Text, 1) = "%" Then
            Call YIncreaseSub("%")
        Else
            Call YIncreaseSub("_")
        End If

    End Sub

    Private Sub YReduce_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
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
            If Val(EBYValue.Text) > -100 And Val(EBYValue.Text) < 100 Then
            Else
                EBYValue.Text = "10%"
            End If
        End If
    End Sub

    Private Sub EBXAxesMoveValue_TextChanged(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles EBXAxesMoveValue.TextChanged
        If Right(EBXAxesMoveValue.Text, 1) = "%" Then
            If Val(EBXAxesMoveValue.Text) > -100 And Val(EBXAxesMoveValue.Text) < 100 Then
            Else
                EBXAxesMoveValue.Text = "10%"
            End If
        End If
        'EBXAxesMoveValue.
    End Sub

    Private Sub BXAxesRight_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles BXAxesRight.Click
        Call BuckUpSet()
        XAxesMoveValue = Val(EBXAxesMoveValue.Text)
        If DDTime.Items(DDTime.SelectedItemIndex).Tag = 0 Then
            MultTime = 1
        Else
            MultTime = Val(DDTime.Items(DDTime.SelectedItemIndex).Tag) * (1 / 86400)
        End If
        If Right(EBXAxesMoveValue.Text, 1) = "%" Then
            Call XAxesRightSub("%")
        Else
            Call XAxesRightSub("_")
        End If

    End Sub


    Private Sub BXAxesLeft_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        XAxesMoveValue = Val(EBXAxesMoveValue.Text)
        If DDTime.Items(DDTime.SelectedItemIndex).Tag = 0 Then
            MultTime = 1
        Else
            MultTime = Val(DDTime.Items(DDTime.SelectedItemIndex).Tag) * (1 / 86400)
        End If
        If Right(EBXAxesMoveValue.Text, 1) = "%" Then
            Call XAxesLeftSub("%")
        Else
            Call XAxesLeftSub("_")
        End If
    End Sub

    Private Sub EBXMoveValue_TextChanged(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        'If Right(EBXMoveValue.Text, 1) = "%" Then
        '    If Val(EBXMoveValue.Text) > 0 And Val(EBXMoveValue.Text) < 100 Then
        '    Else
        '        EBXMoveValue.Text = "10%"
        '    End If
        'End If
    End Sub

    Private Sub XminIncrease_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles XminIncrease.Click
        Call BuckUpSet()
        XMoveValue = Val(EBXAxesMoveValue.Text)
        If DDTime.Items(DDTime.SelectedItemIndex).Tag = 0 Then
            MultTime = 1
        Else
            MultTime = Val(DDTime.Items(DDTime.SelectedItemIndex).Tag) * (1 / 86400)
        End If
        If Right(EBXAxesMoveValue.Text, 1) = "%" Then
            Call XminIncreaseSub("%")
        Else
            Call XminIncreaseSub("_")
        End If
    End Sub

   
    Private Sub XmaxIncrease_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles XmaxIncrease.Click
        Call BuckUpSet()
        XMoveValue = Val(EBXAxesMoveValue.Text)
        If DDTime.Items(DDTime.SelectedItemIndex).Tag = 0 Then
            MultTime = 1
        Else
            MultTime = Val(DDTime.Items(DDTime.SelectedItemIndex).Tag) * (1 / 86400)
        End If
        If Right(EBXAxesMoveValue.Text, 1) = "%" Then
            Call XmaxIncreaseSub("%")
        Else
            Call XmaxIncreaseSub("_")
        End If
    End Sub

    Private Sub XminReduce_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        XMoveValue = Val(EBXAxesMoveValue.Text)
        If DDTime.Items(DDTime.SelectedItemIndex).Tag = 0 Then
            MultTime = 1
        Else
            MultTime = Val(DDTime.Items(DDTime.SelectedItemIndex).Tag) * (1 / 86400)
        End If
        If Right(EBXAxesMoveValue.Text, 1) = "%" Then
            Call XminReduceSub("%")
        Else
            Call XminReduceSub("_")
        End If
    End Sub

    Private Sub XmaxReduce_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        XMoveValue = Val(EBXAxesMoveValue.Text)
        If DDTime.Items(DDTime.SelectedItemIndex).Tag = 0 Then
            MultTime = 1
        Else
            MultTime = Val(DDTime.Items(DDTime.SelectedItemIndex).Tag) * (1 / 86400)
        End If
        If Right(EBXAxesMoveValue.Text, 1) = "%" Then
            Call XmaxReduceSub("%")
        Else
            Call XmaxReduceSub("_")
        End If
    End Sub

    Private Sub DDTime_SelectionChanged(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles DDTime.SelectionChanged
        If DDTime.Items(DDTime.SelectedItemIndex).Tag = 0 Then
            MultTime = 1
        Else
            MultTime = Val(DDTime.Items(DDTime.SelectedItemIndex).Tag) * (1 / 86400)
        End If
    End Sub

    Private Sub DDSubLines_SelectionChanged(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles DDSubLines.SelectionChanged
        Call SubLine_Sub(Val(DDSubLines.Items(DDSubLines.SelectedItemIndex).Tag))
    End Sub

    Private Sub BSet_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles BSet.Click
        Call BuckUpSet()
        XMoveValue = Val(BSetValue.Text)
        If DDTime.Items(DDTime.SelectedItemIndex).Tag = 0 Then
            MultTime = 1
        Else
            MultTime = Val(DDTime.Items(DDTime.SelectedItemIndex).Tag) * (1 / 86400)
        End If
        If Right(BSetValue.Text, 1) = "%" Then
        Else
            Call SetStepXAxisSub()
        End If

    End Sub

    Private Sub YminIncrease_Click_1(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles YminIncrease.Click
        Call BuckUpSet()
        YminValue = Val(EBYValue.Text)

        'EBYminValue.Text = EBYmaxValue.Text
        If Right(EBYValue.Text, 1) = "%" Then
            Call YminIncreaseSub("%")
        Else
            Call YminIncreaseSub("_")
        End If

    End Sub

    Private Sub BSetYStep_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles BSetYStep.Click
        Call BuckUpSet()
        YValue = Val(BSetYStepValue.Text)
        If Right(BSetYStepValue.Text, 1) = "%" Then

        Else
            Call SetYstepSub()
        End If
    End Sub

    Private Sub BSetYmax_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles BSetYmax.Click
        Call BuckUpSet()
        YmaxValue = Val(BSetYmaxValue.Text)

        'EBYminValue.Text = EBYmaxValue.Text
        If Right(BSetYmaxValue.Text, 1) = "%" Then

        Else
            Call SetYmaxSub()
        End If
    End Sub

    Private Sub BSetYmin_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles BSetYmin.Click
        Call BuckUpSet()
        YminValue = Val(BSetYminValue.Text)

        'EBYminValue.Text = EBYmaxValue.Text
        If Right(BSetYminValue.Text, 1) = "%" Then

        Else
            Call SetYminSub()
        End If
    End Sub

    Private Sub BSetXmax_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles BSetXmax.Click
        Call BuckUpSet()
        XMoveValue = Val(BSetXmaxValue.Text)
        'If DDTime.Items(DDTime.SelectedItemIndex).Tag = 0 Then
        '    MultTime = 1
        'Else
        '    MultTime = Val(DDTime.Items(DDTime.SelectedItemIndex).Tag) * (1 / 86400)
        'End If
        If Right(BSetXmaxValue.Text, 1) = "%" Then

        Else
            Call XmaxSetSub()
        End If

    End Sub

    Private Sub BSetXmin_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles BSetXmin.Click
        Call BuckUpSet()
        XMoveValue = Val(BSetXminValue.Text)
        If Right(BSetXminValue.Text, 1) = "%" Then

        Else
            Call XminSetSub()
        End If

    End Sub

    Private Sub PlusMinesX_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles PlusMinesX.Click
        If Left(EBXAxesMoveValue.Text, 1) = "-" Then
            EBXAxesMoveValue.Text = EBXAxesMoveValue.Text.Substring(1, EBXAxesMoveValue.Text.Length - 1)
        Else
            EBXAxesMoveValue.Text = "-" + EBXAxesMoveValue.Text
        End If
    End Sub

    Private Sub PlusMinesY_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles PlusMinesY.Click

        If Left(EBYValue.Text, 1) = "-" Then
            EBYValue.Text = EBYValue.Text.Substring(1, EBYValue.Text.Length - 1)
        Else
            EBYValue.Text = "-" + EBYValue.Text
        End If
    End Sub

    Private Sub OnOff_persent_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles OnOff_persent.Click
        If Right(EBXAxesMoveValue.Text, 1) = "%" Then
            EBXAxesMoveValue.Text = EBXAxesMoveValue.Text.Substring(0, EBXAxesMoveValue.Text.Length - 1)
        Else
            EBXAxesMoveValue.Text = EBXAxesMoveValue.Text + "%"
        End If
    End Sub

    Private Sub CurrentXmax_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles CurrentXmax.Click
        BSetXmaxValue.Text = CurrentXmaxSub().Replace(",", ".")
    End Sub

    Private Sub CurrentValue_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles CurrentValue.Click
        If DDTime.Items(DDTime.SelectedItemIndex).Tag = 0 Then
            MultTime = 1
        Else
            MultTime = Val(DDTime.Items(DDTime.SelectedItemIndex).Tag) * (1 / 86400)
        End If
        BSetValue.Text = CurrentValueSub().Replace(",", ".")
    End Sub

    Private Sub CurrentXmin_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles CurrentXmin.Click
        BSetXminValue.Text = CurrentXminSub().Replace(",", ".")
    End Sub

    Private Sub UpPoint_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles UpPoint.Click
        Call BuckUpSet()
        Call PointUp()
    End Sub

    Private Sub DownPoint_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles DownPoint.Click
        Call BuckUpSet()
        Call PointDown()
    End Sub

    Private Sub OnOff_persentY_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles OnOff_persentY.Click
        If Right(EBYValue.Text, 1) = "%" Then
            EBYValue.Text = EBYValue.Text.Substring(0, EBYValue.Text.Length - 1)
        Else
            EBYValue.Text = EBYValue.Text + "%"
        End If
    End Sub

    Private Sub Button2_Click_1(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles Button2.Click
        Call BuckUpSet()
        YValue = Val(EBYValue.Text)

        'EBYminValue.Text = EBYmaxValue.Text
        If Right(EBYValue.Text, 1) = "%" Then
            Call YShiftSub("%")
        Else
            Call YShiftSub("_")
        End If
    End Sub

    Private Sub CurrentYmax_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles CurrentYmax.Click
        BSetYmaxValue.Text = CurrentYmaxSub().Replace(",", ".")
    End Sub

    Private Sub CurrentYmin_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles CurrentYmin.Click
        BSetYminValue.Text = CurrentYminSub().Replace(",", ".")
    End Sub

    Private Sub CurrentYstep_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles CurrentYstep.Click
        BSetYStepValue.Text = CurrentYstepSub().Replace(",", ".")
    End Sub

    Private Sub BuckUp_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles BuckUp.Click
        Call BuckUpSub()
    End Sub

    Private Sub AutoX_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles AutoX.Click
        Call AutoXSub()
    End Sub

    

    Private Sub AutoY_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles AutoY.Click
        Call AutoYSub()
    End Sub

    

    Private Sub DropDown1_SelectionChanged(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles DDSubLinesY.SelectionChanged, DropDown1.SelectionChanged

    End Sub


    Private Sub DDSubLinesY_SelectionChanged_1(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles DDSubLinesY.SelectionChanged
        Call SubLineY_Sub(Val(DDSubLinesY.Items(DDSubLinesY.SelectedItemIndex).Tag))
    End Sub

   

End Class
