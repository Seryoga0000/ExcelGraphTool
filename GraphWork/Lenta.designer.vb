Partial Class Lenta
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Требуется для поддержки конструктора композиции классов Windows.Forms
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'Этот вызов установлен конструктором компонентов.
        InitializeComponent()

    End Sub

    'Компонент переопределяет метод dispose для очистки списка элементов.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Является обязательной для конструктора компонентов
    Private components As System.ComponentModel.IContainer

    'ПРИМЕЧАНИЕ. Следующая процедура является обязательной для конструктора компонентов
    'Для ее изменения используйте конструктор компонентов.
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim RibbonDropDownItemImpl1 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl2 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl3 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl4 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl5 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl6 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl7 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl8 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl9 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDialogLauncherImpl1 As Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher = Me.Factory.CreateRibbonDialogLauncher
        Dim RibbonDropDownItemImpl10 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl11 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl12 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl13 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl14 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl15 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl16 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl17 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl18 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl19 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl20 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl21 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl22 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl23 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl24 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl25 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl26 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl27 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl28 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl29 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl30 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl31 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl32 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Me.Button6 = Me.Factory.CreateRibbonButton
        Me.DropDown1 = Me.Factory.CreateRibbonDropDown
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.TabGraph = Me.Factory.CreateRibbonTab
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.IncreaseGraphButton = Me.Factory.CreateRibbonButton
        Me.PercentIncrease = Me.Factory.CreateRibbonEditBox
        Me.ReduceGraphButton = Me.Factory.CreateRibbonButton
        Me.Separator2 = Me.Factory.CreateRibbonSeparator
        Me.XminIncrease = Me.Factory.CreateRibbonButton
        Me.PlusMinesX = Me.Factory.CreateRibbonButton
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.BXAxesRight = Me.Factory.CreateRibbonButton
        Me.EBXAxesMoveValue = Me.Factory.CreateRibbonEditBox
        Me.OnOff_persent = Me.Factory.CreateRibbonButton
        Me.XmaxIncrease = Me.Factory.CreateRibbonButton
        Me.DDTime = Me.Factory.CreateRibbonDropDown
        Me.Button4 = Me.Factory.CreateRibbonButton
        Me.Separator4 = Me.Factory.CreateRibbonSeparator
        Me.DDSubLines = Me.Factory.CreateRibbonDropDown
        Me.AutoX = Me.Factory.CreateRibbonButton
        Me.Label2 = Me.Factory.CreateRibbonLabel
        Me.Label4 = Me.Factory.CreateRibbonLabel
        Me.Separator5 = Me.Factory.CreateRibbonSeparator
        Me.BSetXmax = Me.Factory.CreateRibbonButton
        Me.BSet = Me.Factory.CreateRibbonButton
        Me.BSetXmin = Me.Factory.CreateRibbonButton
        Me.BSetXmaxValue = Me.Factory.CreateRibbonEditBox
        Me.BSetValue = Me.Factory.CreateRibbonEditBox
        Me.BSetXminValue = Me.Factory.CreateRibbonEditBox
        Me.CurrentXmax = Me.Factory.CreateRibbonButton
        Me.CurrentValue = Me.Factory.CreateRibbonButton
        Me.CurrentXmin = Me.Factory.CreateRibbonButton
        Me.Label1 = Me.Factory.CreateRibbonLabel
        Me.Label3 = Me.Factory.CreateRibbonLabel
        Me.Separator3 = Me.Factory.CreateRibbonSeparator
        Me.YmaxIncrease = Me.Factory.CreateRibbonButton
        Me.PlusMinesY = Me.Factory.CreateRibbonButton
        Me.YminIncrease = Me.Factory.CreateRibbonButton
        Me.Button2 = Me.Factory.CreateRibbonButton
        Me.EBYValue = Me.Factory.CreateRibbonEditBox
        Me.OnOff_persentY = Me.Factory.CreateRibbonButton
        Me.UpPoint = Me.Factory.CreateRibbonButton
        Me.YIncrease = Me.Factory.CreateRibbonButton
        Me.DownPoint = Me.Factory.CreateRibbonButton
        Me.Separator1 = Me.Factory.CreateRibbonSeparator
        Me.BSetYmax = Me.Factory.CreateRibbonButton
        Me.BSetYStep = Me.Factory.CreateRibbonButton
        Me.BSetYmin = Me.Factory.CreateRibbonButton
        Me.BSetYmaxValue = Me.Factory.CreateRibbonEditBox
        Me.BSetYStepValue = Me.Factory.CreateRibbonEditBox
        Me.BSetYminValue = Me.Factory.CreateRibbonEditBox
        Me.CurrentYmax = Me.Factory.CreateRibbonButton
        Me.CurrentYstep = Me.Factory.CreateRibbonButton
        Me.CurrentYmin = Me.Factory.CreateRibbonButton
        Me.Separator6 = Me.Factory.CreateRibbonSeparator
        Me.DDSubLinesY = Me.Factory.CreateRibbonDropDown
        Me.AutoY = Me.Factory.CreateRibbonButton
        Me.Separator7 = Me.Factory.CreateRibbonSeparator
        Me.BuckUp = Me.Factory.CreateRibbonButton
        Me.Help = Me.Factory.CreateRibbonButton
        Me.Label5 = Me.Factory.CreateRibbonLabel
        Me.Tab1.SuspendLayout()
        Me.TabGraph.SuspendLayout()
        Me.Group2.SuspendLayout()
        '
        'Button6
        '
        Me.Button6.Label = "    Ymin ->"
        Me.Button6.Name = "Button6"
        '
        'DropDown1
        '
        RibbonDropDownItemImpl1.Label = "2"
        RibbonDropDownItemImpl1.Tag = "2"
        RibbonDropDownItemImpl2.Label = "3"
        RibbonDropDownItemImpl2.Tag = "3"
        RibbonDropDownItemImpl3.Label = "4"
        RibbonDropDownItemImpl3.Tag = "4"
        RibbonDropDownItemImpl4.Label = "5"
        RibbonDropDownItemImpl4.Tag = "5"
        RibbonDropDownItemImpl5.Label = "6"
        RibbonDropDownItemImpl5.Tag = "6"
        RibbonDropDownItemImpl6.Label = "8"
        RibbonDropDownItemImpl6.Tag = "8"
        RibbonDropDownItemImpl7.Label = "10"
        RibbonDropDownItemImpl7.Tag = "10"
        RibbonDropDownItemImpl8.Label = "12"
        RibbonDropDownItemImpl8.Tag = "12"
        RibbonDropDownItemImpl9.Label = "24"
        RibbonDropDownItemImpl9.Tag = "24"
        Me.DropDown1.Items.Add(RibbonDropDownItemImpl1)
        Me.DropDown1.Items.Add(RibbonDropDownItemImpl2)
        Me.DropDown1.Items.Add(RibbonDropDownItemImpl3)
        Me.DropDown1.Items.Add(RibbonDropDownItemImpl4)
        Me.DropDown1.Items.Add(RibbonDropDownItemImpl5)
        Me.DropDown1.Items.Add(RibbonDropDownItemImpl6)
        Me.DropDown1.Items.Add(RibbonDropDownItemImpl7)
        Me.DropDown1.Items.Add(RibbonDropDownItemImpl8)
        Me.DropDown1.Items.Add(RibbonDropDownItemImpl9)
        Me.DropDown1.Label = "SubLineY"
        Me.DropDown1.Name = "DropDown1"
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Label = "TabAddIns"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Label = "Group1"
        Me.Group1.Name = "Group1"
        '
        'TabGraph
        '
        Me.TabGraph.Groups.Add(Me.Group2)
        Me.TabGraph.Label = "Graph"
        Me.TabGraph.Name = "TabGraph"
        '
        'Group2
        '
        Me.Group2.DialogLauncher = RibbonDialogLauncherImpl1
        Me.Group2.Items.Add(Me.IncreaseGraphButton)
        Me.Group2.Items.Add(Me.PercentIncrease)
        Me.Group2.Items.Add(Me.ReduceGraphButton)
        Me.Group2.Items.Add(Me.Separator2)
        Me.Group2.Items.Add(Me.XminIncrease)
        Me.Group2.Items.Add(Me.PlusMinesX)
        Me.Group2.Items.Add(Me.Button3)
        Me.Group2.Items.Add(Me.BXAxesRight)
        Me.Group2.Items.Add(Me.EBXAxesMoveValue)
        Me.Group2.Items.Add(Me.OnOff_persent)
        Me.Group2.Items.Add(Me.XmaxIncrease)
        Me.Group2.Items.Add(Me.DDTime)
        Me.Group2.Items.Add(Me.Button4)
        Me.Group2.Items.Add(Me.Separator4)
        Me.Group2.Items.Add(Me.DDSubLines)
        Me.Group2.Items.Add(Me.AutoX)
        Me.Group2.Items.Add(Me.Label2)
        Me.Group2.Items.Add(Me.Label4)
        Me.Group2.Items.Add(Me.Separator5)
        Me.Group2.Items.Add(Me.BSetXmax)
        Me.Group2.Items.Add(Me.BSet)
        Me.Group2.Items.Add(Me.BSetXmin)
        Me.Group2.Items.Add(Me.BSetXmaxValue)
        Me.Group2.Items.Add(Me.BSetValue)
        Me.Group2.Items.Add(Me.BSetXminValue)
        Me.Group2.Items.Add(Me.CurrentXmax)
        Me.Group2.Items.Add(Me.CurrentValue)
        Me.Group2.Items.Add(Me.CurrentXmin)
        Me.Group2.Items.Add(Me.Label1)
        Me.Group2.Items.Add(Me.Label3)
        Me.Group2.Items.Add(Me.Separator3)
        Me.Group2.Items.Add(Me.YmaxIncrease)
        Me.Group2.Items.Add(Me.PlusMinesY)
        Me.Group2.Items.Add(Me.YminIncrease)
        Me.Group2.Items.Add(Me.Button2)
        Me.Group2.Items.Add(Me.EBYValue)
        Me.Group2.Items.Add(Me.OnOff_persentY)
        Me.Group2.Items.Add(Me.UpPoint)
        Me.Group2.Items.Add(Me.YIncrease)
        Me.Group2.Items.Add(Me.DownPoint)
        Me.Group2.Items.Add(Me.Separator1)
        Me.Group2.Items.Add(Me.BSetYmax)
        Me.Group2.Items.Add(Me.BSetYStep)
        Me.Group2.Items.Add(Me.BSetYmin)
        Me.Group2.Items.Add(Me.BSetYmaxValue)
        Me.Group2.Items.Add(Me.BSetYStepValue)
        Me.Group2.Items.Add(Me.BSetYminValue)
        Me.Group2.Items.Add(Me.CurrentYmax)
        Me.Group2.Items.Add(Me.CurrentYstep)
        Me.Group2.Items.Add(Me.CurrentYmin)
        Me.Group2.Items.Add(Me.Separator6)
        Me.Group2.Items.Add(Me.DDSubLinesY)
        Me.Group2.Items.Add(Me.AutoY)
        Me.Group2.Items.Add(Me.Separator7)
        Me.Group2.Items.Add(Me.BuckUp)
        Me.Group2.Items.Add(Me.Label5)
        Me.Group2.Items.Add(Me.Help)
        Me.Group2.Label = "Graph"
        Me.Group2.Name = "Group2"
        '
        'IncreaseGraphButton
        '
        Me.IncreaseGraphButton.Label = "     <- X ->"
        Me.IncreaseGraphButton.Name = "IncreaseGraphButton"
        '
        'PercentIncrease
        '
        Me.PercentIncrease.Label = "%"
        Me.PercentIncrease.Name = "PercentIncrease"
        Me.PercentIncrease.Text = "10"
        '
        'ReduceGraphButton
        '
        Me.ReduceGraphButton.Label = "     -> X <-"
        Me.ReduceGraphButton.Name = "ReduceGraphButton"
        '
        'Separator2
        '
        Me.Separator2.Name = "Separator2"
        '
        'XminIncrease
        '
        Me.XminIncrease.Label = "->Xmin     "
        Me.XminIncrease.Name = "XminIncrease"
        '
        'PlusMinesX
        '
        Me.PlusMinesX.Label = "    +/-        "
        Me.PlusMinesX.Name = "PlusMinesX"
        '
        'Button3
        '
        Me.Button3.Label = "StartPoint"
        Me.Button3.Name = "Button3"
        '
        'BXAxesRight
        '
        Me.BXAxesRight.Label = "       -> X ->    "
        Me.BXAxesRight.Name = "BXAxesRight"
        '
        'EBXAxesMoveValue
        '
        Me.EBXAxesMoveValue.Label = " "
        Me.EBXAxesMoveValue.Name = "EBXAxesMoveValue"
        Me.EBXAxesMoveValue.ShowLabel = False
        Me.EBXAxesMoveValue.Text = "10%"
        '
        'OnOff_persent
        '
        Me.OnOff_persent.Label = "          %          "
        Me.OnOff_persent.Name = "OnOff_persent"
        '
        'XmaxIncrease
        '
        Me.XmaxIncrease.Label = "    ->Xmax                  "
        Me.XmaxIncrease.Name = "XmaxIncrease"
        '
        'DDTime
        '
        RibbonDropDownItemImpl10.Label = " "
        RibbonDropDownItemImpl10.Tag = "0"
        RibbonDropDownItemImpl11.Label = "Sec"
        RibbonDropDownItemImpl11.Tag = "1"
        RibbonDropDownItemImpl12.Label = "Min"
        RibbonDropDownItemImpl12.Tag = "60"
        RibbonDropDownItemImpl13.Label = "Hour"
        RibbonDropDownItemImpl13.Tag = "3600"
        RibbonDropDownItemImpl14.Label = "Day"
        RibbonDropDownItemImpl14.Tag = "86400"
        Me.DDTime.Items.Add(RibbonDropDownItemImpl10)
        Me.DDTime.Items.Add(RibbonDropDownItemImpl11)
        Me.DDTime.Items.Add(RibbonDropDownItemImpl12)
        Me.DDTime.Items.Add(RibbonDropDownItemImpl13)
        Me.DDTime.Items.Add(RibbonDropDownItemImpl14)
        Me.DDTime.Label = " x "
        Me.DDTime.Name = "DDTime"
        '
        'Button4
        '
        Me.Button4.Label = "    EndPoint               "
        Me.Button4.Name = "Button4"
        '
        'Separator4
        '
        Me.Separator4.Name = "Separator4"
        '
        'DDSubLines
        '
        RibbonDropDownItemImpl15.Label = "2"
        RibbonDropDownItemImpl15.Tag = "2"
        RibbonDropDownItemImpl16.Label = "3"
        RibbonDropDownItemImpl16.Tag = "3"
        RibbonDropDownItemImpl17.Label = "4"
        RibbonDropDownItemImpl17.Tag = "4"
        RibbonDropDownItemImpl18.Label = "5"
        RibbonDropDownItemImpl18.Tag = "5"
        RibbonDropDownItemImpl19.Label = "6"
        RibbonDropDownItemImpl19.Tag = "6"
        RibbonDropDownItemImpl20.Label = "8"
        RibbonDropDownItemImpl20.Tag = "8"
        RibbonDropDownItemImpl21.Label = "10"
        RibbonDropDownItemImpl21.Tag = "10"
        RibbonDropDownItemImpl22.Label = "12"
        RibbonDropDownItemImpl22.Tag = "12"
        RibbonDropDownItemImpl23.Label = "24"
        RibbonDropDownItemImpl23.Tag = "24"
        Me.DDSubLines.Items.Add(RibbonDropDownItemImpl15)
        Me.DDSubLines.Items.Add(RibbonDropDownItemImpl16)
        Me.DDSubLines.Items.Add(RibbonDropDownItemImpl17)
        Me.DDSubLines.Items.Add(RibbonDropDownItemImpl18)
        Me.DDSubLines.Items.Add(RibbonDropDownItemImpl19)
        Me.DDSubLines.Items.Add(RibbonDropDownItemImpl20)
        Me.DDSubLines.Items.Add(RibbonDropDownItemImpl21)
        Me.DDSubLines.Items.Add(RibbonDropDownItemImpl22)
        Me.DDSubLines.Items.Add(RibbonDropDownItemImpl23)
        Me.DDSubLines.Label = "SubLineX"
        Me.DDSubLines.Name = "DDSubLines"
        '
        'AutoX
        '
        Me.AutoX.Label = "AutoX"
        Me.AutoX.Name = "AutoX"
        '
        'Label2
        '
        Me.Label2.Label = " "
        Me.Label2.Name = "Label2"
        '
        'Label4
        '
        Me.Label4.Label = " "
        Me.Label4.Name = "Label4"
        '
        'Separator5
        '
        Me.Separator5.Name = "Separator5"
        '
        'BSetXmax
        '
        Me.BSetXmax.Label = "SetXmax"
        Me.BSetXmax.Name = "BSetXmax"
        '
        'BSet
        '
        Me.BSet.Label = "SetStep"
        Me.BSet.Name = "BSet"
        '
        'BSetXmin
        '
        Me.BSetXmin.Label = "SetXmin"
        Me.BSetXmin.Name = "BSetXmin"
        '
        'BSetXmaxValue
        '
        Me.BSetXmaxValue.Label = " "
        Me.BSetXmaxValue.Name = "BSetXmaxValue"
        Me.BSetXmaxValue.ShowLabel = False
        Me.BSetXmaxValue.Text = Nothing
        '
        'BSetValue
        '
        Me.BSetValue.Label = " "
        Me.BSetValue.Name = "BSetValue"
        Me.BSetValue.ShowLabel = False
        Me.BSetValue.Text = Nothing
        '
        'BSetXminValue
        '
        Me.BSetXminValue.Label = " "
        Me.BSetXminValue.Name = "BSetXminValue"
        Me.BSetXminValue.ShowLabel = False
        Me.BSetXminValue.Text = Nothing
        '
        'CurrentXmax
        '
        Me.CurrentXmax.Label = "..."
        Me.CurrentXmax.Name = "CurrentXmax"
        '
        'CurrentValue
        '
        Me.CurrentValue.Label = "..."
        Me.CurrentValue.Name = "CurrentValue"
        '
        'CurrentXmin
        '
        Me.CurrentXmin.Label = "..."
        Me.CurrentXmin.Name = "CurrentXmin"
        '
        'Label1
        '
        Me.Label1.Label = " "
        Me.Label1.Name = "Label1"
        '
        'Label3
        '
        Me.Label3.Label = " "
        Me.Label3.Name = "Label3"
        '
        'Separator3
        '
        Me.Separator3.Name = "Separator3"
        '
        'YmaxIncrease
        '
        Me.YmaxIncrease.Label = "->Ymax"
        Me.YmaxIncrease.Name = "YmaxIncrease"
        '
        'PlusMinesY
        '
        Me.PlusMinesY.Label = "     +/- "
        Me.PlusMinesY.Name = "PlusMinesY"
        '
        'YminIncrease
        '
        Me.YminIncrease.Label = "->Ymin"
        Me.YminIncrease.Name = "YminIncrease"
        '
        'Button2
        '
        Me.Button2.Label = "     -> Y ->         "
        Me.Button2.Name = "Button2"
        '
        'EBYValue
        '
        Me.EBYValue.Label = " "
        Me.EBYValue.Name = "EBYValue"
        Me.EBYValue.Text = "10%"
        '
        'OnOff_persentY
        '
        Me.OnOff_persentY.Label = "           %            "
        Me.OnOff_persentY.Name = "OnOff_persentY"
        '
        'UpPoint
        '
        Me.UpPoint.Label = "UpPoint      "
        Me.UpPoint.Name = "UpPoint"
        '
        'YIncrease
        '
        Me.YIncrease.Label = "    <- Y ->    "
        Me.YIncrease.Name = "YIncrease"
        '
        'DownPoint
        '
        Me.DownPoint.Label = "DownPoint"
        Me.DownPoint.Name = "DownPoint"
        '
        'Separator1
        '
        Me.Separator1.Name = "Separator1"
        '
        'BSetYmax
        '
        Me.BSetYmax.Label = "SetYmax"
        Me.BSetYmax.Name = "BSetYmax"
        '
        'BSetYStep
        '
        Me.BSetYStep.Label = "SetYStep"
        Me.BSetYStep.Name = "BSetYStep"
        '
        'BSetYmin
        '
        Me.BSetYmin.Label = "SetYmin"
        Me.BSetYmin.Name = "BSetYmin"
        '
        'BSetYmaxValue
        '
        Me.BSetYmaxValue.Label = " "
        Me.BSetYmaxValue.Name = "BSetYmaxValue"
        Me.BSetYmaxValue.ShowLabel = False
        Me.BSetYmaxValue.Text = Nothing
        '
        'BSetYStepValue
        '
        Me.BSetYStepValue.Label = " "
        Me.BSetYStepValue.Name = "BSetYStepValue"
        Me.BSetYStepValue.ShowLabel = False
        Me.BSetYStepValue.Text = Nothing
        '
        'BSetYminValue
        '
        Me.BSetYminValue.Label = " "
        Me.BSetYminValue.Name = "BSetYminValue"
        Me.BSetYminValue.ShowLabel = False
        Me.BSetYminValue.Text = Nothing
        '
        'CurrentYmax
        '
        Me.CurrentYmax.Label = "..."
        Me.CurrentYmax.Name = "CurrentYmax"
        '
        'CurrentYstep
        '
        Me.CurrentYstep.Label = "..."
        Me.CurrentYstep.Name = "CurrentYstep"
        '
        'CurrentYmin
        '
        Me.CurrentYmin.Label = "..."
        Me.CurrentYmin.Name = "CurrentYmin"
        '
        'Separator6
        '
        Me.Separator6.Name = "Separator6"
        '
        'DDSubLinesY
        '
        RibbonDropDownItemImpl24.Label = "2"
        RibbonDropDownItemImpl24.Tag = "2"
        RibbonDropDownItemImpl25.Label = "3"
        RibbonDropDownItemImpl25.Tag = "3"
        RibbonDropDownItemImpl26.Label = "4"
        RibbonDropDownItemImpl26.Tag = "4"
        RibbonDropDownItemImpl27.Label = "5"
        RibbonDropDownItemImpl27.Tag = "5"
        RibbonDropDownItemImpl28.Label = "6"
        RibbonDropDownItemImpl28.Tag = "6"
        RibbonDropDownItemImpl29.Label = "8"
        RibbonDropDownItemImpl29.Tag = "8"
        RibbonDropDownItemImpl30.Label = "10"
        RibbonDropDownItemImpl30.Tag = "10"
        RibbonDropDownItemImpl31.Label = "12"
        RibbonDropDownItemImpl31.Tag = "12"
        RibbonDropDownItemImpl32.Label = "24"
        RibbonDropDownItemImpl32.Tag = "24"
        Me.DDSubLinesY.Items.Add(RibbonDropDownItemImpl24)
        Me.DDSubLinesY.Items.Add(RibbonDropDownItemImpl25)
        Me.DDSubLinesY.Items.Add(RibbonDropDownItemImpl26)
        Me.DDSubLinesY.Items.Add(RibbonDropDownItemImpl27)
        Me.DDSubLinesY.Items.Add(RibbonDropDownItemImpl28)
        Me.DDSubLinesY.Items.Add(RibbonDropDownItemImpl29)
        Me.DDSubLinesY.Items.Add(RibbonDropDownItemImpl30)
        Me.DDSubLinesY.Items.Add(RibbonDropDownItemImpl31)
        Me.DDSubLinesY.Items.Add(RibbonDropDownItemImpl32)
        Me.DDSubLinesY.Label = "SubLineY"
        Me.DDSubLinesY.Name = "DDSubLinesY"
        '
        'AutoY
        '
        Me.AutoY.Label = "AutoY"
        Me.AutoY.Name = "AutoY"
        '
        'Separator7
        '
        Me.Separator7.Name = "Separator7"
        '
        'BuckUp
        '
        Me.BuckUp.Label = "BuckUp"
        Me.BuckUp.Name = "BuckUp"
        '
        'Help
        '
        Me.Help.Label = "Help"
        Me.Help.Name = "Help"
        '
        'Label5
        '
        Me.Label5.Label = " "
        Me.Label5.Name = "Label5"
        '
        'Lenta
        '
        Me.Name = "Lenta"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tabs.Add(Me.TabGraph)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.TabGraph.ResumeLayout(False)
        Me.TabGraph.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents TabGraph As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents IncreaseGraphButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents PercentIncrease As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents ReduceGraphButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button4 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator3 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents YmaxIncrease As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents YminIncrease As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BXAxesRight As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents XminIncrease As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents XmaxIncrease As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents YIncrease As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents EBYValue As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents Button6 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator2 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents DDTime As Microsoft.Office.Tools.Ribbon.RibbonDropDown
    Friend WithEvents BSet As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents DDSubLines As Microsoft.Office.Tools.Ribbon.RibbonDropDown
    Friend WithEvents BSetYStep As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BSetYmax As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BSetYmin As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BSetXmin As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BSetXmax As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents PlusMinesX As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents PlusMinesY As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents OnOff_persent As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents EBXAxesMoveValue As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents Separator4 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents BSetXmaxValue As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents BSetValue As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents BSetXminValue As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents BSetYmaxValue As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents BSetYStepValue As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents BSetYminValue As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents CurrentXmax As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents CurrentValue As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents CurrentXmin As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Label1 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Label3 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents OnOff_persentY As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents UpPoint As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents DownPoint As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator1 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents Button2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents CurrentYmax As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents CurrentYstep As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents CurrentYmin As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Label2 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Label4 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Separator5 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents Separator6 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents BuckUp As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents AutoX As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents AutoY As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents DDSubLinesY As Microsoft.Office.Tools.Ribbon.RibbonDropDown
    Friend WithEvents Separator7 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents DropDown1 As Microsoft.Office.Tools.Ribbon.RibbonDropDown
    Friend WithEvents Label5 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Help As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Lenta
        Get
            Return Me.GetRibbon(Of Lenta)()
        End Get
    End Property
End Class
