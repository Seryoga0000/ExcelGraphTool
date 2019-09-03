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
        Dim RibbonDialogLauncherImpl1 As Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher = Me.Factory.CreateRibbonDialogLauncher
        Dim RibbonDropDownItemImpl1 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl2 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl3 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl4 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl5 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl6 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl7 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl8 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl9 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl10 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl11 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl12 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl13 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl14 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.TabGraph = Me.Factory.CreateRibbonTab
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.IncreaseGraphButton = Me.Factory.CreateRibbonButton
        Me.PercentIncrease = Me.Factory.CreateRibbonEditBox
        Me.ReduceGraphButton = Me.Factory.CreateRibbonButton
        Me.BXAxesRight = Me.Factory.CreateRibbonButton
        Me.EBXAxesMoveValue = Me.Factory.CreateRibbonEditBox
        Me.BXAxesLeft = Me.Factory.CreateRibbonButton
        Me.XminIncrease = Me.Factory.CreateRibbonButton
        Me.EBXMoveValue = Me.Factory.CreateRibbonEditBox
        Me.XminReduce = Me.Factory.CreateRibbonButton
        Me.XmaxIncrease = Me.Factory.CreateRibbonButton
        Me.Label1 = Me.Factory.CreateRibbonLabel
        Me.XmaxReduce = Me.Factory.CreateRibbonButton
        Me.Separator1 = Me.Factory.CreateRibbonSeparator
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.EBStartPoint = Me.Factory.CreateRibbonEditBox
        Me.Label3 = Me.Factory.CreateRibbonLabel
        Me.Button4 = Me.Factory.CreateRibbonButton
        Me.EBEndPoint = Me.Factory.CreateRibbonEditBox
        Me.Label2 = Me.Factory.CreateRibbonLabel
        Me.Separator3 = Me.Factory.CreateRibbonSeparator
        Me.YIncrease = Me.Factory.CreateRibbonButton
        Me.EBYValue = Me.Factory.CreateRibbonEditBox
        Me.YReduce = Me.Factory.CreateRibbonButton
        Me.EBYmaxValue = Me.Factory.CreateRibbonEditBox
        Me.YmaxIncrease = Me.Factory.CreateRibbonButton
        Me.YmaxReduce = Me.Factory.CreateRibbonButton
        Me.YminIncrease = Me.Factory.CreateRibbonButton
        Me.YminReduce = Me.Factory.CreateRibbonButton
        Me.EBYminValue = Me.Factory.CreateRibbonEditBox
        Me.Button6 = Me.Factory.CreateRibbonButton
        Me.Separator2 = Me.Factory.CreateRibbonSeparator
        Me.BSet = Me.Factory.CreateRibbonButton
        Me.DDTime = Me.Factory.CreateRibbonDropDown
        Me.DDSubLines = Me.Factory.CreateRibbonDropDown
        Me.Tab1.SuspendLayout()
        Me.TabGraph.SuspendLayout()
        Me.Group2.SuspendLayout()
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
        Me.TabGraph.Label = "График"
        Me.TabGraph.Name = "TabGraph"
        '
        'Group2
        '
        Me.Group2.DialogLauncher = RibbonDialogLauncherImpl1
        Me.Group2.Items.Add(Me.IncreaseGraphButton)
        Me.Group2.Items.Add(Me.PercentIncrease)
        Me.Group2.Items.Add(Me.ReduceGraphButton)
        Me.Group2.Items.Add(Me.BXAxesRight)
        Me.Group2.Items.Add(Me.EBXAxesMoveValue)
        Me.Group2.Items.Add(Me.BXAxesLeft)
        Me.Group2.Items.Add(Me.XminIncrease)
        Me.Group2.Items.Add(Me.EBXMoveValue)
        Me.Group2.Items.Add(Me.XminReduce)
        Me.Group2.Items.Add(Me.XmaxIncrease)
        Me.Group2.Items.Add(Me.Label1)
        Me.Group2.Items.Add(Me.XmaxReduce)
        Me.Group2.Items.Add(Me.Separator2)
        Me.Group2.Items.Add(Me.DDTime)
        Me.Group2.Items.Add(Me.BSet)
        Me.Group2.Items.Add(Me.DDSubLines)
        Me.Group2.Items.Add(Me.Separator1)
        Me.Group2.Items.Add(Me.Button3)
        Me.Group2.Items.Add(Me.EBStartPoint)
        Me.Group2.Items.Add(Me.Label3)
        Me.Group2.Items.Add(Me.Button4)
        Me.Group2.Items.Add(Me.EBEndPoint)
        Me.Group2.Items.Add(Me.Label2)
        Me.Group2.Items.Add(Me.Separator3)
        Me.Group2.Items.Add(Me.YIncrease)
        Me.Group2.Items.Add(Me.EBYValue)
        Me.Group2.Items.Add(Me.YReduce)
        Me.Group2.Items.Add(Me.EBYmaxValue)
        Me.Group2.Items.Add(Me.YmaxIncrease)
        Me.Group2.Items.Add(Me.YmaxReduce)
        Me.Group2.Items.Add(Me.YminIncrease)
        Me.Group2.Items.Add(Me.YminReduce)
        Me.Group2.Items.Add(Me.EBYminValue)
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
        'BXAxesRight
        '
        Me.BXAxesRight.Label = "   -> X ->"
        Me.BXAxesRight.Name = "BXAxesRight"
        '
        'EBXAxesMoveValue
        '
        Me.EBXAxesMoveValue.Label = " "
        Me.EBXAxesMoveValue.Name = "EBXAxesMoveValue"
        Me.EBXAxesMoveValue.Text = "10"
        '
        'BXAxesLeft
        '
        Me.BXAxesLeft.Label = "   <- X <-"
        Me.BXAxesLeft.Name = "BXAxesLeft"
        '
        'XminIncrease
        '
        Me.XminIncrease.Label = "    Xmin ->"
        Me.XminIncrease.Name = "XminIncrease"
        '
        'EBXMoveValue
        '
        Me.EBXMoveValue.Label = "            "
        Me.EBXMoveValue.Name = "EBXMoveValue"
        Me.EBXMoveValue.Text = "10"
        '
        'XminReduce
        '
        Me.XminReduce.Label = "    Xmin <-"
        Me.XminReduce.Name = "XminReduce"
        '
        'XmaxIncrease
        '
        Me.XmaxIncrease.Label = "-> Xmax"
        Me.XmaxIncrease.Name = "XmaxIncrease"
        '
        'Label1
        '
        Me.Label1.Label = " "
        Me.Label1.Name = "Label1"
        '
        'XmaxReduce
        '
        Me.XmaxReduce.Label = "<- Xmax"
        Me.XmaxReduce.Name = "XmaxReduce"
        '
        'Separator1
        '
        Me.Separator1.Name = "Separator1"
        '
        'Button3
        '
        Me.Button3.Label = "StartPoint"
        Me.Button3.Name = "Button3"
        '
        'EBStartPoint
        '
        Me.EBStartPoint.Enabled = False
        Me.EBStartPoint.Label = " "
        Me.EBStartPoint.Name = "EBStartPoint"
        Me.EBStartPoint.Text = "0"
        '
        'Label3
        '
        Me.Label3.Label = " "
        Me.Label3.Name = "Label3"
        '
        'Button4
        '
        Me.Button4.Label = "EndPoint"
        Me.Button4.Name = "Button4"
        '
        'EBEndPoint
        '
        Me.EBEndPoint.Enabled = False
        Me.EBEndPoint.Label = " "
        Me.EBEndPoint.Name = "EBEndPoint"
        Me.EBEndPoint.Text = "0"
        '
        'Label2
        '
        Me.Label2.Label = " "
        Me.Label2.Name = "Label2"
        '
        'Separator3
        '
        Me.Separator3.Name = "Separator3"
        '
        'YIncrease
        '
        Me.YIncrease.Label = "    <- Y ->"
        Me.YIncrease.Name = "YIncrease"
        '
        'EBYValue
        '
        Me.EBYValue.Label = " "
        Me.EBYValue.Name = "EBYValue"
        Me.EBYValue.Text = "1"
        '
        'YReduce
        '
        Me.YReduce.Label = "    -> Y <-"
        Me.YReduce.Name = "YReduce"
        '
        'EBYmaxValue
        '
        Me.EBYmaxValue.Label = " "
        Me.EBYmaxValue.Name = "EBYmaxValue"
        Me.EBYmaxValue.Tag = ""
        Me.EBYmaxValue.Text = "100"
        '
        'YmaxIncrease
        '
        Me.YmaxIncrease.Label = "    Ymax ->"
        Me.YmaxIncrease.Name = "YmaxIncrease"
        '
        'YmaxReduce
        '
        Me.YmaxReduce.Label = "    Ymax <-"
        Me.YmaxReduce.Name = "YmaxReduce"
        '
        'YminIncrease
        '
        Me.YminIncrease.Label = "    Ymin ->"
        Me.YminIncrease.Name = "YminIncrease"
        '
        'YminReduce
        '
        Me.YminReduce.Label = "    Ymin <-"
        Me.YminReduce.Name = "YminReduce"
        '
        'EBYminValue
        '
        Me.EBYminValue.Label = " "
        Me.EBYminValue.Name = "EBYminValue"
        Me.EBYminValue.Text = "100"
        '
        'Button6
        '
        Me.Button6.Label = "    Ymin ->"
        Me.Button6.Name = "Button6"
        '
        'Separator2
        '
        Me.Separator2.Name = "Separator2"
        '
        'BSet
        '
        Me.BSet.Label = "Set"
        Me.BSet.Name = "BSet"
        '
        'DDTime
        '
        RibbonDropDownItemImpl1.Label = " "
        RibbonDropDownItemImpl1.Tag = "0"
        RibbonDropDownItemImpl2.Label = "Sec"
        RibbonDropDownItemImpl2.Tag = "1"
        RibbonDropDownItemImpl3.Label = "Min"
        RibbonDropDownItemImpl3.Tag = "60"
        RibbonDropDownItemImpl4.Label = "Hour"
        RibbonDropDownItemImpl4.Tag = "3600"
        RibbonDropDownItemImpl5.Label = "Day"
        RibbonDropDownItemImpl5.Tag = "86400"
        Me.DDTime.Items.Add(RibbonDropDownItemImpl1)
        Me.DDTime.Items.Add(RibbonDropDownItemImpl2)
        Me.DDTime.Items.Add(RibbonDropDownItemImpl3)
        Me.DDTime.Items.Add(RibbonDropDownItemImpl4)
        Me.DDTime.Items.Add(RibbonDropDownItemImpl5)
        Me.DDTime.Label = " XTime x"
        Me.DDTime.Name = "DDTime"
        '
        'DDSubLines
        '
        RibbonDropDownItemImpl6.Label = "2"
        RibbonDropDownItemImpl6.Tag = "2"
        RibbonDropDownItemImpl7.Label = "3"
        RibbonDropDownItemImpl7.Tag = "3"
        RibbonDropDownItemImpl8.Label = "4"
        RibbonDropDownItemImpl8.Tag = "4"
        RibbonDropDownItemImpl9.Label = "5"
        RibbonDropDownItemImpl9.Tag = "5"
        RibbonDropDownItemImpl10.Label = "6"
        RibbonDropDownItemImpl10.Tag = "6"
        RibbonDropDownItemImpl11.Label = "8"
        RibbonDropDownItemImpl11.Tag = "8"
        RibbonDropDownItemImpl12.Label = "10"
        RibbonDropDownItemImpl12.Tag = "10"
        RibbonDropDownItemImpl13.Label = "12"
        RibbonDropDownItemImpl13.Tag = "12"
        RibbonDropDownItemImpl14.Label = "24"
        RibbonDropDownItemImpl14.Tag = "24"
        Me.DDSubLines.Items.Add(RibbonDropDownItemImpl6)
        Me.DDSubLines.Items.Add(RibbonDropDownItemImpl7)
        Me.DDSubLines.Items.Add(RibbonDropDownItemImpl8)
        Me.DDSubLines.Items.Add(RibbonDropDownItemImpl9)
        Me.DDSubLines.Items.Add(RibbonDropDownItemImpl10)
        Me.DDSubLines.Items.Add(RibbonDropDownItemImpl11)
        Me.DDSubLines.Items.Add(RibbonDropDownItemImpl12)
        Me.DDSubLines.Items.Add(RibbonDropDownItemImpl13)
        Me.DDSubLines.Items.Add(RibbonDropDownItemImpl14)
        Me.DDSubLines.Label = "SubLine"
        Me.DDSubLines.Name = "DDSubLines"
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
    Friend WithEvents Separator1 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents EBStartPoint As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents Button4 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents EBEndPoint As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents Label2 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Label3 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Separator3 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents YmaxIncrease As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents YminReduce As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents EBYminValue As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents YmaxReduce As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents YminIncrease As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents EBYmaxValue As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents BXAxesRight As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents EBXAxesMoveValue As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents BXAxesLeft As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents XminIncrease As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents EBXMoveValue As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents XminReduce As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents XmaxIncrease As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Label1 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents XmaxReduce As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents YIncrease As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents EBYValue As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents YReduce As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button6 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator2 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents DDTime As Microsoft.Office.Tools.Ribbon.RibbonDropDown
    Friend WithEvents BSet As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents DDSubLines As Microsoft.Office.Tools.Ribbon.RibbonDropDown
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Lenta
        Get
            Return Me.GetRibbon(Of Lenta)()
        End Get
    End Property
End Class
