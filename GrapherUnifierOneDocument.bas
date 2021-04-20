Option Explicit

'#Uses "Unify_AxisMinMax_TicksMarks_TickLabels_TitleText.bas"
'#Uses "SupportFunctions.bas"
'#Uses "Unify_PlotDataLimits.bas"

Dim xAxisPattern As New Unify_AxisMinMax_TicksMarks_TickLabels_TitleText
Dim plotLimits As New Unify_PlotDataLimits

Const m_axisLineW As Double = 0.06
Const m_axisColor As grfColor = grfColorBlack
Const m_axisMajorW As Double = 0.04
Const m_axisMajorColor As grfColor = grfColorBlack40
Const m_axisMinorW As Double = 0.04
Const m_axisMinorColor As grfColor = grfColorBlack20

Const m_xAxisPos As Double = 3
Const m_xAxisLength As Double = 16

Const m_fontFace As String = "Calibri"
Const m_fontSizeAxisLabel As Integer = 14
Const m_fontSizeLegend As Integer = 14
Const m_fontSizeAxisTitle As Integer= 16
Const m_fontSizeGraphTitle As Integer= 18
Const m_fontOpacity As Integer = 100
Const m_fontColor As grfColor = grfColorBlack
Const m_fontBold As Boolean = False
Const m_fontItalic As Boolean = False
Const m_fontUnderline As Boolean = False
Const m_fontStrikethrough As Boolean = False

Dim m_plot As AutoPlot
Dim m_clipping As AutoClipping


'**********************************************************************
' Set X axis position = m_xAxisPos And length = m_xAxisLength
' used only for main (first) axis in graph
'**********************************************************************
Sub setXAxisPosition(axis As AutoAxis)
	If 	doubleEq(axis.xPos, m_xAxisPos) = False Or doubleEq(axis.length, m_xAxisLength) = False Then
		axis.xPos = m_xAxisPos
		axis.length = m_xAxisLength
	End If
End Sub

Sub setSolidLine(ByRef l As AutoLine)
	If  l.style <> "Solid" Then
		l.style = "Solid"
	End If
End Sub


'**********************************************************************
' Set axis line width = m_axisLineW
' used for all axes in graph
'**********************************************************************
Sub setAxesLineW(ByRef l As AutoLine)
	If  doubleEq(l.width, m_axisLineW) = False Then
		l.width = m_axisLineW
	End If
	setSolidLine(l)
End Sub

'**********************************************************************
' Set axis: ticks = off, grid (major and minor) = on, grid lines width and color
' used only for main axes in graph (first - X, second - Y)
'**********************************************************************
Sub setMainAxesLine(ByRef axis As AutoAxis)
	Dim l As AutoLine

	Set l = axis.Line
	setAxesLineW(l)
	If  l.foreColor <> m_axisColor Then
		l.foreColor = m_axisColor
	End If


	Dim tickMarks As AutoAxisTickMarks
	Set tickMarks = axis.Tickmarks
	If tickMarks.MajorSide <> grfTicksOff Or tickMarks.MinorSide <> grfTicksOff	Then
		tickMarks.MajorSide = grfTicksOff
		tickMarks.MinorSide = grfTicksOff
	End If

	Dim grid As AutoAxisGrid
	Set grid = axis.Grid
	If 	grid.AtMajorTicks <> True Or grid.AtMinorTicks <> True Then
		grid.AtMajorTicks = True
		grid.AtMinorTicks = True
	End If

	Set l = grid.MajorLine
	If l.foreColor <> m_axisMajorColor Or doubleEq(l.width, m_axisMajorW) = False Then
		l.foreColor = m_axisMajorColor
		l.width = m_axisMajorW
	End If
	setSolidLine(l)

	Set l = grid.MinorLine
	If 	l.foreColor <> m_axisMinorColor Or doubleEq(l.width, m_axisMinorW) = False Then
		l.foreColor = m_axisMinorColor
		l.width = m_axisMinorW
	End If
	setSolidLine(l)
End Sub

Sub setAxes(ByRef axes As AutoAxes)


	Dim xAxis  As AutoAxis
	Set xAxis = axes.Item(1)
	Dim yAxis  As AutoAxis
 	Set yAxis = axes.Item(2)

	setXAxisPosition(xAxis)
	Dim xLink As AutoAxisLink
	Set xLink = xAxis.Link
	Dim yLink As AutoAxisLink
	Set yLink = yAxis.Link


	If _
		xLink.ToAxis <> yAxis.Name Or _
		xLink.xPos <> False Or _
		xLink.yPos <> True Or _
		xLink.length <> False Or _
		xLink.YPosOption <> grfPORightOrBottom  Or _
		yLink.ToAxis <> xAxis.Name Or _
		yLink.xPos <> True Or _
		yLink.yPos <> False Or _
		yLink.length <> False Or _
		yLink.XPosOption <> grfPOLeftOrTop _
	Then
		xLink.length = False
		yLink.length = False
		xLink.ToAxis = ""
		yLink.ToAxis = ""
		xLink.xPos = False
		yLink.xPos = False

		xLink.ToAxis = yAxis.Name
		yLink.ToAxis = xAxis.Name

		xLink.yPos = True
		xLink.YPosOption = grfPORightOrBottom
		yLink.xPos = True
		yLink.XPosOption = grfPOLeftOrTop
	End If

	setMainAxesLine(xAxis)
	setMainAxesLine(yAxis)

	setFont(xAxis.title.Font, m_fontSizeAxisTitle)
	setFont(yAxis.title.Font, m_fontSizeAxisTitle)

	setFont(xAxis.TickLabels.MajorFont, m_fontSizeAxisLabel)
	setFont(yAxis.TickLabels.MajorFont, m_fontSizeAxisLabel)

	Dim j As Integer
	Dim axis  As AutoAxis
	Dim grid As AutoAxisGrid
	For j = 3 To axes.Count
		Set axis = axes.Item(j)
		setAxesLineW(axis.line)
		Set grid = axes.Item(j).Grid
		If 	grid.AtMajorTicks <> False Or grid.AtMinorTicks <> False Then
			grid.AtMajorTicks = False
			grid.AtMinorTicks = False
		End If
		setFontColorChange(axis.title.Font, m_fontSizeAxisTitle, axis.Line.foreColor)
		setFontColorChange(axis.TickLabels.MajorFont, m_fontSizeAxisLabel, axis.Line.foreColor)

		If j <= 4 Then
			If axis.axisType = grfXAxis Then
				Set xLink = axis.Link
				If _
					xLink.ToAxis <> yAxis.Name Or _
					xLink.xPos <> True Or _
					xLink.yPos <> True Or _
					xLink.length <> False Or _
					xLink.YPosOption <> grfPOLeftOrTop Or _
					xLink.XPosOption <> grfPORightOrBottom _
				Then
					xLink.length = False
					xLink.ToAxis = ""
					xLink.xPos = True
					xLink.yPos = True
					xLink.ToAxis = yAxis.Name
					xLink.XPosOption = grfPORightOrBottom
					xLink.YPosOption = grfPOLeftOrTop
				End If

			End If
			If axis.axisType = grfYAxis Then
				Set yLink = axis.Link
				If _
					yLink.ToAxis <> xAxis.Name Or _
					yLink.xPos <> True Or _
					yLink.yPos <> True Or _
					yLink.length <> False Or _
					yLink.YPosOption <> grfPOLeftOrTop Or _
					yLink.XPosOption <> grfPORightOrBottom _
				Then
					yLink.length = False
					yLink.ToAxis = ""
					yLink.xPos = True
					yLink.yPos = True
					yLink.ToAxis = xAxis.Name
					yLink.XPosOption = grfPORightOrBottom
					yLink.YPosOption = grfPOLeftOrTop
				End If
			End If
			If axis.TickLabels.MajorSide <> grfTicksTopRight Or axis.TickLabels.MinorSide <> grfTicksTopRight Then
				axis.TickLabels.MajorSide = grfTicksTopRight
				axis.TickLabels.MinorSide = grfTicksTopRight
			End If
		End If
	Next j
End Sub


Sub Main
	Debug.Clear
	Dim app As Application
	Set app = CreateObject("Grapher.Application")
	Debug.Print("Application found (or created)")
	app.Visible = True

	Dim docs As Documents
	Set docs = app.Documents

	Dim doc As Document
	Set doc = docs.Active
	Debug.Print("Active document set")
	Dim shapes As AutoShapes
	Set shapes = doc.Shapes

	Dim t0, t1 As Double

	Dim i As Integer
	Dim iFirstGraph As Boolean
	iFirstGraph = True
	For i = 1 To shapes.Count
		
		If shapes.Item(i).Type = grfShapeGraph Then
			Dim graph As AutoGraph
			Set graph = shapes.Item(i)

			Debug.Print("Working with graph " +graph.Name + " | " + str$(i))
			Dim axes As AutoAxes
			Set axes = graph.Axes

			If iFirstGraph = True Then
				iFirstGraph = False
				If graph.Axes.Count > 0 And graph.Plots.Count >0 Then
					t0 = Timer
					xAxisPattern.initialize(axes.Item(1), True)
					t1 = Timer
					Debug.Print("First graph init XAxis pattern | " + Str$(t1 - t0) + " s")
					t0 = t1
					plotLimits.initialize(graph.Plots.item(1))
					t1 = Timer
					Debug.Print("First graph init plot limist | " + Str$(t1 - t0) + " s")
				End If
			Else
				If graph.Axes.Count > 0 Then
					t0 = Timer
					xAxisPattern.unify(axes.Item(1))
					t1 = Timer
					Debug.Print("Unify XAxis | " + Str$(t1 - t0) + " s")
				End If
			End If
			t0 = Timer
			setAxes(axes)
			t1 = Timer
			Debug.Print("Set axes representation | " + Str$(t1 - t0) + " s")
			t0 = t1

			setLegend(graph)
			t1 = Timer
			Debug.Print("Set legend | " + Str$(t1 - t0) + " s")
			t0 = t1
			setGraph(graph)
			t1 = Timer
			Debug.Print("Set graph | " + Str$(t1 - t0) + " s")
			t0 = t1
			plotLimits.unify(graph.Plots)
			t1 = Timer
			Debug.Print("Unify plots | " + Str$(t1 - t0) + " s")
		End If
	Next i
End Sub

Sub setLegend(ByRef graph As AutoGraph)
	Dim legends As AutoLegends
	Set legends = graph.Legends
	Dim i As Integer
	For i = 1 To legends.Count
		Dim legend As AutoLegend
		Set legend = legends.Item(i)

		Dim lt As Object 'AutoLegendTitle is missing
		Set lt = legend.title
		If lt.Text <> "" Then
			lt.Text = ""
		End If
		
		Dim f As AutoFill
		Set f = legend.Fill

		If f.BackColor <> grfColorWhite Or f.PatternName <> "None" Then
			f.BackColor = grfColorWhite
			f.PatternName = "None"
		End If
		
		If legend.line.style <> "Invisible" Then
			legend.Line.style = "Invisible"
		End If

		setFont(lt.font, m_fontSizeLegend)

		Dim j As Integer
		For j = 1 To legend.EntryCount
			setFont(legend.EntryFont(j), m_fontSizeLegend)
		Next j

		If legend.DisplayShadow <> False Then
			legend.DisplayShadow = False
		End If
		
	Next i
End Sub

Sub setFont(ByRef font As Object, ByVal size As Integer)
	setFontColorChange(font, size, m_fontColor)
End Sub

Sub setFontColorChange(ByRef font As Object, ByVal size As Integer, fontColor As grfColor)
	If _
		font.face <> m_fontFace Or _
		font.size <> size Or _
		font.Opacity <> m_fontOpacity Or _
		font.Bold <> m_fontBold Or _
		font.Italic <> m_fontItalic Or _
		font.Underline <> m_fontUnderline Or _
		font.StrikeThrough <> m_fontStrikethrough Or _
		font.color <> fontColor _
	Then
		font.face = m_fontFace
		font.size = size
		font.Opacity = m_fontOpacity
		font.Bold = m_fontBold
		font.Italic = m_fontItalic
		font.Underline = m_fontUnderline
		font.StrikeThrough = m_fontStrikethrough
		font.color = fontColor
	End If
End Sub

Sub setGraphTitle (ByRef title As AutoGraphTitle, ByVal titleName As String)
		setFont(title.Font, m_fontSizeGraphTitle)
		If title.text <> titleName Then
			title.text = titleName
		End If
		If title.line.style <> "Invisible" Then
			title.Line.style = "Invisible"
		End If
		If title.Position <> grfCenterTop Then
			title.Position = grfCenterTop
		End If
		If doubleEq(title.xOffset, 0.0) = False Then
			title.xOffset = 0.0
		End If
		If doubleEq(title.yOffset, 0.0) = False Then
			title.yOffset = 0
		End If
		If doubleEq(title.angle, 0.0) = False Then
			title.angle = 0
		End If
End Sub

Sub setGraph(ByRef graph As AutoGraph)
	setGraphTitle(graph.title, graph.Name)

	If graph.BackLine.Style <> "Invisible" Then
		graph.BackLine.Style = "Invisible"
	End If
	If graph.BackFill.PatternName <> "None" Then
		graph.BackFill.PatternName = "None"
	End If
End Sub
