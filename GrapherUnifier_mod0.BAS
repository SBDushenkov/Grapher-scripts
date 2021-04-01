Option Explicit

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

Dim m_xAxis As AutoAxis
Dim m_plot As AutoPlot
Dim m_clipping As AutoClipping
Dim m_xTickLabels As AutoAxisTickLabels
Dim m_xTickMarks As AutoAxisTickMarks



Sub setXAxisPosition(axis As AutoAxis)
	If 	axis.xPos <> m_xAxisPos Or axis.length <> m_xAxisLength Then
		axis.xPos = m_xAxisPos
		axis.length = m_xAxisLength
	End If
End Sub

Sub setAxesLineW(ByRef l As AutoLine)
	If  l.width <> m_axisLineW Then
		l.width = m_axisLineW
	End If
End Sub

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
	If l.foreColor <> m_axisMajorColor Or l.width <> m_axisMajorW Then
		l.foreColor = m_axisMajorColor
		l.width = m_axisMajorW
	End If

	Set l = grid.MinorLine
	If 	l.foreColor <> m_axisMinorColor Or l.width <> m_axisMinorW Then
		l.foreColor = m_axisMinorColor
		l.width = m_axisMinorW
	End If


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
		xLink.YPosOption <> grfPORightOrBottom Or _
		xLink.XPosOffset <> 0 Or _
		xLink.YPosOffset <> 0 _
	Then
		xLink.ToAxis = ""
		xLink.xPos = False
		xLink.yPos = True
		xLink.ToAxis = yAxis.Name
		xLink.YPosOption = grfPORightOrBottom
		xLink.XPosOffset = 0
		xLink.YPosOffset = 0
	End If

	If _
		yLink.ToAxis <> xAxis.Name Or _
		yLink.xPos <> True Or _
		yLink.yPos <> False Or _
		yLink.XPosOption <> grfPOLeftOrTop Or _
		yLink.XPosOffset <> 0 Or _
		yLink.YPosOffset <> 0 _
	Then
		yLink.ToAxis = ""
		yLink.xPos = True
		yLink.yPos = False
		yLink.ToAxis = xAxis.Name
		yLink.XPosOption = grfPOLeftOrTop
		yLink.XPosOffset = 0
		yLink.YPosOffset = 0
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
		setFontNoColorChange(axis.title.Font, m_fontSizeAxisTitle)
		setFontNoColorChange(axis.TickLabels.MajorFont, m_fontSizeAxisLabel)
	Next j
End Sub


Sub Main
	Dim app As Application
	Set app = CreateObject("Grapher.Application")
	app.Visible = True

	Dim docs As Documents
	Set docs = app.Documents

	Dim doc As Document
	Set doc = docs.Active

	Dim shapes As AutoShapes
	Set shapes = doc.Shapes

	Dim i As Integer
	Dim iFirstGraph As Integer
	iFirstGraph = -1
	For i = 1 To shapes.Count
		
		If shapes.Item(i).Type = grfShapeGraph Then
			Dim graph As AutoGraph
			Set graph = shapes.Item(i)

			Dim axes As AutoAxes
			Set axes = graph.Axes

			If iFirstGraph < 0 Then
				iFirstGraph = i
				initTicks(axes)
				initPlot(graph)
			Else
				setTicks(axes)
			End If
			setAxes(axes)
			setLegend(graph)
			setGraph(graph)
			setPlots(graph.Plots)
		End If
	Next i
End Sub

Sub initPlot(ByRef graph As AutoGraph)
	If graph.Plots.Count >0 Then
		Set m_plot = graph.Plots(1)
		Set m_clipping = m_plot.Clipping
	End If
End Sub

Sub setPlots(ByRef plots As AutoPlots)
	If Not m_plot Is Nothing Then
			Dim i As Integer
			For i = 1 To plots.Count
				Dim plot As AutoPlot
				Set plot = plots(i)
				If _
					plot.FirstRow <> m_plot.FirstRow Or _
					plot.LastRow <> m_plot.LastRow Or _
					plot.AutoFirstRow <> m_plot.AutoFirstRow Or _
					plot.AutoLastRow <> m_plot.AutoLastRow _
				Then
					plot.FirstRow = m_plot.FirstRow
					plot.LastRow = m_plot.LastRow
					plot.AutoFirstRow = m_plot.AutoFirstRow
					plot.AutoLastRow = m_plot.AutoLastRow
				End If
				Dim clipping As AutoClipping
				Set clipping = plot.Clipping
				clipping.DrawToLimits = m_clipping.DrawToLimits
				If clipping.XMinMode <> clipping.XMinMode Then
					clipping.XMinMode = m_clipping.XMinMode
				End If
				If clipping.XMinMode = grfClipCustom And clipping.xMin <> m_clipping.xMin Then
					clipping.xMin = m_clipping.xMin
				End If
				If clipping.XMaxMode <> clipping.XMaxMode Then
					clipping.XMaxMode = m_clipping.XMaxMode
				End If
				If clipping.XMaxMode = grfClipCustom And clipping.xMax <> m_clipping.xMax Then
					clipping.xMax = m_clipping.xMax
				End If
			Next i
	End If


End Sub


Sub initTicks(ByRef axes As AutoAxes)
	Dim xAxis  As AutoAxis

	Set xAxis = axes.Item(1)
	Set m_xAxis = xAxis
	Set m_xTickMarks = xAxis.Tickmarks
	Set m_xTickLabels = xAxis.TickLabels
End Sub

Sub setTicks(ByRef axes As AutoAxes)
	Dim xAxis  As AutoAxis
	Set xAxis = axes.Item(1)
	Dim yAxis  As AutoAxis
	Set yAxis = axes.Item(2)

	Dim xTickLabels As AutoAxisTickLabels
	Set xTickLabels = xAxis.TickLabels

	If xAxis.AutoMin <> m_xAxis.AutoMin Then
		xAxis.AutoMin = m_xAxis.AutoMin
	End If
	If m_xAxis.AutoMin = False And  xAxis.Min <> m_xAxis.Min Then
		xAxis.Min = m_xAxis.Min
	End If

	If xAxis.AutoMax <> m_xAxis.AutoMax Then
		xAxis.AutoMax = m_xAxis.AutoMax
	End If
	If m_xAxis.AutoMax = False And xAxis.Max <> m_xAxis.Max Then
		xAxis.Max = m_xAxis.Max
	End If


	Dim xlf As AutoLabelFormat
	Set xlf = xTickLabels.MajorFormat

	Dim _xlf As AutoLabelFormat
	Set _xlf = m_xTickLabels.MajorFormat

	If 	xlf.NumericFormat <> _xlf.NumericFormat Or _
		xTickLabels.UseDateTimeFormat <> m_xTickLabels.UseDateTimeFormat Or _
		xlf.DateTimeString <> _xlf.DateTimeString _
	Then
		xlf.NumericFormat = _xlf.NumericFormat
		xTickLabels.UseDateTimeFormat = m_xTickLabels.UseDateTimeFormat
		xlf.DateTimeString = _xlf.DateTimeString
	End If

	Dim xTickMarks As AutoAxisTickMarks
	Set xTickMarks = xAxis.Tickmarks

	If _
		xTickMarks.DTSpacing  <> m_xTickMarks.DTSpacing Or _
		xTickMarks.MajorSpacing <> m_xTickMarks.MajorSpacing Or _
		xTickMarks.EnableDTSpacing <> m_xTickMarks.EnableDTSpacing Or _
		xTickMarks.DTUnits <> m_xTickMarks.DTUnits _
	Then
		xTickMarks.DTSpacing = m_xTickMarks.DTSpacing
		xTickMarks.MajorSpacing = m_xTickMarks.MajorSpacing
		xTickMarks.EnableDTSpacing = m_xTickMarks.EnableDTSpacing
		xTickMarks.DTUnits = m_xTickMarks.DTUnits
	End If

	If _
		xTickMarks.FirstTickValue <> m_xTickMarks.FirstTickValue Or _
		xTickMarks.FirstTickMode <> m_xTickMarks.FirstTickMode Or _
		xTickMarks.StartAt <> m_xTickMarks.StartAt Or _
		xTickMarks.StartAtAuto <> m_xTickMarks.StartAtAuto _
	Then
		xTickMarks.FirstTickValue = m_xTickMarks.FirstTickValue
		xTickMarks.FirstTickMode = m_xTickMarks.FirstTickMode
		xTickMarks.StartAt = m_xTickMarks.StartAt
		xTickMarks.StartAtAuto = m_xTickMarks.StartAtAuto
	End If

	If _
		xTickMarks.DTMinorSpacing <> m_xTickMarks.DTMinorSpacing Or  _
		xTickMarks.DTUnits <> m_xTickMarks.DTUnits Or _
		xTickMarks.EnableDTMinorSpacing <> m_xTickMarks.EnableDTMinorSpacing _
	Then
		xTickMarks.DTMinorSpacing = m_xTickMarks.DTMinorSpacing
		xTickMarks.DTUnits = m_xTickMarks.DTUnits
		xTickMarks.EnableDTMinorSpacing = m_xTickMarks.EnableDTMinorSpacing
	End If
	If 	xTickMarks.MinorDivision <> m_xTickMarks.MinorDivision Or xTickMarks.MinorDivision <> m_xTickMarks.MinorDivision Then
		xTickMarks.MinorDivision = m_xTickMarks.MinorDivision
		xTickMarks.MinorDivision = m_xTickMarks.MinorDivision
	End If

	If xAxis.title.Text <> m_xAxis.title.Text Then
		xAxis.title.Text = m_xAxis.title.Text
	End If
	
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
	If font.color <> m_fontColor Then
		font.color = m_fontColor
	End If
	
	setFontNoColorChange(font, size)
End Sub

Sub setFontNoColorChange(ByRef font As Object, ByVal size As Integer)
	If _
		font.face <> m_fontFace Or _
		font.size <> size Or _
		font.Opacity <> m_fontOpacity Or _
		font.Bold <> m_fontBold Or _
		font.Italic <> m_fontItalic Or _
		font.Underline <> m_fontUnderline Or _
		font.StrikeThrough <> m_fontStrikethrough _
	Then
		font.face = m_fontFace
		font.size = size
		font.Opacity = m_fontOpacity
		font.Bold = m_fontBold
		font.Italic = m_fontItalic
		font.Underline = m_fontUnderline
		font.StrikeThrough = m_fontStrikethrough
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
		If title.xOffset <> 0 Then
			title.xOffset = 0
		End If
		If title.yOffset <> 0 Then
			title.yOffset = 0
		End If
		If title.angle <> 0 Then
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
