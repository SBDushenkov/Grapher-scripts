Attribute VB_Name = "Unify_PlotDataLimits"
'#Language "WWB-COM"

Option Explicit

'#Uses "SupportFunctions.bas"

'**********************************************************************
' Class getting one plot as pattern end copy settings to specified plots
' Works with limits and clipping
'**********************************************************************
Public Class Unify_PlotDataLimits
	Private m_Worksheet As String
	Private m_plot As AutoPlot
	Private m_clipping As AutoClipping
	Private m_AutoFirstRow As Boolean
	Private m_AutoLastRow As Boolean
	Private m_FirstRow As Long
	Private m_LastRow As Long

	Private m_DrawToLimits As Long
	Private m_xMin As Double
	Private m_XMinMode As Long
	Private m_xMax As Double
	Private m_XMaxMode As Long

	Public Sub initialize(ByRef plot As AutoPlot)
		Set m_plot = plot
		Set m_clipping = m_plot.Clipping
		m_Worksheet = plot.worksheet
		m_AutoFirstRow = m_plot.AutoFirstRow
		m_AutoLastRow = m_plot.AutoLastRow
		m_FirstRow = m_plot.FirstRow
		m_LastRow = m_plot.LastRow

		m_DrawToLimits = m_clipping.DrawToLimits
		m_xMin = m_clipping.xMin
		m_XMinMode = m_clipping.XMinMode
		m_xMax = m_clipping.xMax
		m_XMaxMode = m_clipping.XMaxMode
	End Sub


	Public Sub unify(ByRef plots As AutoPlots)
		If Not m_plot Is Nothing Then
			Dim plot As AutoPlot
			Dim clipping As AutoClipping
			Dim i As Integer
			For i = 1 To plots.Count
				Set plot = plots(i)
				If plot.worksheet = m_Worksheet Then
					If m_AutoLastRow > plot.LastRow Then
						plot.LastRow = max(m_AutoLastRow, plot.LastRow) + m_AutoLastRow - m_FirstRow
					End If

					If plot.AutoFirstRow <> m_AutoFirstRow Then
						plot.AutoFirstRow = m_AutoFirstRow
					End If
					If m_AutoFirstRow = False And plot.FirstRow <> m_FirstRow Then
						plot.FirstRow = m_FirstRow
					End If

					If plot.AutoLastRow <> m_AutoLastRow Then
						plot.AutoLastRow = m_AutoLastRow
					End If
					If m_AutoLastRow = False And plot.LastRow <> m_LastRow Then
						plot.LastRow = m_LastRow
					End If

					Set clipping = plot.Clipping

					If m_xMin > clipping.xMax Then
						clipping.xMax = max(m_xMax, clipping.xMax) + m_xMax - m_xMin
					End If
					If clipping.XMinMode <> m_XMinMode Then
						clipping.XMinMode = m_XMinMode
					End If
					If m_XMinMode = grfClipCustom And doubleEq(clipping.xMin, m_xMin) = False Then
						clipping.xMin = m_xMin
					End If

					If clipping.XMaxMode <> m_XMaxMode Then
						clipping.XMaxMode = m_XMaxMode
					End If
					If m_XMaxMode = grfClipCustom And doubleEq(clipping.xMax, m_xMax) = False Then
						clipping.xMax = m_xMax
					End If

					If clipping.DrawToLimits <> m_DrawToLimits Then
						clipping.DrawToLimits = m_DrawToLimits
					End If
				End If
			Next i

		Else
			Debug.Print "'AutoPlots' object is nothing"
	
		End If

	End Sub

End Class
