Attribute VB_Name = "patternAxis"
'#Language "WWB-COM"

Option Explicit

'#Uses "SupportFunctions.bas"

Public Class Unify_AxisMinMax_TicksMarks_TickLabels_TitleText
	Private m_isPrimary As Boolean
	' axis globale
	Private m_axis As AutoAxis
	'Private m_AutoMin As Boolean
	Private m_Min As Double
	'Private m_AutoMax As Boolean
	Private m_Max As Double

	' title
	Private m_TitleText As String

	' labels
	Private m_TickLabels As AutoAxisTickLabels
	Private m_LabelFormat As AutoLabelFormat
	' label format DateTime
	Private m_UseDateTimeFormat As Boolean
	Private m_DateTimeString As String

	' label format NOT DateTime
	Private m_NumericFormat As Long
	Private m_Thousands As Boolean
	Private m_prefix As String
	Private m_postfix As String

	' ticks
	Private m_TickMarks As AutoAxisTickMarks
	Private m_DTMinorSpacing As Double
	Private m_DTMinorUnits As Long
	Private m_DTSpacing As Double
	Private m_DTUnits As Long
	Private m_EnableDTMinorSpacing As Boolean
	Private m_EnableDTSpacing As Boolean
	'Private m_FirstTickMode As Long
	Private m_FirstTickValue As Double
	'Private m_LastTickMode As Long
	Private m_LastTickValue As Double
	Private m_MajorSpacing As Double
	Private m_MajorSpacingAuto As Long
	Private m_MinorDivision As Integer
	Private m_StartAt As Double
	Private m_StartAtAuto As Long
	Private m_LabelDivisor As Double
	Private m_MajorFreq As Integer

	' auto mode (FirstTickMode, LastTickMode, AutoMin, AutoMax )
	' are meaningless because some graphs can have plots with different limits
	' StartAtAuto = true is necessary because value for initial graph is not calculating automaticly
	' basicly even if it is set to auto (for initial XAxis) value is not changing and is wrong

	Public Sub initialize(ByRef axis As AutoAxis, ByVal isPrimary As Boolean)
		Set m_axis = axis
		Set m_TickMarks = axis.Tickmarks
		Set m_TickLabels = axis.TickLabels
		Set m_LabelFormat = m_TickLabels.MajorFormat


		'm_AutoMin = m_axis.AutoMin
		m_Min = m_axis.Min
		'm_AutoMax = m_axis.AutoMax
		m_Max = m_axis.Max
		m_TitleText = m_axis.title.Text

		m_UseDateTimeFormat = m_TickLabels.UseDateTimeFormat
		m_DateTimeString = m_LabelFormat.DateTimeString

		m_LabelDivisor = m_TickLabels.LabelDivisor
		m_MajorFreq = m_TickLabels.MajorFreq

		m_NumericFormat = m_LabelFormat.NumericFormat
		m_Thousands = m_LabelFormat.Thousands
		m_prefix = m_LabelFormat.prefix
		m_postfix = m_LabelFormat.Postfix

		m_DTMinorSpacing = m_TickMarks.DTMinorSpacing
		m_DTMinorUnits = m_TickMarks.DTMinorUnits
		m_DTSpacing = m_TickMarks.DTSpacing
		m_DTUnits = m_TickMarks.DTUnits
		m_EnableDTMinorSpacing = m_TickMarks.EnableDTMinorSpacing
		m_EnableDTSpacing = m_TickMarks.EnableDTSpacing
		'm_FirstTickMode = m_TickMarks.FirstTickMode
		m_FirstTickValue = m_TickMarks.FirstTickValue
		'm_LastTickMode = m_TickMarks.LastTickMode
		m_LastTickValue = m_TickMarks.LastTickValue
		m_MajorSpacing = m_TickMarks.MajorSpacing
		m_MajorSpacingAuto = m_TickMarks.MajorSpacingAuto
		m_MinorDivision = m_TickMarks.MinorDivision
		m_StartAt = m_TickMarks.StartAt
		m_StartAtAuto = m_TickMarks.StartAtAuto

	End Sub

	Public Sub unify(ByRef axis As AutoAxis)
			Dim TickLabels As AutoAxisTickLabels
			Set TickLabels = axis.TickLabels

			Dim axisLabelFormat As AutoLabelFormat
			Set axisLabelFormat = TickLabels.MajorFormat

			If axis.AutoMin <> False Then
				axis.AutoMin = False
			End If
			If axis.AutoMax <> False Then
				axis.AutoMax = False
			End If
			If m_Min > axis.Max Then
				axis.Max = max(m_Max, axis.Max) + m_Max - m_Min
			End If
			If doubleEq(m_Min, axis.Min) = False Then
				axis.Min = m_Min
			End If
			If doubleEq(m_Max, axis.Max) = False Then
				axis.Max = m_Max
			End If

			If TickLabels.UseDateTimeFormat <> m_UseDateTimeFormat Then
				TickLabels.UseDateTimeFormat = m_UseDateTimeFormat
			End If

			If m_UseDateTimeFormat = True And axisLabelFormat.DateTimeString <> m_DateTimeString Then
				axisLabelFormat.DateTimeString = m_DateTimeString
			End If

			If _
				m_UseDateTimeFormat = False And _
				(axisLabelFormat.NumericFormat <> m_NumericFormat Or _
				axisLabelFormat.Thousands <> m_Thousands Or _
				axisLabelFormat.prefix <> m_prefix Or _
				axisLabelFormat.Postfix <> m_postfix) _
			Then
				axisLabelFormat.NumericFormat = m_NumericFormat
				axisLabelFormat.Thousands = m_Thousands
				axisLabelFormat.prefix = m_prefix
				axisLabelFormat.Postfix = m_postfix
			End If

			If TickLabels.LabelDivisor <> m_LabelDivisor Then
				TickLabels.LabelDivisor = m_LabelDivisor
			End If
			If TickLabels.MajorFreq  <> m_MajorFreq Then
				TickLabels.MajorFreq = m_MajorFreq
			End If

			Dim TickMarks As AutoAxisTickMarks
			Set TickMarks = axis.Tickmarks

			If m_EnableDTSpacing = True Then
				If _
					doubleEq(TickMarks.DTSpacing, m_DTSpacing) = False Or _
					TickMarks.EnableDTSpacing <> m_EnableDTSpacing Or _
					TickMarks.DTUnits <> m_DTUnits _
				Then
					TickMarks.DTSpacing = m_DTSpacing
					TickMarks.EnableDTSpacing = m_EnableDTSpacing
					TickMarks.DTUnits = m_DTUnits
				End If
			Else
				If TickMarks.MajorSpacingAuto <>  m_MajorSpacingAuto Then
					TickMarks.MajorSpacingAuto =  m_MajorSpacingAuto
				End If
				If m_MajorSpacingAuto = False And doubleEq(TickMarks.MajorSpacing, m_MajorSpacing) = False Then
					TickMarks.MajorSpacing = m_MajorSpacing
				End If
			End If
			

			If m_EnableDTMinorSpacing = True Then
				If _
					TickMarks.DTMinorSpacing <> m_DTMinorSpacing Or  _
					TickMarks.DTMinorUnits <> m_DTMinorUnits Or _
					TickMarks.EnableDTMinorSpacing <> m_EnableDTMinorSpacing _
				Then
					TickMarks.DTMinorSpacing = m_DTMinorSpacing
					TickMarks.DTMinorUnits = m_DTMinorUnits
					TickMarks.EnableDTMinorSpacing = m_EnableDTMinorSpacing
				End If
			Else
				If TickMarks.MinorDivision <>  m_MinorDivision Then
					TickMarks.MinorDivision =  m_MinorDivision
				End If
			End If

			If TickMarks.FirstTickMode <> grfTickCustom Then
				TickMarks.FirstTickMode = grfTickCustom
			End If

			If TickMarks.LastTickMode <> grfTickCustom Then
				TickMarks.LastTickMode = grfTickCustom
			End If

			If m_FirstTickValue > TickMarks.LastTickValue Then
				TickMarks.LastTickValue = max(m_LastTickValue, TickMarks.LastTickValue) + m_LastTickValue - m_FirstTickValue
			End If
			If doubleEq(TickMarks.FirstTickValue, m_FirstTickValue) = False	Then
				TickMarks.FirstTickValue = m_FirstTickValue
			End If
			If doubleEq(TickMarks.LastTickValue, m_LastTickValue) = False Then
				TickMarks.LastTickValue = m_LastTickValue
			End If


			If m_StartAtAuto = True And TickMarks.StartAtAuto <> m_StartAtAuto Then
				TickMarks.StartAtAuto = m_StartAtAuto
			End If
			If m_StartAtAuto <> True And doubleEq(TickMarks.StartAt, m_StartAt) = False Then
				TickMarks.StartAt  = m_StartAt
			End If

			If axis.title.Text <> m_TitleText Then
				axis.title.Text = m_TitleText
			End If
	End Sub

End Class
