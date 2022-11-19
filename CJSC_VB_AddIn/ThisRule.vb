Imports Inventor

Public Class ThisRule
	Inherits AbstractRule

	Dim oInv As Application = Nothing
	Dim oIPT As PartDocument = Nothing
	Dim oPCD As PartComponentDefinition = Nothing

	Public Sub Main()

		oInv = ThisApplication
		oIPT = oInv.ActiveDocument ' ThisDoc.Document
		oPCD = oIPT.ComponentDefinition

		Dim oMaxAreaSB As SurfaceBody = Nothing

		'Dim AreaSortedSBs = From x In oPCD.SurfaceBodies Where x.Visible = True Order By x.MassProperties.Area
		Dim AreaSortedSBs = From x In oPCD.SurfaceBodies Where (x.Visible = True And x.IsSolid = True) Order By x.MassProperties.Area
		oMaxAreaSB = AreaSortedSBs.Last() ' Public member ... not found !!!

	End Sub
End Class