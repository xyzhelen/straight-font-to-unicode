Sub to_StraightBarbara()
'
'Straight Barbara Macro
'
'Last updated: 5-May-2021 by Helen Zhang
'
	Selection.Find.ClearFormatting
	Selection.Find.Replacement.ClearFormatting
	Selection.Find.Replacement.Font.Name = "Times New Roman"
	Selection.Find.Format = True
	Selection.Find.MatchCase = True
	With Selection.Find
		.Text = ChrW(353)
		.Replacement.Text = ChrW(167)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(99) + ChrW(780)
		.Replacement.Text = ChrW(198)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "k" + ChrW(787)
		.Replacement.Text = ChrW(251)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(113) + ChrW(787)
		.Replacement.Text = ChrW(207)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(411) + ChrW(787)
		.Replacement.Text = ChrW(195)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(120) + ChrW(780)
		.Replacement.Text = ChrW(197)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "c" + ChrW(787)
		.Replacement.Text = ChrW(141)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "m" + ChrW(787)
		.Replacement.Text = ChrW(181)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "n" + ChrW(787)
		.Replacement.Text = ChrW(186)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "l" + ChrW(787)
		.Replacement.Text = ChrW(194)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "y" + ChrW(787)
		.Replacement.Text = ChrW(180)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "w" + ChrW(787)
		.Replacement.Text = ChrW(183)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(660)
		.Replacement.Text = ChrW(214)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(322)
		.Replacement.Text = ChrW(168)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(601)
		.Replacement.Text = ChrW(59)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "p" + ChrW(787)
		.Replacement.Text = ChrW(185)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(952)
		.Replacement.Text = ChrW(196)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(183)
		.Replacement.Text = ChrW(165)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(695)
		.Replacement.Text = ChrW(191)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
End Sub
