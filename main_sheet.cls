VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CheckBox1_Click()

If CheckBox1.Value = True Then
CADDY.Visible = True
Range("B45") = "Caddy all in one no ring"

ElseIf CheckBox1.Value = False Then
CADDY.Visible = False
Range("B45") = ""

End If

End Sub

Private Sub CheckBox10_Click()

If CheckBox10.Value = True Then
MUDRING3.Visible = True
CheckBox66.Visible = True
Range("R17") = "3/4''"
CheckBox65.Visible = True
Range("T17") = "5/8''"
CheckBox67.Visible = True
Range("V17") = "ADJ"
Range("B46") = "3 d"

ElseIf CheckBox8.Value = False Then
MUDRING3.Visible = False
CheckBox66.Visible = False
Range("R17") = ""
CheckBox65.Visible = False
Range("T17") = ""
CheckBox67.Visible = False
Range("V17") = ""
Range("B46") = ""
End If

End Sub

Private Sub CheckBox11_Click()

If CheckBox11.Value = True Then
MUDRING3.Visible = True
CheckBox69.Visible = True
Range("R18") = "3/4''"
CheckBox68.Visible = True
Range("T18") = "5/8''"
CheckBox70.Visible = True
Range("V18") = "ADJ"
Range("B46") = "4 d"

ElseIf CheckBox8.Value = False Then
MUDRING3.Visible = False
CheckBox69.Visible = False
Range("R18") = ""
CheckBox68.Visible = False
Range("T18") = ""
CheckBox70.Visible = False
Range("V18") = ""
Range("B46") = ""
End If

End Sub

Private Sub CheckBox12_Click()

If CheckBox12.Value = True Then
TSRBS1625.Visible = True
Range("B44") = "Telescoping 16-24"

ElseIf CheckBox12.Value = False Then
TSRBS1625.Visible = False
Range("B44") = ""

End If

End Sub

Private Sub CheckBox13_Click()

If CheckBox13.Value = True Then
STS_2346.Visible = True
Range("B44") = "Snap to stud electrical box"

ElseIf CheckBox13.Value = False Then
STS_2346.Visible = False
Range("B44") = ""

End If

End Sub

Private Sub CheckBox14_Click()

If CheckBox14.Value = True Then
TSGB_16.Visible = True
Range("B44") = "Ceiling  Bracket Fan Rated (35lb.)"

ElseIf CheckBox14.Value = False Then
TSGB_16.Visible = False
Range("B44") = ""

End If

End Sub

Private Sub CheckBox15_Click()

If CheckBox15.Value = True Then
HD512.Visible = True
Range("B44") = "Ceiling  Bracket Fan Rated (35lb.)"

ElseIf CheckBox15.Value = False Then
HD512.Visible = False
Range("B44") = ""

End If

End Sub

Private Sub CheckBox16_Click()

If CheckBox16.Value = True Then
CEILING_BR.Visible = True
Range("B44") = "Ceiling fan box w/brace"

ElseIf CheckBox16.Value = False Then
CEILING_BR.Visible = False
Range("B44") = ""

End If

End Sub

Private Sub CheckBox17_Click()

If CheckBox17.Value = True Then
RACO.Visible = True
Range("B44") = "T-grid box"

ElseIf CheckBox17.Value = False Then
RACO.Visible = False
Range("B44") = ""

End If

End Sub

Private Sub CheckBox18_Click()

Range("AA36").Value = Range("R44").Value
End Sub

Private Sub CheckBox19_Click()

Range("AA38").Value = Range("R46").Value
End Sub

Private Sub CheckBox2_Click()

If CheckBox2.Value = True Then
FOUR_SQ.Visible = True
Range("B45") = "4 sq. sq x 2-1/8"

ElseIf CheckBox2.Value = False Then
FOUR_SQ.Visible = False
Range("B45") = ""

End If

End Sub

Private Sub CheckBox20_Click()

Range("AA40").Value = Range("R48").Value
End Sub

Private Sub CheckBox21_Click()

If CheckBox21.Value = True Then
CheckBox35.Visible = True
CheckBox36.Visible = True
CheckBox37.Visible = True

Range("Q30") = "3/4''"
Range("Q31") = "1''"
Range("Q32") = "1 1/4''"
Range("Q33") = "NA"

ElseIf CheckBox21.Value = False Then

CheckBox35.Visible = False
CheckBox36.Visible = False
CheckBox37.Visible = False
Range("Q30") = ""
Range("Q31") = ""
Range("Q32") = ""
Range("Q33") = ""

End If
End Sub

Private Sub CheckBox22_Click()

If CheckBox22.Value = True Then

CheckBox38.Visible = True
CheckBox39.Visible = True
CheckBox40.Visible = True
CheckBox41.Visible = True
Range("T30") = "38AST"
Range("T31") = "3838AST"
Range("T32") = "40AST"
Range("T33") = "4040AST"

ElseIf CheckBox22.Value = False Then

CheckBox38.Visible = False
CheckBox39.Visible = False
CheckBox40.Visible = False
CheckBox41.Visible = False
Range("T30") = ""
Range("T31") = ""
Range("T32") = ""
Range("T33") = ""

End If
End Sub

Private Sub CheckBox23_Click()
If CheckBox23.Value = True Then
CheckBox42.Visible = True
CheckBox43.Visible = True
CheckBox44.Visible = True

Range("T30") = "3/4''"
Range("T31") = "1''"
Range("T32") = "1 1/4''"
Range("T33") = "NA"

ElseIf CheckBox23.Value = False Then
CheckBox42.Visible = False
CheckBox43.Visible = False
CheckBox44.Visible = False
Range("T30") = ""
Range("T31") = ""
Range("T32") = ""
Range("T33") = ""

End If
End Sub

Private Sub CheckBox24_Click()

Range("B45").Value = Range("T5").Value

End Sub

Private Sub CheckBox25_Click()

Range("B44").Value = Range("T24").Value

End Sub


Private Sub CheckBox3_Click()

If CheckBox3.Value = True Then
MASONRY.Visible = True
Range("B45") = "1 Gang Masonry"

ElseIf CheckBox3.Value = False Then
MASONRY.Visible = False
Range("B45") = ""

End If

End Sub



Private Sub CheckBox4_Click()

If CheckBox4.Value = True Then
GANG3.Visible = True
Range("B45") = "3 gang"

ElseIf CheckBox4.Value = False Then
GANG3.Visible = False
Range("B45") = ""

End If
End Sub



Private Sub CheckBox45_Click()

If CheckBox2.Value = True Then
FOUR_SQ.Visible = True
Range("B45") = "4 in sq x 1-1/2"


ElseIf CheckBox2.Value = False Then
FOUR_SQ.Visible = False
Range("B45") = ""

End If

End Sub

Private Sub CheckBox46_Click()

If CheckBox46.Value = True Then
Range("AA40") = "20A Duplex Standard"

ElseIf CheckBox46.Value = False Then
Range("AA40") = ""

End If

End Sub

Private Sub CheckBox47_Click()

If CheckBox47.Value = True Then
Range("AA35") = "20A Switch single pole"

ElseIf CheckBox47.Value = False Then
Range("AA35") = ""

End If

End Sub

Private Sub CheckBox48_Click()

If CheckBox48.Value = True Then
Range("AA36") = "20A 3 way switch"

ElseIf CheckBox48.Value = False Then
Range("AA36") = ""

End If

End Sub

Private Sub CheckBox49_Click()

If CheckBox49.Value = True Then
Range("AA37") = "20A GFCI"

ElseIf CheckBox49.Value = False Then
Range("AA37") = ""

End If

End Sub

Private Sub CheckBox5_Click()

If CheckBox5.Value = True Then
GANG3.Visible = True
Range("B45") = "4 gang"

ElseIf CheckBox5.Value = False Then
GANG3.Visible = False
Range("B45") = ""

End If

End Sub

Private Sub CheckBox50_Click()

If CheckBox2.Value = True Then
FOUR_SQ.Visible = True
Range("B45") = "4-11 sq. sq x 2-1/8"


ElseIf CheckBox2.Value = False Then
FOUR_SQ.Visible = False
Range("B45") = ""

End If

End Sub

Private Sub CheckBox51_Click()

If CheckBox51.Value = True Then
Range("B45") = "2 Gang Masonry"

ElseIf CheckBox51.Value = False Then
Range("B45") = ""

End If

End Sub

Private Sub CheckBox52_Click()

If CheckBox52.Value = True Then
Range("B45") = "Octagon x 2-1/2 Deep"

ElseIf CheckBox52.Value = False Then
Range("B45") = ""

End If

End Sub

Private Sub CheckBox53_Click()

If CheckBox53.Value = True Then
Range("B45") = "Old Work x 3-1/2 Deep"

ElseIf CheckBox53.Value = False Then
Range("B45") = ""

End If

End Sub

Private Sub CheckBox54_Click()

If CheckBox54.Value = True Then
Range("B44") = "4 Sq. Universal Hubbell"

ElseIf CheckBox54.Value = False Then
Range("B44") = ""

End If

End Sub

Private Sub CheckBox55_Click()

If CheckBox55.Value = True Then
Range("B44") = "Floor Mount Triple Box"

ElseIf CheckBox55.Value = False Then
Range("B44") = ""

End If

End Sub

Private Sub CheckBox56_Click()

If CheckBox56.Value = True Then
Range("B44") = "Floor Mount  Single Box"

ElseIf CheckBox56.Value = False Then
Range("B44") = ""

End If

End Sub

Private Sub CheckBox57_Click()

If CheckBox57.Value = True Then
Range("B44") = "Caddy Open Back Telescoping"

ElseIf CheckBox57.Value = False Then
Range("B44") = ""

End If

End Sub

Private Sub CheckBox58_Click()

If CheckBox58.Value = True Then
Range("AA42") = "20 A Single Receptacle"

ElseIf CheckBox58.Value = False Then
Range("AA42") = ""

End If

End Sub

Private Sub CheckBox59_Click()

If CheckBox59.Value = True Then
Range("D46") = "5/8 in"

ElseIf CheckBox59.Value = False Then
Range("D46") = ""

End If

End Sub

Private Sub CheckBox6_Click()

If CheckBox6.Value = True Then
FIRE.Visible = True
Range("B48") = "4 Sq. FR Insert"

ElseIf CheckBox6.Value = False Then
FIRE.Visible = False
Range("B48") = ""

End If

End Sub

Private Sub CheckBox60_Click()

If CheckBox60.Value = True Then
Range("D46") = "3/4 in"

ElseIf CheckBox60.Value = False Then
Range("D46") = ""

End If

End Sub

Private Sub CheckBox61_Click()

If CheckBox61.Value = True Then
Range("D47") = "5/8 in"

ElseIf CheckBox61.Value = False Then
Range("D47") = ""

End If

End Sub

Private Sub CheckBox62_Click()

If CheckBox62.Value = True Then
Range("D46") = "ADJ"

ElseIf CheckBox62.Value = False Then
Range("D46") = ""

End If

End Sub

Private Sub CheckBox63_Click()

If CheckBox63.Value = True Then
Range("D47") = "3/4 in"

ElseIf CheckBox63.Value = False Then
Range("D47") = ""

End If

End Sub

Private Sub CheckBox64_Click()

If CheckBox64.Value = True Then
Range("D47") = "ADJ"

ElseIf CheckBox64.Value = False Then
Range("D47") = ""

End If

End Sub

Private Sub CheckBox65_Click()

If CheckBox65.Value = True Then
Range("D47") = "5/8 in"

ElseIf CheckBox65.Value = False Then
Range("D47") = ""

End If

End Sub

Private Sub CheckBox66_Click()

If CheckBox66.Value = True Then
Range("D47") = "3/4 in"

ElseIf CheckBox66.Value = False Then
Range("D47") = ""

End If

End Sub

Private Sub CheckBox67_Click()

If CheckBox67.Value = True Then
Range("D47") = "ADJ"

ElseIf CheckBox67.Value = False Then
Range("D47") = ""

End If

End Sub

Private Sub CheckBox68_Click()

If CheckBox68.Value = True Then
Range("D47") = "5/8 in"

ElseIf CheckBox68.Value = False Then
Range("D47") = ""

End If

End Sub

Private Sub CheckBox69_Click()

If CheckBox69.Value = True Then
Range("D47") = "3/4 in"

ElseIf CheckBox69.Value = False Then
Range("D47") = ""

End If

End Sub

Private Sub CheckBox7_Click()

If CheckBox7.Value = True Then
GROUND.Visible = True
Range("B47") = "Ground pigtail"

ElseIf CheckBox7.Value = False Then
GROUND.Visible = False
Range("B47") = ""

End If

End Sub

Private Sub CheckBox70_Click()

If CheckBox70.Value = True Then
Range("D47") = "ADJ"

ElseIf CheckBox70.Value = False Then
Range("D47") = ""

End If

End Sub

Private Sub CheckBox71_Click()

If CheckBox58.Value = True Then
Range("AA43") = "20 A GFCI- Non TR"

ElseIf CheckBox49.Value = False Then
Range("AA43") = ""

End If

End Sub

Private Sub CheckBox72_Click()

If CheckBox72.Value = True Then
Range("AA44") = "20 A Duplex - TR"

ElseIf CheckBox72.Value = False Then
Range("AA44") = ""

End If

End Sub

Private Sub CheckBox73_Click()

If CheckBox73.Value = True Then
Range("AA43") = "20 A GFCI- Non TR"

ElseIf CheckBox73.Value = False Then
Range("AA43") = ""

End If

End Sub

Private Sub CheckBox74_Click()

If CheckBox74.Value = True Then
Range("AA46") = "20 A GFCI -TR"

ElseIf CheckBox74.Value = False Then
Range("AA46") = ""

End If

End Sub

Private Sub CheckBox75_Click()

If CheckBox75.Value = True Then
Range("AA47") = "20 A Duplex w/ USB- Non TR"

ElseIf CheckBox75.Value = False Then
Range("AA47") = ""

End If

End Sub

Private Sub CheckBox76_Click()

If CheckBox76.Value = True Then
Range("AA48") = "20 A Duplex TR"

ElseIf CheckBox76.Value = False Then
Range("AA48") = ""

End If

End Sub

Private Sub CheckBox77_Click()

If CheckBox77.Value = True Then
Range("AA38") = "20A 4 way"

ElseIf CheckBox77.Value = False Then
Range("AA38") = ""

End If

End Sub

Private Sub CheckBox78_Click()

If CheckBox78.Value = True Then
Range("AA39") = "Low Volt Spec"

ElseIf CheckBox78.Value = False Then
Range("AA39") = ""

End If

End Sub

Private Sub CheckBox79_Click()
If CheckBox79.Value = True Then
Range("J48") = "Yes"

ElseIf CheckBox79.Value = False Then

Range("J48") = "No"
End If
End Sub

Private Sub CheckBox8_Click()

If CheckBox8.Value = True Then
MUDRING1.Visible = True
CheckBox60.Visible = True
Range("R15") = "3/4''"
CheckBox59.Visible = True
Range("T15") = "5/8''"
CheckBox62.Visible = True
Range("V15") = "ADJ"
Range("B46") = "1 d"

ElseIf CheckBox8.Value = False Then
MUDRING1.Visible = False
CheckBox60.Visible = False
Range("R15") = ""
CheckBox59.Visible = False
Range("T15") = ""
CheckBox62.Visible = False
Range("V15") = ""
Range("B46") = ""

End If

End Sub

Private Sub CheckBox80_Click()
If CheckBox80.Value = True Then

Range("J51") = "Yes"

ElseIf CheckBox80.Value = False Then

Range("J51") = "No"

End If
End Sub

Private Sub CheckBox83_Click()
If CheckBox83.Value = True Then

Range("J49") = "Yes"

ElseIf CheckBox83.Value = False Then

Range("J49") = "No"

End If
End Sub

Private Sub CheckBox84_Click()
If CheckBox84.Value = True Then

Range("K50") = "Yes"

ElseIf CheckBox84.Value = False Then

Range("K50") = "No"

End If
End Sub

Private Sub CheckBox85_Click()

If CheckBox85.Value = True Then

Range("C47") = "Raised Industrial Cover"

ElseIf CheckBox85.Value = False Then

Range("C47") = ""

End If

End Sub

Private Sub CheckBox9_Click()

If CheckBox9.Value = True Then
MUDRING1.Visible = True
CheckBox63.Visible = True
Range("R16") = "3/4''"
CheckBox61.Visible = True
Range("T16") = "5/8''"
CheckBox64.Visible = True
Range("V16") = "ADJ"
Range("B46") = "2 d"

ElseIf CheckBox8.Value = False Then
MUDRING1.Visible = False
CheckBox63.Visible = False
Range("R16") = ""
CheckBox61.Visible = False
Range("T16") = ""
CheckBox64.Visible = False
Range("V16") = ""
Range("B46") = ""

End If

End Sub

Private Sub ComboBox1_Change()
ComboBox1.List = Array("-", "MC Single 38AST", "MC Duble 3838AST", "MC Single 40AST", "MC Duble 4040AST", "EMT 3/4 in", "EMT 1 in", "EMT 1 1/4in")
Range("B49") = ComboBox1.Text
Range("X33") = ComboBox1.Text
Range("G4") = ComboBox1.Text
End Sub

Private Sub ComboBox10_Change()
ComboBox10.List = Array("-", "Tr", "SL", "PI", "PO")
Range("Y25") = ComboBox10.Text
End Sub

Private Sub ComboBox11_Change()
ComboBox11.List = Array("-", "Tr", "SL", "PI", "PO")
Range("AC25") = ComboBox11.Text
End Sub

Private Sub ComboBox12_Change()
ComboBox12.List = Array("-", "Tr", "SL", "PI", "PO")
Range("AG25") = ComboBox12.Text
End Sub

Private Sub ComboBox13_Change()
ComboBox13.List = Array("-", "12/2", "12/3", "10/2", "EMT", "GRC")
Range("W22") = ComboBox13.Text
End Sub

Private Sub ComboBox14_Change()
ComboBox14.List = Array("-", "1/2 in", "3/4 in", "1 in", "1 1/4 in")
Range("W19") = ComboBox14.Text
End Sub

Private Sub ComboBox15_Change()
ComboBox15.List = Array("-", "12/2", "12/3", "10/2", "EMT", "GRC")
Range("AA22") = ComboBox15.Text
End Sub

Private Sub ComboBox16_Change()
ComboBox16.List = Array("-", "1/2 in", "3/4 in", "1 in", "1 1/4 in")
Range("AA19") = ComboBox16.Text
End Sub

Private Sub ComboBox17_Change()
ComboBox17.List = Array("-", "12/2", "12/3", "10/2", "EMT", "GRC")
Range("AE22") = ComboBox17.Text
End Sub

Private Sub ComboBox18_Change()
ComboBox18.List = Array("-", "1/2 in", "3/4 in", "1 in", "1 1/4 in")
Range("AE19") = ComboBox18.Text
End Sub

Private Sub ComboBox19_Change()
ComboBox19.List = Array("-", "12/2", "12/3", "10/2", "EMT", "GRC")
Range("Y22") = ComboBox19.Text
End Sub

Private Sub ComboBox2_Change()
ComboBox2.List = Array("-", "MC Single 38AST", "MC Duble 3838AST", "MC Single 40AST", "MC Duble 4040AST", "EMT 3/4 in", "EMT 1 in", "EMT 1 1/4in")
Range("B50") = ComboBox2.Text
Range("AB33") = ComboBox2.Text
Range("H5") = ComboBox2.Text
End Sub

Private Sub ComboBox20_Change()
ComboBox20.List = Array("-", "1/2 in", "3/4 in", "1 in", "1 1/4 in")
Range("Y19") = ComboBox20.Text
End Sub

Private Sub ComboBox21_Change()
ComboBox21.List = Array("-", "12/2", "12/3", "10/2", "EMT", "GRC")
Range("AC22") = ComboBox21.Text
End Sub

Private Sub ComboBox22_Change()
ComboBox22.List = Array("-", "1/2 in", "3/4 in", "1 in", "1 1/4 in")
Range("AC19") = ComboBox22.Text
End Sub

Private Sub ComboBox23_Change()
ComboBox23.List = Array("-", "12/2", "12/3", "10/2", "EMT", "GRC")
Range("AG22") = ComboBox23.Text
End Sub

Private Sub ComboBox24_Change()
ComboBox24.List = Array("-", "1/2 in", "3/4 in", "1 in", "1 1/4 in")
Range("AG19") = ComboBox24.Text
End Sub

Private Sub ComboBox3_Change()
ComboBox3.List = Array("-", "MC Single 38AST", "MC Duble 3838AST", "MC Single 40AST", "MC Duble 4040AST", "EMT 3/4 in", "EMT 1 in", "EMT 1 1/4in")
Range("AE33") = ComboBox3.Text
Range("B51") = ComboBox3.Text
Range("I6") = ComboBox3.Text
End Sub

Private Sub ComboBox4_Change()
ComboBox4.List = Array("None", "20A Switch single pole", "20A Switch 3 Way", "20A Swith 4 Way", "Low Volt Spec", "20 A Single Receptacle", "20 A Duplex - TR", "20A GFCI", "20 A GFCI - TR")
Range("AA36") = ComboBox4.Text
End Sub

Private Sub ComboBox5_Change()
ComboBox5.List = Array("None", "20A Switch single pole", "20A Switch 3 Way", "20A Swith 4 Way", "Low Volt Spec", "20 A Single Receptacle", "20 A Duplex - TR", "20A GFCI", "20 A GFCI - TR")
Range("AA38") = ComboBox5.Text
End Sub

Private Sub ComboBox6_Change()
ComboBox6.List = Array("None", "20A Switch single pole", "20A Switch 3 Way", "20A Swith 4 Way", "Low Volt Spec", "20 A Single Receptacle", "20 A Duplex - TR", "20A GFCI", "20 A GFCI - TR")
Range("AA40") = ComboBox6.Text
End Sub

Private Sub ComboBox7_Change()
ComboBox7.List = Array("-", "Tr", "SL", "PI", "PO")
Range("W25") = ComboBox7.Text
End Sub

Private Sub ComboBox8_Change()
ComboBox8.List = Array("-", "Tr", "SL", "PI", "PO")
Range("AA25") = ComboBox8.Text
End Sub

Private Sub ComboBox9_Change()
ComboBox9.List = Array("-", "Tr", "SL", "PI", "PO")
Range("AE25") = ComboBox9.Text
End Sub

Private Sub CommandButton1_Click()



End Sub

Private Sub CommandButton2_Click()

WriteFile

End Sub

Private Sub CommandButton3_Click()
PDFActiveSheet
End Sub

