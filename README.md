# testing
test docs




'################################################# Slide 4 - Demographics ###############################################

Sub slide_4()

Dim Demo_1 As String
Dim agebox As Shape

Sheets("Slide4").Select

Demo_1 = Range("Q2").Value ' average age


''''''''''''''''''''''''''''''''''''''''TEXTBOX age
ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 1.5, 200, 20).Select
With Selection
        .name = "agebox"
        .Characters.Text = "Average age " & Demo_1 & " years old"
        .Font.name = "Century Gothic"
        .Font.Size = 12
        .Font.Bold = False
        .Font.Color = RGB(0, 0, 0)
 End With
Set abc = ActiveSheet.Shapes("agebox")
With abc
    .Fill.Transparency = 0
    .Fill.Visible = msoFalse
    .Line.Visible = False
End With

'Sorting region

Range("T2:AD13").Sort key1:=Range("AB2"), _
      order1:=xlDescending, Header:=xlYes




'Naming the boxes


Range("H6:I7").Select
    Selection.Copy
    Range("H10").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "Gender"
    
Range("F6:F7").Select
    Selection.Copy
    Range("F10").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "Income"

Range("CP10:CP12").Select
    Selection.Copy
    Range("CP18").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "relationship"
    
Range("AI2:AK7").Select
    Selection.Copy
    Range("C44").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "region"

Range("CP13").Select
    Selection.Copy
    Range("CP20").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "kids"
    
Range("DH15:DH20").Select
    Selection.Copy
    Range("DH32").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "Sexuality"
    
Range("EB43:EB48").Select
    Selection.Copy
    Range("EB50").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "Ethnicity"

''''''''' pasting it all

'first part: standard code, copy and paste

Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim iCht As Integer

Application.CutCopyMode = False

Set PPApp = GetObject(, "Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

' define spreadheet

Sheets("Slide4").Select

' Paste the shapes individually

ActiveSheet.Shapes("agebox").Copy
With PPPres.Slides(4)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 12.07 * 28.34
        .Top = 4.82 * 28.34
    End With
End With


ActiveSheet.Shapes("Gender").Copy
With PPPres.Slides(4)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 4.07 * 28.34
        .Top = 6.2 * 28.34
    End With
End With


ActiveSheet.Shapes("Income").Copy
With PPPres.Slides(4)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 28.4 * 28.34
        .Top = 15.43 * 28.34
    End With
End With

ActiveSheet.Shapes("Work_Status").Copy

With PPPres.Slides(4)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 22.85 * 28.34
        .Top = 4.72 * 28.34
    End With
End With

ActiveSheet.Shapes("relationship").Copy

With PPPres.Slides(4)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 3.48 * 28.34
        .Top = 10.25 * 28.34
    End With
End With

ActiveSheet.Shapes("kids").Copy

With PPPres.Slides(4)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 3.48 * 28.34
        .Top = 15.41 * 28.34
    End With
End With

ActiveSheet.Shapes("region").Copy

With PPPres.Slides(4)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 11.33 * 28.34
        .Top = 12.71 * 28.34
    End With
End With


ActiveSheet.Shapes("Age Chart").Copy

With PPPres.Slides(4)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 10.82 * 28.34
        .Top = 5.5 * 28.34
    End With
End With

ActiveSheet.Shapes("Social Grade").Copy

With PPPres.Slides(4)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 22.85 * 28.34
        .Top = 10.93 * 28.34
    End With
End With

ActiveSheet.Shapes("Sexuality").Copy

With PPPres.Slides(4)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 37.69 * 28.34
        .Top = 1.49 * 28.34
    End With
End With

ActiveSheet.Shapes("Sexuality_Chart").Copy

With PPPres.Slides(4)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 37.06 * 28.34
        .Top = 7.03 * 28.34
    End With
End With

ActiveSheet.Shapes("Ethnicity").Copy

With PPPres.Slides(4)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 48.33 * 28.34
        .Top = 1.49 * 28.34
    End With
End With

ActiveSheet.Shapes("Ethnicity_Chart").Copy

With PPPres.Slides(4)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 47.29 * 28.34
        .Top = 7.03 * 28.34
    End With
End With

' End of pasting

End Sub


'################################################ Slide 5 - Life events ###############################################

Sub slide_5()





Dim life3, life4, life5, life6, life7 As String
Dim life3box, life4box, life5box, life6box, life7box As Shape
Dim abc As Shape


Sheets("Slide5").Select

'sorting life events by percentage
Range("A2:K20").Sort key1:=Range("I2"), _
      order1:=xlDescending, Header:=xlYes
      
Application.Wait (Now + TimeValue("00:00:02"))

Range("U2:AE20").Sort key1:=Range("AC2"), _
      order1:=xlDescending, Header:=xlYes
      
Application.Wait (Now + TimeValue("00:00:02"))
      
' life1 = life events past 12 months
Range("P2:R7").Select
    Selection.Copy
    Range("P9").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "life1"
    
' life2 = life events next 12 months
Range("AJ2:AL7").Select
    Selection.Copy
    Range("AJ9").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "life2"


life3 = Range("BF3").Value 'alone
life4 = Range("BF4").Value 'with parents
life5 = Range("BF5").Value 'with partner
life6 = Range("BF6").Value 'with other relatives
life7 = Range("BF7").Value 'with children
life8 = Range("BF8").Value 'with other adults

ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 1.5, 80, 20).Select
With Selection
        .name = "life3box"
        .Characters.Text = life3
        .Font.name = "Century Gothic"
        .Font.Size = 14
        .Font.Bold = False
        .Font.Color = RGB(0, 0, 0)
 End With
 Set abc = ActiveSheet.Shapes("life3box")
 With abc
    .Fill.Transparency = 0
    .Fill.Visible = msoFalse
    .Line.Visible = False
 End With

ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 1.5, 80, 20).Select
With Selection
        .name = "life4box"
        .Characters.Text = life4
        .Font.name = "Century Gothic"
        .Font.Size = 14
        .Font.Bold = False
        .Font.Color = RGB(0, 0, 0)
 End With
 Set abc = ActiveSheet.Shapes("life4box")
 With abc
    .Fill.Transparency = 0
    .Fill.Visible = msoFalse
    .Line.Visible = False
 End With
 
 ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 1.5, 80, 20).Select
With Selection
        .name = "life5box"
        .Characters.Text = life5
        .Font.name = "Century Gothic"
        .Font.Size = 14
        .Font.Bold = False
        .Font.Color = RGB(0, 0, 0)
 End With
 Set abc = ActiveSheet.Shapes("life5box")
 With abc
    .Fill.Transparency = 0
    .Fill.Visible = msoFalse
    .Line.Visible = False
 End With
 
 ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 1.5, 80, 20).Select
With Selection
        .name = "life6box"
        .Characters.Text = life6
        .Font.name = "Century Gothic"
        .Font.Size = 14
        .Font.Bold = False
        .Font.Color = RGB(0, 0, 0)
 End With
 Set abc = ActiveSheet.Shapes("life6box")
 With abc
    .Fill.Transparency = 0
    .Fill.Visible = msoFalse
    .Line.Visible = False
 End With
 
 ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 1.5, 80, 20).Select
With Selection
        .name = "life7box"
        .Characters.Text = life7
        .Font.name = "Century Gothic"
        .Font.Size = 14
        .Font.Bold = False
        .Font.Color = RGB(0, 0, 0)
 End With
 Set abc = ActiveSheet.Shapes("life7box")
 With abc
    .Fill.Transparency = 0
    .Fill.Visible = msoFalse
    .Line.Visible = False
 End With
 
 
  ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 1.5, 80, 20).Select
With Selection
        .name = "life8box"
        .Characters.Text = life8
        .Font.name = "Century Gothic"
        .Font.Size = 14
        .Font.Bold = False
        .Font.Color = RGB(0, 0, 0)
 End With
 Set abc = ActiveSheet.Shapes("life8box")
 With abc
    .Fill.Transparency = 0
    .Fill.Visible = msoFalse
    .Line.Visible = False
 End With
''''''''' pasting it all

'first part: standard code, copy and paste

Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim iCht As Integer

Application.CutCopyMode = False

Set PPApp = GetObject(, "Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

' define spreadheet

Sheets("Slide5").Select

' Paste the shapes individually

ActiveSheet.Shapes("life1").Copy
With PPPres.Slides(5)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 4.21 * 28.34
        .Top = 4.07 * 28.34
    End With
End With

ActiveSheet.Shapes("life2").Copy
With PPPres.Slides(5)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 18.91 * 28.34
        .Top = 3.96 * 28.34
    End With
End With


ActiveSheet.Shapes("life3box").Copy
With PPPres.Slides(5)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 2.4 * 28.34
        .Top = 15.95 * 28.34
    End With
End With

ActiveSheet.Shapes("life4box").Copy
With PPPres.Slides(5)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 7.86 * 28.34
        .Top = 15.95 * 28.34
    End With
End With

ActiveSheet.Shapes("life5box").Copy
With PPPres.Slides(5)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 12.89 * 28.34
        .Top = 15.95 * 28.34
    End With
End With

ActiveSheet.Shapes("life6box").Copy
With PPPres.Slides(5)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 18.32 * 28.34
        .Top = 15.95 * 28.34
    End With
End With

ActiveSheet.Shapes("life7box").Copy
With PPPres.Slides(5)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 23.49 * 28.34
        .Top = 15.95 * 28.34
    End With
End With

ActiveSheet.Shapes("life8box").Copy
With PPPres.Slides(5)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 29.06 * 28.34
        .Top = 15.95 * 28.34
    End With
End With



End Sub


'################################################ Slide 6 - Statements  #######################################
Sub slide_6()

Sheets("slide6").Select

'sorting by percentages


Range("A2:m210").Sort key1:=Range("k2"), _
      order1:=xlDescending, Header:=xlYes




End Sub


'################################################ Slide 7 - Passions  #######################################


Sub slide_7()


Sheets("Slide7").Select

'PAIN_1
Range("A2:K22").Sort key1:=Range("I2"), _
      order1:=xlDescending, Header:=xlYes
      
Application.Wait (Now + TimeValue("00:00:05"))
      
'sorting by Index
'PAIN_2
Range("T2:AD22").Sort key1:=Range("AD2"), _
      order1:=xlDescending, Header:=xlYes
      
Application.Wait (Now + TimeValue("00:00:05"))

''''''''' pasting it all

' Naming the boxes

Range("O9:P13").Select
    Selection.Copy
    Range("O21").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "PAS_1"


Range("AH9:AI13").Select
    Selection.Copy
    Range("AH21").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "PAS_2"

'first part: standard code, copy and paste

Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim iCht As Integer

Application.CutCopyMode = False

Set PPApp = GetObject(, "Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

' define spreadheet

Sheets("Slide7").Select

' Paste the shapes individually

ActiveSheet.Shapes("PAS_1").Copy
With PPPres.Slides(7)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 4.08 * 28.34
        .Top = 7.34 * 28.34
    End With
End With

ActiveSheet.Shapes("PAS_2").Copy
With PPPres.Slides(7)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 19.5 * 28.34
        .Top = 7.34 * 28.34
    End With
End With

End Sub

'################################################ Slide 8 - Passions  #######################################
Sub slide_8()



Sheets("Slide8").Select

'sorting by percentages

'PAIN_1
Range("A2:K48").Sort key1:=Range("I2"), _
      order1:=xlDescending, Header:=xlYes

Application.Wait (Now + TimeValue("00:00:05"))

      
'sorting by Index
'PAIN_2
Range("U2:AF48").Sort key1:=Range("Af2"), _
      order1:=xlDescending, Header:=xlYes

Application.Wait (Now + TimeValue("00:00:05"))

' Naming the boxes

Range("B71:K73").Select
    Selection.Copy
    Range("B80").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "PASS_1"


Range("V71:AE73").Select
    Selection.Copy
    Range("V80").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "PASS_2"

''''''''' pasting it all

'first part: standard code, copy and paste

Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim iCht As Integer

Application.CutCopyMode = False

Set PPApp = GetObject(, "Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

' define spreadheet

Sheets("Slide8").Select

' Paste the shapes individually

ActiveSheet.Shapes("PASS_1").Copy
With PPPres.Slides(8)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 3.36 * 28.34
        .Top = 4.76 * 28.34
    End With
End With

ActiveSheet.Shapes("PASS_2").Copy
With PPPres.Slides(8)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 3.36 * 28.34
        .Top = 11.03 * 28.34
    End With
End With
End Sub

'################################################ Slide 9 - BRANDS  #######################################

Sub slide_9()
'sorting by percentages


Sheets("Slide9").Select

'BRAND_1
Range("A2:K109").Sort key1:=Range("I2"), _
      order1:=xlDescending, Header:=xlYes
      
Application.Wait (Now + TimeValue("00:00:05"))

'BRAND_2
Range("V2:AG109").Sort key1:=Range("AG2"), _
      order1:=xlDescending, Header:=xlYes
      
Application.Wait (Now + TimeValue("00:00:05"))

' Naming the boxes

Range("B120:K122").Select
    Selection.Copy
    Range("B135").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "BRAND_1"


Range("W120:AF122").Select
    Selection.Copy
    Range("W135").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "BRAND_2"

''''''''' pasting it all

'first part: standard code, copy and paste

Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim iCht As Integer

Application.CutCopyMode = False

Set PPApp = GetObject(, "Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

' define spreadheet

Sheets("Slide9").Select

' Paste the shapes individually

ActiveSheet.Shapes("BRAND_1").Copy
With PPPres.Slides(9)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 3.36 * 28.34
        .Top = 4.76 * 28.34
    End With
End With

ActiveSheet.Shapes("BRAND_2").Copy
With PPPres.Slides(9)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 3.36 * 28.34
        .Top = 11.03 * 28.34
    End With
End With
End Sub


'################################################ Slide 10 - Sports  #######################################
Sub slide_10()


Sheets("Slide10").Select

'sorting by Combination

'Spor_1
Range("A2:L45").Sort key1:=Range("i2"), _
      order1:=xlDescending, Header:=xlYes
      
Application.Wait (Now + TimeValue("00:00:05"))
      
'Spor_2
Range("Y2:AJ46").Sort key1:=Range("AG2"), _
      order1:=xlDescending, Header:=xlYes
      
Application.Wait (Now + TimeValue("00:00:05"))
      
'Spor_3
Range("AT2:BE38").Sort key1:=Range("BB2"), _
      order1:=xlDescending, Header:=xlYes

Application.Wait (Now + TimeValue("00:00:05"))

' Naming the boxes

Range("C49:F54").Select
    Selection.Copy
    Range("C60").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "Spor_1"


Range("AA49:AD54").Select
    Selection.Copy
    Range("AA60").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "Spor_2"

''''''''' pasting it all

'first part: standard code, copy and paste

Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim iCht As Integer

Application.CutCopyMode = False

Set PPApp = GetObject(, "Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

' define spreadheet

Sheets("Slide10").Select

' Paste the shapes individually

ActiveSheet.Shapes("Spor_1").Copy
With PPPres.Slides(10)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 2.68 * 28.34
        .Top = 3.43 * 28.34
    End With
End With

ActiveSheet.Shapes("Spor_2").Copy
With PPPres.Slides(10)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 18.81 * 28.34
        .Top = 3.43 * 28.34
    End With
End With

ActiveSheet.Shapes("Spor_3").Copy
With PPPres.Slides(10)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 4.63 * 28.34
        .Top = 12.85 * 28.34
    End With
End With
End Sub


'################################################ Slide 11 - Passions #######################################

Sub slide_11()


Sheets("Slide11").Select

'sorting by combination formula


'Sport_1
Range("A2:L45").Sort key1:=Range("i2"), _
      order1:=xlDescending, Header:=xlYes
'Sport_2
Range("T2:AE46").Sort key1:=Range("AB2"), _
      order1:=xlDescending, Header:=xlYes
'Sport_3
Range("AL2:AW46").Sort key1:=Range("AT2"), _
      order1:=xlDescending, Header:=xlYes
'Ent_1
Range("BF2:BQ48").Sort key1:=Range("BN2"), _
      order1:=xlDescending, Header:=xlYes
'Ent_2
Range("BZ2:CK48").Sort key1:=Range("CH2"), _
      order1:=xlDescending, Header:=xlYes


''''''''' pasting it all

'first part: standard code, copy and paste

Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim iCht As Integer

Application.CutCopyMode = False

Set PPApp = GetObject(, "Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

' define spreadheet

Sheets("Slide11").Select

' Paste the shapes individually

ActiveSheet.Shapes("Sport_1").Copy
With PPPres.Slides(11)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 2.62 * 28.34
        .Top = 3.72 * 28.34
    End With
End With

ActiveSheet.Shapes("Sport_2").Copy
With PPPres.Slides(11)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 12.46 * 28.34
        .Top = 3.72 * 28.34
    End With
End With

ActiveSheet.Shapes("Sport_3").Copy
With PPPres.Slides(11)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 22.3 * 28.34
        .Top = 3.73 * 28.34
    End With
End With

ActiveSheet.Shapes("Ent_1").Copy
With PPPres.Slides(11)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 2.62 * 28.34
        .Top = 11.38 * 28.34
    End With
End With

ActiveSheet.Shapes("Ent_2").Copy
With PPPres.Slides(11)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 17.37 * 28.34
        .Top = 11.38 * 28.34
    End With
End With


End Sub


'################################################ Slide 12 - Word of Mouth #######################################

Sub slide_12()


Sheets("Slide12").Select

'sorting by combination formula


'WOM_1
Range("A2:L24").Sort key1:=Range("I2"), _
      order1:=xlDescending, Header:=xlYes
'WOM_2
Range("V2:AG24").Sort key1:=Range("AD2"), _
      order1:=xlDescending, Header:=xlYes

' Naming the boxes

Range("P22:R27").Select
    Selection.Copy
    Range("P30").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "WOM_1"

Range("AK22:AM27").Select
    Selection.Copy
    Range("AK30").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "WOM_2"

''''''''' pasting it all

'first part: standard code, copy and paste

Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim iCht As Integer

Application.CutCopyMode = False

Set PPApp = GetObject(, "Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

' define spreadheet

Sheets("Slide12").Select

' Paste the shapes individually

ActiveSheet.Shapes("WOM_1").Copy
With PPPres.Slides(12)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 3.59 * 28.34
        .Top = 6.19 * 28.34
    End With
End With

ActiveSheet.Shapes("WOM_2").Copy
With PPPres.Slides(12)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 19.11 * 28.34
        .Top = 8.83 * 28.34
    End With
End With

End Sub


'############################### Slide 14 - Touchpoints #######################################'


Sub slide_14()

'sorting
Sheets("Slide14").Select


Range("A2:K60").Sort key1:=Range("I2"), _
    order1:=xlDescending, Header:=xlYes
    
''''''''' pasting it all

'first part: standard code, copy and paste

Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim iCht As Integer

Application.CutCopyMode = False

Set PPApp = GetObject(, "Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

' define spreadheet

Sheets("Slide14").Select

' Paste the shapes individually

ActiveSheet.Shapes("TouchPoints").Copy
With PPPres.Slides(14)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 1.53 * 28.34
        .Top = 3.26 * 28.34
    End With
End With

End Sub

'############################### Slide 15 - Technology adoption #########################

Sub slide_15()


Sheets("Slide15").Select

'sorting by percentages

'Tech_1
Range("A2:N35").Sort key1:=Range("i2"), _
      order1:=xlDescending, Header:=xlYes

'Tech_3
Range("W2:AG5").Sort key1:=Range("AG2"), _
      order1:=xlDescending, Header:=xlYes
      

Range("AL2:AN5").Select
    Selection.Copy
    Range("AL15").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "TECH_3"

''''''''' pasting it all

'first part: standard code, copy and paste

Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim iCht As Integer

Application.CutCopyMode = False

Set PPApp = GetObject(, "Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

' define spreadheet

Sheets("Slide15").Select

' Paste the shapes individually

ActiveSheet.Shapes("Tech_1").Copy
With PPPres.Slides(15)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 1.8 * 28.34
        .Top = 3.94 * 28.34
    End With
End With
      
ActiveSheet.Shapes("Tech_2").Copy
With PPPres.Slides(15)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 22.03 * 28.34
        .Top = 11.58 * 28.34
    End With
End With
      
ActiveSheet.Shapes("Tech_3").Copy
With PPPres.Slides(15)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 22.64 * 28.34
        .Top = 3.89 * 28.34
    End With
End With
      

End Sub

'############################### Slide 16 Media usage #################################

Sub slide_16()


Sheets("Slide16").Select
      
'
'Naming the boxes

Range("P15:Aa20").Select
    Selection.Copy
    Range("P29").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "USAGE_1"
    
Range("P22:Aa23").Select
    Selection.Copy
    Range("P35").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "HOURS_1"
    
''''''''' pasting it all

'first part: standard code, copy and paste

Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim iCht As Integer

Application.CutCopyMode = False

Set PPApp = GetObject(, "Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

' define spreadheet

Sheets("Slide16").Select

' Paste the shapes individually

ActiveSheet.Shapes("HOURS_1").Copy
With PPPres.Slides(16)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 2.21 * 28.34
        .Top = 7.07 * 28.34
    End With
End With


ActiveSheet.Shapes("USAGE_1").Copy
With PPPres.Slides(16)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 2.21 * 28.34
        .Top = 11.67 * 28.34
    End With
End With


End Sub





'############################### Slide 17 Media Motivations #################################

Sub slide_17()


Sheets("Slide17").Select

'Sorting by percentage

Range("A2:K8").Sort key1:=Range("I2"), _
      order1:=xlDescending, Header:=xlYes
      
Application.Wait (Now + TimeValue("00:00:02"))

Range("s2:ac26").Sort key1:=Range("Aa2"), _
      order1:=xlDescending, Header:=xlYes

Application.Wait (Now + TimeValue("00:00:02"))

Range("am2:aw8").Sort key1:=Range("Au2"), _
      order1:=xlDescending, Header:=xlYes
      
Application.Wait (Now + TimeValue("00:00:02"))

Range("bf2:bq17").Sort key1:=Range("bq2"), _
      order1:=xlDescending, Header:=xlYes
      
Application.Wait (Now + TimeValue("00:00:02"))

'Naming the boxes


Range("d15:i16").Select
    Selection.Copy
    Range("D21").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "Mobile_1"

Range("s29:t33").Select
    Selection.Copy
    Range("w29").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "Mobile_2"
    
Range("bh24:bj28").Select
    Selection.Copy
    Range("bn29").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "Mobile_3"


''''''''' pasting it all

'first part: standard code, copy and paste

Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim iCht As Integer

Application.CutCopyMode = False

Set PPApp = GetObject(, "Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

' define spreadheet

Sheets("Slide17").Select

' Paste the shapes individually

ActiveSheet.Shapes("Mobile_1").Copy
With PPPres.Slides(17)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 1.69 * 28.34
        .Top = 4.28 * 28.34
    End With
End With

' moved table to right side of slide - position updated

ActiveSheet.Shapes("Mobile_2").Copy
With PPPres.Slides(17)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 18.8 * 28.34
        .Top = 12 * 28.34
    End With
End With

' moved mobile chart to left of slide - position updated accordingly

ActiveSheet.Shapes("Mobile_Chart").Copy
With PPPres.Slides(17)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 1.69 * 28.34
        .Top = 8.81 * 28.34
    End With
End With

ActiveSheet.Shapes("Mobile_3").Copy
With PPPres.Slides(17)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 17.53 * 28.34
        .Top = 3.84 * 28.34
    End With
End With

End Sub


'################################################ Slide 19 - TV METHODS #######################################

Sub slide_19()


Sheets("Slide19").Select


'sorting by combination formula

'TV_1
Range("A2:AL16").Sort key1:=Range("L2"), _
      order1:=xlDescending, Header:=xlYes

'TV_2
Range("Z2:AK68").Sort key1:=Range("AH2"), _
      order1:=xlDescending, Header:=xlYes

'TV_3
Range("AV2:BG68").Sort key1:=Range("BG2"), _
      order1:=xlDescending, Header:=xlYes


' Naming the boxes

Range("O29:Q33").Select
    Selection.Copy
    Range("O22").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "TV_1"


Range("AO2:AP7").Select
    Selection.Copy
    Range("AO12").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "TV_2"


Range("BK2:BL7").Select
    Selection.Copy
    Range("BK12").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "TV_3"


''''''''' pasting it all

'first part: standard code, copy and paste

Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim iCht As Integer

Application.CutCopyMode = False

Set PPApp = GetObject(, "Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

' define spreadheet

Sheets("Slide19").Select

' Paste the shapes individually

ActiveSheet.Shapes("TV_1").Copy
With PPPres.Slides(19)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 16.92 * 28.34
        .Top = 6.13 * 28.34
    End With
End With

ActiveSheet.Shapes("TV_2").Copy
With PPPres.Slides(19)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 1.75 * 28.34
        .Top = 4.44 * 28.34
    End With
End With

ActiveSheet.Shapes("TV_3").Copy
With PPPres.Slides(19)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 8.96 * 28.34
        .Top = 4.44 * 28.34
    End With
End With


End Sub


'############################## Slide 20 TV programmes & Channels #################################

Sub slide_20()

Sheets("Slide20").Select


'sorting by combination formula

'TV_1a
Range("A2:L75").Sort key1:=Range("I2"), _
      order1:=xlDescending, Header:=xlYes
      
Application.Wait (Now + TimeValue("00:00:05"))

'TV_2a
Range("AB2:AM75").Sort key1:=Range("AM2"), _
      order1:=xlDescending, Header:=xlYes
      
Application.Wait (Now + TimeValue("00:00:05"))

'TV_3a
Range("AW2:BG19").Sort key1:=Range("BE2"), _
      order1:=xlDescending, Header:=xlYes

' Naming the boxes

Range("Q18:U20").Select
    Selection.Copy
    Range("Q25").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "TV_1a"
    
Range("AO18:AS20").Select
    Selection.Copy
    Range("AO25").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "TV_2a"

Range("BL3:BM12").Select
    Selection.Copy
    Range("BL15").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "TV_3a"

''''''''' pasting it all

'first part: standard code, copy and paste

Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim iCht As Integer

Application.CutCopyMode = False

Set PPApp = GetObject(, "Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

' define spreadheet

Sheets("Slide20").Select

' Paste the shapes individually

ActiveSheet.Shapes("TV_1a").Copy
With PPPres.Slides(20)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 1.67 * 28.34
        .Top = 4.71 * 28.34
    End With
End With

ActiveSheet.Shapes("TV_2a").Copy
With PPPres.Slides(20)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 1.67 * 28.34
        .Top = 11.04 * 28.34
    End With
End With

ActiveSheet.Shapes("TV_3a").Copy
With PPPres.Slides(20)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 24.04 * 28.34
        .Top = 3.83 * 28.34
    End With
End With

End Sub
'############################## Slide 21 Diary Data ####################################

Sub Slide_21()

''''''''' pasting it all

'first part: standard code, copy and paste

Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim iCht As Integer

Application.CutCopyMode = False

Set PPApp = GetObject(, "Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

' define spreadheet

Sheets("Slide21").Select

' Paste the shapes individually

ActiveSheet.Shapes("Diary_1").Copy
With PPPres.Slides(21)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 2.25 * 28.34
        .Top = 4.65 * 28.34
    End With
End With

End Sub

'############################### Slide 22 Multi Screen #################################

Sub slide_22()



Sheets("Slide22").Select

' Naming the boxes

Range("P17:S20").Select
    Selection.Copy
    Range("P22").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "Multi_1"


''''''''' pasting it all

'first part: standard code, copy and paste

Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim iCht As Integer

Application.CutCopyMode = False

Set PPApp = GetObject(, "Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

' define spreadheet

Sheets("Slide22").Select

' Paste the shapes individually

ActiveSheet.Shapes("Multi_1").Copy
With PPPres.Slides(22)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 8.33 * 28.34
        .Top = 5.31 * 28.34
    End With
End With

   

End Sub


'############################## Slide 24 Online Activities ###########################

Sub slide_24()

Sheets("Slide24").Select


'Digi_3
Range("A2:N15").Sort key1:=Range("I2"), _
      order1:=xlDescending, Header:=xlYes
      
'sorting by combination formula

'Digi_1
Range("Y2:AL64").Sort key1:=Range("AG2"), _
      order1:=xlDescending, Header:=xlYes
      
'Digi_2
Range("AS2:BF64").Sort key1:=Range("BD2"), _
      order1:=xlDescending, Header:=xlYes
      
      
' Naming the boxes

Range("AO3:AP7").Select
    Selection.Copy
    Range("AO9").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "Digi_1"
    
Range("BH3:BI7").Select
    Selection.Copy
    Range("BH9").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "Digi_2"

''''''''' pasting it all

'first part: standard code, copy and paste

Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim iCht As Integer

Application.CutCopyMode = False

Set PPApp = GetObject(, "Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

' define spreadheet

Sheets("Slide24").Select

' Paste the shapes individually

ActiveSheet.Shapes("Digi_1").Copy
With PPPres.Slides(24)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 0.27 * 28.34
        .Top = 5.19 * 28.34
    End With
End With

ActiveSheet.Shapes("Digi_2").Copy
With PPPres.Slides(24)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 8.42 * 28.34
        .Top = 5.19 * 28.34
    End With
End With

ActiveSheet.Shapes("Digi_3").Copy
With PPPres.Slides(24)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 16.93 * 28.34
        .Top = 4.08 * 28.34
    End With
End With


End Sub

'############################# Slide 25 - Social media ############################

Sub slide_25()

Sheets("Slide25").Select

'social_1
Range("A2:K20").Sort key1:=Range("I2"), _
      order1:=xlDescending, Header:=xlYes

'social_2
Range("W2:AH33").Sort key1:=Range("AE2"), _
      order1:=xlDescending, Header:=xlYes

' Naming the boxes

Range("C29:D38").Select
    Selection.Copy
    Range("C42").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "Social_1"

''''''''' pasting it all

'first part: standard code, copy and paste

Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim iCht As Integer

Application.CutCopyMode = False

Set PPApp = GetObject(, "Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

' define spreadheet

Sheets("Slide25").Select

' Paste the shapes individually

ActiveSheet.Shapes("Social_1").Copy
With PPPres.Slides(25)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 1.74 * 28.34
        .Top = 4.09 * 28.34
    End With
End With

ActiveSheet.Shapes("Social_2").Copy
With PPPres.Slides(25)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 12.18 * 28.34
        .Top = 4.51 * 28.34
    End With
End With

End Sub

'############################### Slide 26 Social Media #################################

Sub slide_26()


Sheets("Slide26").Select


'sorting by combination formula

'Social_1a
Range("V2:AG19").Sort key1:=Range("AG2"), _
      order1:=xlDescending, Header:=xlYes
      
      
'Social_3a
Range("A2:N17").Sort key1:=Range("N2"), _
      order1:=xlDescending, Header:=xlYes
      
' Naming the boxes


Range("AK27:AM31").Select
    Selection.Copy
    Range("AK35").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "Social_1a"

Range("O34:Q38").Select
    Selection.Copy
    Range("O40").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "Social_3a"

    
''''''''' pasting it all

'first part: standard code, copy and paste

Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim iCht As Integer

Application.CutCopyMode = False

Set PPApp = GetObject(, "Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

' define spreadheet

Sheets("Slide26").Select

' Paste the shapes individually

ActiveSheet.Shapes("Social_1a").Copy
With PPPres.Slides(26)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 1.67 * 28.34
        .Top = 3.83 * 28.34
    End With
End With

ActiveSheet.Shapes("Social_2a").Copy
With PPPres.Slides(26)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 1.69 * 28.34
        .Top = 11.92 * 28.34
    End With
End With
    
    
ActiveSheet.Shapes("Social_3a").Copy
With PPPres.Slides(26)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 17.02 * 28.34
        .Top = 7.96 * 28.34
    End With
End With



End Sub

'############################### Slide 28 Radio #################################

Sub slide_28()



Sheets("Slide28").Select


'sorting by combination formula

'Radio_1
Range("A2:L65").Sort key1:=Range("L2"), _
      order1:=xlDescending, Header:=xlYes


'Radio_2
Range("AG2:AT13").Sort key1:=Range("AT2"), _
      order1:=xlDescending, Header:=xlYes
      
      
'Radio4
Range("BC2:BM21").Sort key1:=Range("BK2"), _
      order1:=xlDescending, Header:=xlYes


' Naming the boxes


Range("U8:Y10").Select
    Selection.Copy
    Range("U15").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "Radio1"

Range("AL53:AN56").Select
    Selection.Copy
    Range("AL60").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "Radio2"
    


''''''''' pasting it all

'first part: standard code, copy and paste

Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim iCht As Integer

Application.CutCopyMode = False

Set PPApp = GetObject(, "Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

' define spreadheet

Sheets("Slide28").Select

' Paste the shapes individually

ActiveSheet.Shapes("Radio1").Copy
With PPPres.Slides(28)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 1.89 * 28.34
        .Top = 4.87 * 28.34
    End With
End With

ActiveSheet.Shapes("Radio2").Copy
With PPPres.Slides(28)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 1.67 * 28.34
        .Top = 10.78 * 28.34
    End With
End With
    
    
ActiveSheet.Shapes("Radio4").Copy
With PPPres.Slides(28)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 17.13 * 28.34
        .Top = 4.24 * 28.34
    End With
End With

End Sub


'############################### Slide 29 Press #################################

Sub slide_29()


Sheets("Slide29").Select


'Sorting attitudes by formula

Range("A2:N16").Sort key1:=Range("N2"), _
      order1:=xlDescending, Header:=xlYes
      
Application.Wait (Now + TimeValue("00:00:05"))

Range("BD2:BQ16").Sort key1:=Range("BQ2"), _
      order1:=xlDescending, Header:=xlYes
      
Application.Wait (Now + TimeValue("00:00:05"))
      
'Sorting Sections by percentage

Range("V2:AH21").Sort key1:=Range("AD2"), _
      order1:=xlDescending, Header:=xlYes

Range("BY2:CK21").Sort key1:=Range("CG2"), _
      order1:=xlDescending, Header:=xlYes

'Sorting newspaper titles by percentage

Range("DH2:DT29").Sort key1:=Range("DS2"), _
      order1:=xlDescending, Header:=xlYes
      
'Sorting magazine titles by percentage

Range("EC2:EO92").Sort key1:=Range("EN2"), _
      order1:=xlDescending, Header:=xlYes

' Naming the boxes

Range("C21:D23").Select
    Selection.Copy
    Range("D27").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "NEWS_1"

Range("Y42:AC43").Select
    Selection.Copy
    Range("Y47").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "NEWS_2"

Range("DS36:DU37").Select
    Selection.Copy
    Range("DH40").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "NEWS_3"
    
Range("AT15:AW15").Select
    Selection.Copy
    Range("AT20").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "NEWS_4"
    
    
Range("BG21:BH23").Select
    Selection.Copy
    Range("BF28").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "MAGS_1"
    
    
    
Range("CD42:CH43").Select
    Selection.Copy
    Range("CD48").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "MAGS_2"
    
    
Range("EQ36:ES37").Select
    Selection.Copy
    Range("EQ40").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "MAGS_3"

Range("CW15:CZ15").Select
    Selection.Copy
    Range("CW20").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "MAGS_4"

      

''''''''' pasting it all

'first part: standard code, copy and paste

Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim iCht As Integer

Application.CutCopyMode = False

Set PPApp = GetObject(, "Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

' define spreadheet

Sheets("Slide29").Select

' Paste the shapes individually

ActiveSheet.Shapes("NEWS_1").Copy
With PPPres.Slides(29)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 2.89 * 28.34
        .Top = 4.28 * 28.34
    End With
End With

ActiveSheet.Shapes("NEWS_2").Copy
With PPPres.Slides(29)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 2.89 * 28.34
        .Top = 8.36 * 28.34
    End With
End With

ActiveSheet.Shapes("NEWS_3").Copy
With PPPres.Slides(29)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 3.63 * 28.34
        .Top = 11.78 * 28.34
    End With
End With

ActiveSheet.Shapes("NEWS_4").Copy
With PPPres.Slides(29)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 2.94 * 28.34
        .Top = 16.75 * 28.34
    End With
End With

ActiveSheet.Shapes("MAGS_1").Copy
With PPPres.Slides(29)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 18.18 * 28.34
        .Top = 4.28 * 28.34
    End With
End With

ActiveSheet.Shapes("MAGS_2").Copy
With PPPres.Slides(29)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 18.13 * 28.34
        .Top = 8.36 * 28.34
    End With
End With

ActiveSheet.Shapes("MAGS_3").Copy
With PPPres.Slides(29)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 18.57 * 28.34
        .Top = 11.78 * 28.34
    End With
End With

ActiveSheet.Shapes("MAGS_4").Copy
With PPPres.Slides(29)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 18.18 * 28.34
        .Top = 16.75 * 28.34
    End With
End With

End Sub


'############################### Slide 30 Cinema #################################

Sub slide_30()



Sheets("Slide30").Select

'Sorting by percentage
'CINE_2
Range("V2:AF29").Sort key1:=Range("AD2"), _
      order1:=xlDescending, Header:=xlYes
      
'sorting by combination formula

'CINE_4
Range("AS2:BD9").Sort key1:=Range("BD2"), _
      order1:=xlDescending, Header:=xlYes
      
' Naming the boxes

Range("F21:F25").Select
    Selection.Copy
    Range("F33").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "CINE_1"


Range("BH26:BJ30").Select
    Selection.Copy
    Range("BY33").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "CINE_4"



''''''''' pasting it all

'first part: standard code, copy and paste

Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim iCht As Integer

Application.CutCopyMode = False

Set PPApp = GetObject(, "Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

' define spreadheet

Sheets("Slide30").Select

' Paste the shapes individually

ActiveSheet.Shapes("CINE_1").Copy
With PPPres.Slides(30)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 6.25 * 28.34
        .Top = 4.21 * 28.34
    End With
End With

ActiveSheet.Shapes("CINE_2").Copy
With PPPres.Slides(30)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 17.01 * 28.34
        .Top = 3.85 * 28.34
    End With
End With

' Paste the shapes individually

ActiveSheet.Shapes("CINE_4").Copy
With PPPres.Slides(30)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 1.66 * 28.34
        .Top = 9.82 * 28.34
    End With
End With
End Sub

'############################### Slide 31 OOH #################################

Sub Slide_31()


Sheets("Slide31").Select

'Sorting by percentage
'OHH_3
Range("W2:AG12").Sort key1:=Range("AE2"), _
      order1:=xlDescending, Header:=xlYes
      

      
' Naming the boxes

Range("E20:I21").Select
    Selection.Copy
    Range("E25").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "OOH_1"
    

'first part: standard code, copy and paste

Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim iCht As Integer

Application.CutCopyMode = False

Set PPApp = GetObject(, "Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

' define spreadheet

Sheets("Slide31").Select

' Paste the shapes individually

ActiveSheet.Shapes("OOH_1").Copy
With PPPres.Slides(31)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 9.59 * 28.34
        .Top = 7.17 * 28.34
    End With
End With


ActiveSheet.Shapes("OOH_3").Copy
With PPPres.Slides(31)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 1.79 * 28.34
        .Top = 11.71 * 28.34
    End With
End With

End Sub


'############################### Slide 32 Gaming #################################

Sub Slide_32()

Sheets("Slide32").Select

'Sorting by percentage
'GAMING4
Range("AV2:BF11").Sort key1:=Range("BD2"), _
      order1:=xlDescending, Header:=xlYes

'GAMING3
Range("BU2:CE13").Sort key1:=Range("CE2"), _
      order1:=xlDescending, Header:=xlYes

' Naming the boxes

Range("G10:J11").Select
    Selection.Copy
    Range("G15").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "GAMING1"

Range("AD10:AG11").Select
    Selection.Copy
    Range("AD15").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "GAMING2"
    
Range("BW15:BX20").Select
    Selection.Copy
    Range("bw25").Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.name = "GAMING3"
    
Application.Wait (Now + TimeValue("00:00:02"))

'first part: standard code, copy and paste

Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim iCht As Integer

Application.CutCopyMode = False

Set PPApp = GetObject(, "Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

' define spreadheet

Sheets("Slide32").Select

' Paste the shapes individually

ActiveSheet.Shapes("GAMING1").Copy
With PPPres.Slides(32)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 1.69 * 28.34
        .Top = 6.75 * 28.34
    End With
End With

ActiveSheet.Shapes("GAMING2").Copy
With PPPres.Slides(32)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 1.51 * 28.34
        .Top = 13.07 * 28.34
    End With
End With

ActiveSheet.Shapes("GAMING3").Copy
With PPPres.Slides(32)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 11.95 * 28.34
        .Top = 3.82 * 28.34
    End With
End With

ActiveSheet.Shapes("GAMING4").Copy
With PPPres.Slides(32)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 12.54 * 28.34
        .Top = 10.68 * 28.34
    End With
End With

End Sub

'############################### Slide 34 Sponsorship #################################

Sub slide_34()

Sheets("Slide34").Select

'Sorting by percentage
'Sponsor_1
Range("A2:K9").Sort key1:=Range("I2"), _
      order1:=xlDescending, Header:=xlYes


'first part: standard code, copy and paste

Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim iCht As Integer

Application.CutCopyMode = False

Set PPApp = GetObject(, "Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

' define spreadheet

Sheets("Slide34").Select

' Paste the shapes individually

ActiveSheet.Shapes("Sponsor_1").Copy
With PPPres.Slides(34)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 2.19 * 28.34
        .Top = 3.95 * 28.34
    End With
End With
End Sub



'############################### Slide 35 issues #################################

Sub slide_35()

Sheets("Slide35").Select

'Sorting by percentage

Range("a2:K11").Sort key1:=Range("i2"), _
      order1:=xlDescending, Header:=xlYes


'first part: standard code, copy and paste

Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim iCht As Integer

Application.CutCopyMode = False

Set PPApp = GetObject(, "Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

' define spreadheet

Sheets("Slide35").Select

' Paste the shapes individually

ActiveSheet.Shapes("Chart 2").Copy
With PPPres.Slides(35)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 1.99 * 28.34
        .Top = 2.64 * 28.34
    End With
End With
End Sub

'################################################ Slide 36 - Schwartz Values  #######################################


Sub slide_36()


Sheets("Slide36").Select


''''''''' pasting it all

'first part: standard code, copy and paste

Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim iCht As Integer

Application.CutCopyMode = False

Set PPApp = GetObject(, "Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

' define spreadheet

Sheets("Slide36").Select

' Paste the shapes individually

ActiveSheet.Shapes("Schwartz_1").Copy
With PPPres.Slides(36)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 2.32 * 28.34
        .Top = 3.33 * 28.34
    End With
End With

End Sub

'############################### Slide 37 Media Motivations #################################

Sub slide_37()


Sheets("Slide37").Select

''''''''' pasting it all

'first part: standard code, copy and paste

Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim iCht As Integer

Application.CutCopyMode = False

Set PPApp = GetObject(, "Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

' define spreadheet

Sheets("Slide37").Select

' Paste the shapes individually

ActiveSheet.Shapes("Media_States").Copy
With PPPres.Slides(37)
    .Shapes.Paste
    With .Shapes(.Shapes.Count)
        .Left = 2.32 * 28.34
        .Top = 3.85 * 28.34
    End With
End With

End Sub


