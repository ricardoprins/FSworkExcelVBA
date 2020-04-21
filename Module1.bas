Attribute VB_Name = "Module1"
Dim objRegex As Object
Dim cell As Range
Dim datesColumn As Range

'-------------------------------------------------------
'
' Spreadsheet adjustment subroutines - FamilySearch
' Date Range Adjustment - Portuguese
'
'-------------------------------------------------------
Sub Dates_PT()

    Set objRegex = CreateObject("VBScript.RegExp") ' This object allows usage of regular expressions (RegEx) to search for text patterns (see below)
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row ' This gets the last row of the sheet
    Set datesColumn = Range(Cells(2, 9), Cells(lastRow, 9)) ' This specifies which is the 'Incl.Dates' column, 8th column from cell H2 to cell H_lastRow
    
    VisualAdjustment
    SymbolRemoval
    DaysRemoval
    FixDateInExcelFormat
    InitialReplacementsES
    SwapMonthYear
    HyperlinkColumn
        
    datesColumn.Replace What:="Mai", Replacement:="Maio" ' Replaces Mai for Maio in all Incl.Dates cells
End Sub
'-------------------------------------------------------
'
' Spreadsheet adjustment subroutines - FamilySearch
' Date Range Adjustment - Spanish
'
'-------------------------------------------------------
 Sub Dates_ES()
     lastRow = Cells(Rows.Count, 1).End(xlUp).Row ' This gets the last row of the sheet
     Set datesColumn = Range(Cells(2, 9), Cells(lastRow, 9)) ' This specifies which is the 'Incl.Dates' column, 8th column from cell H2 to cell H_lastRow
    '
    ' These lines below are references to private subroutines (you can see them below)
    '
    ' It is good practice to always subdivide longer tasks into smaller subroutines;
    '
    ' This also has the benefit to make it easier to reuse code in the future if needed, or to identify mistakes on the code.
    '
    ' A private subroutine is a block of code that can only be accessed by a subroutine within the same module, and that isn't
    ' visible on the Macros menu of Excel.
    '
    ' To create one, all you need to do is to add the keyword Private before the declaration of a subroutine.
    '
    ' As seen below, to invoke a subroutine, you just have to type its name. There is no restriction on how many times
    ' a subroutine can be used: for example, if we needed to run the DaysRemoval 10 times, we could just type in the name
    ' 10 times in a sequence, or use a loop (I can show you how to do loops later on as well).
    ' ---
    ' 1 - we format the columns and add the filter;
    ' 2 - we remove the weird symbols;
    ' 3 - we remove all days from the cells, leaving only months and years;
    ' 4 - we remove dates in the numbered excel format, replacing them for YYYY mmm;
    ' 5 - we remove English months, "aprox" words, capitalized letters, et cetera;
    ' 6 - we swap months for years
    VisualAdjustment
    SymbolRemoval
    DaysRemoval
    FixDateInExcelFormat
    InitialReplacementsES
    SwapMonthYear
    HyperlinkColumn
    
    datesColumn.Replace What:="may", Replacement:="mayo" ' Replaces Mai for Maio in all Incl.Dates cells
 End Sub
Private Sub VisualAdjustment()
    '-------------------------------------------------------
    '
    ' Initial adjustment: Freeze top rows, enable filter,
    '                     autofit colum widths
    '
    ' First, it sets the first row in the ActiveWindow object
    ' (the ActiveWindow object is what Excel is displaying when you run the macro)
    ' to be frozen;
    '
    ' Then, it enables AutoFilter from the cell A1 of the ActiveSheet object
    ' (note the following standard: everything in VBA/Excel is done to an object)
    '
    ' The third step is to have all columns from A to K set to Auto Fit
    ' (the object here is the range A:K of the Columns object, the property
    ' is EntireColumn and the action is AutoFit - another pattern to be observed)
    '-------------------------------------------------------
     With ActiveWindow
         .SplitColumn = 0
         .SplitRow = 1
     End With
     ActiveWindow.FreezePanes = True
     ActiveSheet.Range("A1").AutoFilter
     Columns("A:K").EntireColumn.AutoFit
    '--------------------------------------------------------
End Sub

Private Sub SymbolRemoval()

 Set objRegex = CreateObject("VBScript.RegExp") ' This object allows usage of regular expressions (RegEx) to search for text patterns (see below)
     lastRow = Cells(Rows.Count, 1).End(xlUp).Row ' This gets the last row of the sheet
     Set datesColumn = Range(Cells(2, 9), Cells(lastRow, 9)) ' This specifies which is the 'Incl.Dates' column, 8th column from cell H2 to cell H_lastRow
     
     
     For Each cell In datesColumn
          
     With objRegex
        .Pattern = "[^-\w\s.]+"
        .Global = True ' when we set the Global property to true, this means that we will keep looking for what we want to find even after we've found the value once.
        ' the standard value for this property is False - having it set to False makes it easier to program text changes, but it doesn't help us with the symbols here
        ' I've changed it to true because I want this code to replace all the weird symbols to a single dash.
     End With
     ' whenever we use []s for RegEx, it means that it will look for a single character - any of the ones specified within the []s
     ' in here, for example, we are looking for anything that:
     ' - in the beginning is just a dash;
     ' \w means anything that is between a-z, A-Z, 0-9 or a underscore;
     ' \s means any kind of empty space
     ' . is just a dot
     ' a ^ within the square brackets works as a negation mark
     ' in other words, we are looking for anything that is neither a number, a letter, a dot, a dash or an empty space
     ' we just want the weird stuff
     
     If objRegex.test(cell.Value) Then cell.Value = objRegex.Replace(cell.Value, "-")
     Next cell
     objRegex.Global = False
End Sub

Private Sub DaysRemoval()

Set objRegex = CreateObject("VBScript.RegExp") ' This object allows usage of regular expressions (RegEx) to search for text patterns (see below)
lastRow = Cells(Rows.Count, 1).End(xlUp).Row ' This gets the last row of the sheet
Set datesColumn = Range(Cells(2, 9), Cells(lastRow, 9)) ' This specifies which is the 'Incl.Dates' column, 8th column from cell H2 to cell H_lastRow

With datesColumn ' The With block is used when we want to run a specific action within an object multiple times - this way we don't have to repeat the object many times in the code
    .Replace What:="–", Replacement:="-" 'in every value in the range datesColumn, replace What for Replacement
    .Replace What:="aprox. ", Replacement:=""
End With


For Each cell In datesColumn ' the pattern to be looked for in the datesColumn range is checked using a RegEx object - the property Pattern is what we use to check.
' RegEx (Regular Expressions) are a standardized way to represent text blocks.
    
    objRegex.Pattern = "^\d{1,2}-\d{1,2} "
    ' the ^ right after the quotation mark means "at the beginning of the text"
    ' the \d means any digit 0-9
    '
    ' the numbers between the curly brackets mean how many times the character before them repeat:
    ' in here, this means that it will look for either one or two digits
    ' - is just a -
    ' in other words, it is looking for something either like "1-" or "14-": a day in the beginning of the text, followed by a dash
    If objRegex.test(cell.Value) Then cell.Value = objRegex.Replace(cell.Value, "") ' if it finds it, it will replace it by ""(nothing)

    objRegex.Pattern = "^\d{1,2} |^\d{1,2}-"
    If objRegex.test(cell.Value) Then cell.Value = objRegex.Replace(cell.Value, "")
    
    objRegex.Pattern = "^\d{1,2} "
    If objRegex.test(cell.Value) Then cell.Value = objRegex.Replace(cell.Value, "")
    
    objRegex.Pattern = "-\d{1,2} "
    If objRegex.test(cell.Value) Then cell.Value = objRegex.Replace(cell.Value, "-")
Next cell

End Sub

Private Sub FixDateInExcelFormat()

lastRow = Cells(Rows.Count, 1).End(xlUp).Row ' This gets the last row of the sheet
Set datesColumn = Range(Cells(2, 9), Cells(lastRow, 9)) ' This specifies which is the 'Incl.Dates' column, 8th column from cell H2 to cell H_lastRow


    '--------------------------------------------------------
    '
    ' Fixing data in Date format in cells:
    '
    ' This goes for each one of the cells in the K column looking for anything in the Excel Date format (the numerical 35917 thing you've mentioned before)
    ' It then gets the year and the month and writes it down to the cell in text format.
    '
    '--------------------------------------------------------
     For Each cell In datesColumn ' For each is a type of loop, it works like this: "For each of the individual objects in a range of objects, do the following:"
     Trim (cell.Value) ' Trim is a function that erases additional spaces in a given range (in here, the range is the property Value of the cell range - the name cell could be replaced by anything, it is just a name. I could have named it ball, or dinosaur, it would have the same effect.)
        If IsDate(cell.Value) Then ' The if/else structure: "If (condition) Then (do the following)
            tempText1 = cell.Value ' tempText is just a temporary variable (a placeholder for information) I created to make the code easier to read. It is not necessary to do it like this.
            cell.Value = Format(tempText1, "YYYY") & " " & Format(tempText1, "mmm")
        End If ' Since the instructions after the Then keyword are in multiple lines, you need to specify an End If line, to indicate the end of the conditions.
     Next cell ' The for each loop is enclosed by these two lines: the for each declaration and the Next object line.

End Sub

Private Sub InitialReplacementsES()

     lastRow = Cells(Rows.Count, 1).End(xlUp).Row ' This gets the last row of the sheet
     Set datesColumn = Range(Cells(2, 9), Cells(lastRow, 9)) ' This specifies which is the 'Incl.Dates' column, 8th column from cell H2 to cell H_lastRow
    '--------------------------------------------------------
    '
    ' Replacing dashes and wrong month names. Additional lines
    ' can be added as needed, following the same logic.
    '
    ' Initially, all May values will be set to May to make it
    ' easier to code further fixes. Later on, the values will
    ' be replaced to Mayo.
    '
    '--------------------------------------------------------
     With datesColumn
        .Replace What:=".", Replacement:=""
        .Replace What:="aproxim ", Replacement:=""
        .Replace What:="aprox ", Replacement:=""
        .Replace What:="circa ", Replacement:=""
        .Replace What:="January", Replacement:="ene"
        .Replace What:="February", Replacement:="feb"
        .Replace What:="March", Replacement:="mar"
        .Replace What:="April", Replacement:="abr"
        .Replace What:="May", Replacement:="may", MatchCase:=True
        .Replace What:="June", Replacement:="jun"
        .Replace What:="July", Replacement:="jul"
        .Replace What:="August", Replacement:="ago"
        .Replace What:="September", Replacement:="sep"
        .Replace What:="October", Replacement:="oct"
        .Replace What:="November", Replacement:="nov"
        .Replace What:="December", Replacement:="dic"
        .Replace What:="Enero", Replacement:="ene"
        .Replace What:="Febrero", Replacement:="feb"
        .Replace What:="Marzo", Replacement:="mar"
        .Replace What:="Abril", Replacement:="abr"
        .Replace What:="Mayo", Replacement:="may"
        .Replace What:="Junio", Replacement:="jun"
        .Replace What:="Julio", Replacement:="jul"
        .Replace What:="Agosto", Replacement:="ago"
        .Replace What:="Septiembre", Replacement:="sep"
        .Replace What:="Octubre", Replacement:="oct"
        .Replace What:="Novembro", Replacement:="nov"
        .Replace What:="Diciembre", Replacement:="dic"
        .Replace What:="Apr", Replacement:="abr"
        .Replace What:="Aug", Replacement:="ago"
        .Replace What:="Dec", Replacement:="dic"
        .Replace What:="Ene", Replacement:="ene"
        .Replace What:="Feb", Replacement:="feb"
        .Replace What:="Mar", Replacement:="mar"
        .Replace What:="Abr", Replacement:="abr"
        .Replace What:="May", Replacement:="may"
        .Replace What:="Jun", Replacement:="jun"
        .Replace What:="Jul", Replacement:="jul"
        .Replace What:="Ago", Replacement:="ago"
        .Replace What:="Sep", Replacement:="sep"
        .Replace What:="Oct", Replacement:="oct"
        .Replace What:="Nov", Replacement:="nov"
        .Replace What:="Dic", Replacement:="dic"
        .Replace What:="–", Replacement:="-"
        .Replace What:="/", Replacement:=" "
        
     End With
    '--------------------------------------------------------
End Sub

Private Sub InitialReplacementsPT()
     lastRow = Cells(Rows.Count, 1).End(xlUp).Row ' This gets the last row of the sheet
     Set datesColumn = Range(Cells(2, 9), Cells(lastRow, 9)) ' This specifies which is the 'Incl.Dates' column, 8th column from cell H2 to cell H_lastRow
     With datesColumn
        .Replace What:="aproxim. ", Replacement:=""
        .Replace What:="aprox. ", Replacement:=""
        .Replace What:="circa ", Replacement:=""
        .Replace What:="January", Replacement:="Jan"
        .Replace What:="February", Replacement:="Fev"
        .Replace What:="March", Replacement:="Mar"
        .Replace What:="April", Replacement:="Abr"
        .Replace What:="May", Replacement:="Mai"
        .Replace What:="June", Replacement:="Jun"
        .Replace What:="July", Replacement:="Jul"
        .Replace What:="August", Replacement:="Ago"
        .Replace What:="September", Replacement:="Set"
        .Replace What:="October", Replacement:="Out"
        .Replace What:="November", Replacement:="Nov"
        .Replace What:="December", Replacement:="Dez"
        .Replace What:="Janeiro", Replacement:="Jan"
        .Replace What:="Fevereiro", Replacement:="Fev"
        .Replace What:="Março", Replacement:="Mar"
        .Replace What:="Abril", Replacement:="Abr"
        .Replace What:="Maio", Replacement:="Mai"
        .Replace What:="Junho", Replacement:="Jun"
        .Replace What:="Julho", Replacement:="Jul"
        .Replace What:="Agosto", Replacement:="Ago"
        .Replace What:="Setembro", Replacement:="Set"
        .Replace What:="Outubro", Replacement:="Out"
        .Replace What:="Novembro", Replacement:="Nov"
        .Replace What:="Dezembro", Replacement:="Dez"
        .Replace What:="Feb", Replacement:="Fev"
        .Replace What:="Apr", Replacement:="Abr"
        .Replace What:="May", Replacement:="Maio"
        .Replace What:="Aug", Replacement:="Ago"
        .Replace What:="Sep", Replacement:="Set"
        .Replace What:="Oct", Replacement:="Out"
        .Replace What:="Dec", Replacement:="Dez"
        .Replace What:="jan", Replacement:="Jan", MatchCase:=True
        .Replace What:="fev", Replacement:="Fev", MatchCase:=True
        .Replace What:="mar", Replacement:="Mar", MatchCase:=True
        .Replace What:="abr", Replacement:="Abr", MatchCase:=True
        .Replace What:="mai", Replacement:="Mai", MatchCase:=True
        .Replace What:="jun", Replacement:="Jun", MatchCase:=True
        .Replace What:="jul", Replacement:="Jul", MatchCase:=True
        .Replace What:="ago", Replacement:="Ago", MatchCase:=True
        .Replace What:="set", Replacement:="Set", MatchCase:=True
        .Replace What:="out", Replacement:="Out", MatchCase:=True
        .Replace What:="nov", Replacement:="Nov", MatchCase:=True
        .Replace What:="dez", Replacement:="Dez", MatchCase:=True
        .Replace What:="-", Replacement:="-"
        .Replace What:="/", Replacement:="-"
        .Replace What:=".", Replacement:="-"
        .Replace What:=" ", Replacement:="-"
     End With
End Sub

Private Sub SwapMonthYear()

     Set objRegex = CreateObject("VBScript.RegExp") ' This object allows usage of regular expressions (RegEx) to search for text patterns (see below)
     lastRow = Cells(Rows.Count, 1).End(xlUp).Row ' This gets the last row of the sheet
     Set datesColumn = Range(Cells(2, 9), Cells(lastRow, 9)) ' This specifies which is the 'Incl.Dates' column, 8th column from cell H2 to cell H_lastRow
     
 For Each cell In datesColumn
        
    objRegex.Pattern = "\d{4}" ' Escaping correct value YYYY
     If objRegex.test(cell.Value) And Len(cell.Value) = 4 Then GoTo RightCell
     
     objRegex.Pattern = "\d{4}-\d{4}" ' Escaping correct value YYYY-YYYY
     If objRegex.test(cell.Value) And Len(cell.Value) = 9 Then GoTo RightCell
     
     objRegex.Pattern = "\d{4} [a-zA-Z]{3}" ' Escaping correct value YYYY mmm
     If objRegex.test(cell.Value) And Len(cell.Value) = 8 Then
        cell.Replace What:="-", Replacement:=" "
        GoTo RightCell
     End If
     
     objRegex.Pattern = "\d{4} [a-zA-Z]{3}-[a-zA-Z]{3}" ' Escaping correct value YYYY mmm-mmm
     If objRegex.test(cell.Value) And Len(cell.Value) = 12 Then
        tempText1 = Right(cell.Value, 7)
        cell.Value = Left(cell.Value, 4) & " " & tempText1
        GoTo RightCell
     End If
     
     objRegex.Pattern = "\d{4} [a-zA-Z]{3}-\d{4} a-zA-Z{3}" ' Escaping correct value YYYY mmm-YYYY mmm
     If objRegex.test(cell.Value) And Len(cell.Value) = 17 Then
        tempText1 = Left(cell.Value, 8)
        tempText2 = Right(cell.Value, 8)
        cell.Value = Replace(tempText1, "-", " ") & "-" & Replace(tempText2, "-", " ")
        GoTo RightCell
     End If
        
    flag = InStr(cell.Value, "-") ' InStr is a function to give me the position in which something is found within a String
    ' in this case, I want to know where the dash is, because in the Spanish dates, the dashes are being used only once.
    ' I'll use that to my advantage to invert the months and the years: everything after the dash is a year, and everything before it is a month
    If flag = 0 Then
        cell.Value = Right(cell.Value, 4) & " " & Left(cell.Value, 3) ' if no dash is found, I'll just invert the four characters to the right (Year) and the 3 to the left(month)
        ' I am assigning to the cell Value the combination of strings above, linking them with & (we have to use it to add multiple blocks of text together)
    ElseIf flag = 4 Then ' if the dash is in the fourth position, this means that I have something in the mmm-mmm YYYY standard;
        cell.Value = Right(cell.Value, 4) & " " & Left(cell.Value, 7)
    Else
        tempTextLeft = Left(cell.Value, 8) ' if  the dash is anywhere else, the only standard that is left: mmm YYYY-mmm YYYY
        tempTextLeft = Right(tempTextLeft, 4) & " " & Left(tempTextLeft, 3)
        tempTextRight = Right(cell.Value, 8)
        tempTextRight = Right(tempTextRight, 4) & " " & Left(tempTextRight, 3)
        cell.Value = tempTextLeft & "-" & tempTextRight
    End If
RightCell:  ' :)
 Next cell
End Sub

Private Sub HyperlinkColumn()
     Columns("F:F").EntireColumn.Insert
     lastRow = Cells(Rows.Count, 1).End(xlUp).Row
     For i = 2 To lastRow
        testText1 = Cells(i, 10).Value
        linkText = "https://www.familysearch.org/records/images/search-results?dgsNumbers=" & testText1
        ActiveSheet.Hyperlinks.Add Cells(i, 6), Address:=linkText, TextToDisplay:=Cells(i, 10).Text
     Next i
End Sub
