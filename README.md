Instructions
•	Create a simple Excel workbook and VBA macro in which a user is provided a single button to click. Based on the number they provide in a text box above, a different message box will appear.
o	If the user enters a value of 1, display: “You choose to enter the wooded forest of doom!”
o	If the user enters a value of 2, display: “You choose to enter the fiery volcano of doom!”
o	If the user enters a value of 3, display: “You choose to enter the terrifying jungle of doom!”
o	If the user enters a value of 4, display a similar custom message.
o	If the user enters anything else, display: “Try following directions”
Sub Begin_Journey()
    ' Use conditionals to change message box based on user input
    If (Range("B1").Value = 1) Then
        MsgBox ("You choose to enter the wooded forest of doom!")

    ElseIf (Range("B1").Value = 2) Then
        MsgBox ("You choose to enter the fiery volcano of doom!")

    ElseIf (Range("B1").Value = 3) Then
        MsgBox ("You choose to enter the terrifying jungle of doom!")

   ElseIf (Range("B1").Value = 4) Then
        MsgBox ("You choose to enter the bathroom")

    Else
        MsgBox ("Try following directions")

    End If
End Sub
 
Sub Conditionals():
    ' Simple Conditional Example
    ' ------------------------------------------
    If Range("A2").Value > Range("B2").Value Then
        MsgBox ("Num 1 is greater than Num 2")
    End If
    ' Simple Conditional with If, Else, and Elseif
    ' ------------------------------------------
    If Range("A5").Value > Range("B5").Value Then
        MsgBox ("Num 3 is greater than Num 4")

    ElseIf Range("A5").Value < Range("B5").Value Then
        MsgBox ("Num 4 is greater than Num 3")
    Else
        MsgBox ("Num 3 and Num 4 are equal")
    End If

    ' Conditional with Operators (And)
    ' ------------------------------------------
    If (Range("A8").Value > Range("C8").Value And Range("B8").Value > Range("C8").Value) Then
        MsgBox ("Both Num 5 and Num 6 are greater than Num 7")
    End If
    ' Conditional with Operators (OR)
    ' ------------------------------------------
    If (Range("A8").Value > Range("C8").Value Or Range("B8").Value > Range("C8").Value) Then
        MsgBox ("Either Num 5 and/or Num 6 is greater than Num 7")
    End If

End Sub
 
Sub SentenceBreaker()

    ' Retrieve the user sentence and store in variable
    Dim Sentence As String
    Sentence = Cells(1, 2).Value
    MsgBox (Sentence)

    ' Retrieve the user word numbers and store in variables
    Dim num1 As Integer
    Dim num2 As Integer
    Dim num3 As Integer
    num1 = Cells(4, 1).Value
    num2 = Cells(5, 1).Value
    num3 = Cells(6, 1).Value
    MsgBox (num1)
    MsgBox (num2)
    MsgBox (num3)

    ' Split the user's sentence into words
    Dim SentenceArray() As String
    SentenceArray = Split(Sentence, " ")
    ' Use the word numbers to retrieve the specific words in the sentence
    ' Remember to offset by the 0 index
    Cells(4, 2).Value = SentenceArray(num1 - 1)
    Cells(5, 2).Value = SentenceArray(num2 - 1)
    Cells(6, 2).Value = SentenceArray(num3 - 1)
    

End Sub
 
Sub Variables():

    ' Basic String Variable
    ' ----------------------------------------
    Dim name As String
    name = "Gandalf"

    MsgBox (name)

    ' Basic String Concatenation (Combination)
    ' ----------------------------------------
    Dim title As String
    title = "The Great"

    Dim fullname As String
    fullname = name + " " + title

    MsgBox (fullname)

    ' Basic Integer, Double, Long Variables
    ' ----------------------------------------
    Dim age1 As Integer
    Dim age2 As Integer
    age1 = 5
    age2 = 10

    Dim price As Double
    Dim tax As Double
    price = 19.99
    tax = 0.05

    Dim lightspeed As Long
    lightspeed = 299792458

    ' Basic Numeric manipulation
    ' ----------------------------------------
    MsgBox (age1 + age2)
    Cells(1, 1).Value = price * (1 + tax)

    ' String, Numeric Combination (Casting)
    ' ----------------------------------------
    MsgBox ("I am " + Str(age1) + " years old.")

    ' Booleans
    ' ----------------------------------------
    Dim money_grows_on_trees As Boolean
    money_grows_on_trees = False

End Sub
 
Loops
Sub cereal_loop()

Dim i, total As Integer
Dim sheet1, sheet2 As Worksheet

Set sheet1 = Worksheets("Sheet1")
Set sheet2 = Worksheets("Sheet2")

sheet2.Cells(1, 1).Value = "Brand Names"
sheet2.Cells(1, 3).Value = "Average Calories"

For i = 2 To 66
    sheet2.Range("A" & i).Value = sheet1.Cells(i, 1).Value
    total = total + Cells(i, 3).Value
Next i

sheet2.Cells(2, 3).Value = total / 65
    

End Sub
 
Time to put all your skills to use and create a summary of the US population growth.
Instructions
•	Fill in all of the states in Column A on the Summary sheet.
•	Match each state with its population from 2000 in column B
•	Match each state with its project population for 2030 in column C
•	In Column D, get the percent change from 2000 to 2030.
•	If the population is expected to increase, color with a shade of blue
•	If the population is expected to decrease color, with a shade of yellow.
•	All Code should be written as a single VBA script which can be run on any sheet with the same result.
Hints
•	Before writing any code, plan out how you will accomplish each step. Write a summary of what you think needs to get done, and then write down the steps in plain english. This is called pseudo-coding.

United States Population Growth Part 2
Now that we have some clean looking results, let's see what change there are.
Instructions
•	Find the total population of the USA for 2000 and 2030, respectively.
•	Find the great population and State for 2000 and 2030, respectively.
•	Find the lowest population and state for 2000 and 2030, respectively.
•	Complete this activity using worksheet functions and a single VBA script.
•	Use variables where you can.
Hint
•	Remember for worksheet functions, the values start one higher because they don't include headers when counting.
•	Eyeball the results to make sure they are correct. Keep in mind that this becomes tougher to do when data sets become much larger (i.e. thousands of records).
 
United States Population Growth Part 3
The Grand conclusion! We will now place everything into one giant script!
Instructions
•	Delete the results from everything done in steps one and two.
•	Combine our scripts from parts one and two.
•	Make adjustments where needed and run the script.
•	You results should be exactly the same as running them individually.

Sub population():

Dim total As Long, i As Integer, per_change As Single, pop1, pop2 As Long
Dim sheet1, sheet2, sheet3 As Worksheet

Set sheet1 = Worksheets("Census_2000")
Set sheet2 = Worksheets("Projected_2030")
Set sheet3 = Worksheets("Summary")

total = 0

For i = 2 To 52

    ' states
    sheet3.Cells(i, 1).Value = sheet1.Cells(i, 1).Value

    'populations
    pop1 = sheet1.Cells(i, 2).Value
    pop2 = sheet2.Cells(i, 2).Value

    sheet3.Cells(i, 2).Value = pop1
    sheet3.Cells(i, 3).Value = pop2
    
    per_change = ((pop2 - pop1) / pop1) * 100
    sheet3.Cells(i, 4).Value = per_change

    If per_change > 0 Then
        
        sheet3.Cells(i, 4).Interior.ColorIndex = 33
    Else
        sheet3.Cells(i, 4).Interior.ColorIndex = 36
    
    End If
   
Next i

End Sub

 
Sub population():

Dim total As Long, i As Integer, per_change As Single, pop1, pop2 As Long
Dim sheet1, sheet2, sheet3 As Worksheet

Set sheet1 = Worksheets("Census_2000")
Set sheet2 = Worksheets("Projected_2030")
Set sheet3 = Worksheets("Summary")

total = 0

For i = 2 To 52

    ' states
    sheet3.Cells(i, 1).Value = sheet1.Cells(i, 1).Value

    'populations
    pop1 = sheet1.Cells(i, 2).Value
    pop2 = sheet2.Cells(i, 2).Value

    sheet3.Cells(i, 2).Value = pop1
    sheet3.Cells(i, 3).Value = pop2
    
    per_change = ((pop2 - pop1) / pop1) * 100
    sheet3.Cells(i, 4).Value = per_change

    If per_change > 0 Then
        
        sheet3.Cells(i, 4).Interior.ColorIndex = 33
    Else
        sheet3.Cells(i, 4).Interior.ColorIndex = 36
    
    End If
   
Next i

End Sub
 
Sub population():

Dim l, total As Long, i As Integer
Dim sheet1, sheet2, sheet3 As Worksheet

Set sheet1 = Worksheets("Census_2000")
Set sheet2 = Worksheets("Projected_2030")
Set sheet3 = Worksheets("Summary")

total_2000 = 0
total_2030 = 0

For i = 2 To 52

    ' states 
    sheet3.Cells(i,1).Value = sheet1.Cells(i,1).Value

    'populations
    pop1 = sheet1.Cells(i, 2).Value
    pop2 = sheet2.Cells(i, 2).Value

    sheet3.Cells(i, 2).Value = pop1
    sheet3.Cells(i, 3).Value = pop2
    
    per_change = ((pop2 - pop1) / pop1) * 100
    sheet3.Cells(i, 4).Value = per_change

    If per_change > 0 Then
        
        sheet3.Cells(i, 4).Interior.ColorIndex = 33
    Else
        sheet3.Cells(i, 4).Interior.ColorIndex = 36
    
    End If

    perchange = 0

    ' yearly totals
        total_2000 = total_2000 + sheet1.Cells(i, 2).Value
        total_2030 = total_2030 + sheet2.Cells(i, 2).Value
    
    Next i

' Largest population 2000
max_num1 = WorksheetFunction.Max(sheet3.Range("B2:B52"))
max_state1 = WorksheetFunction.Match(max_num1, sheet3.Range("B2:B52"), 0)
sheet3.Cells(4, 9).Value = sheet3.Cells(max_state1 + 1, 1)
sheet3.Cells(4, 10).Value = max_num1

'Largest population 2030
max_num2 = WorksheetFunction.Max(sheet3.Range("C3:B52"))
max_state2 = WorksheetFunction.Match(max_num2, sheet3.Range("C2:C52"), 0)
sheet3.Cells(10, 9).Value = sheet3.Cells(max_state2 + 1, 1)
sheet3.Cells(10, 10).Value = max_num2

' Low population 2000
low_number1 = WorksheetFunction.Min(Range("B2:B52"))
low_state1 = WorksheetFunction.Match(low_number1, Range("B2:B52"), 0)
sheet3.Cells(5, 9).Value = sheet1.Cells(low_state1 + 1, 1)
sheet3.Cells(5, 10).Value = low_number1

' Low population 20
low_number2 = WorksheetFunction.Min(Range("C2:C52"))
low_state2 = WorksheetFunction.Match(low_number2, Range("C2:C52"), 0)
sheet3.Cells(11, 9).Value = sheet1.Cells(low_state2 + 1, 1)
sheet3.Cells(11, 10).Value = low_number2

' Total population
sheet3.Cells(6, 10).Value = total_2000
sheet3.Cells(12, 10).Value = total_2030

End Sub
# ColumbiaHomework
