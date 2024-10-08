# Interest rate converter

This is a Visual Basic  6 program that convert temperature, it acccepts user inputs in Fahrenheit and provide a corresponding output in Celsius.


![live-preview](https://github.com/marveeygoodlife/Interest-rate-converter/blob/main/images/Screenshot%202024-10-08%20173641.png)

 ## Table of Contents

1. [AUTHOR](#marvellous-ediagbonya)
1. [PROBLEM STATEMENT](#problem-statement)
2. [PROBLEM SPECIFICATION](#problem-specification)
3. [ANALYSIS](#analysis)
4. [DESIGN](#design)
5. [IMPLEMENTATION](#implementation)
5. [RESULT](#result)

## Marvellous Ediagbonya

Noun,  
Abuja  
07015824775  
CIT 104 PRACTICAL  
7th October 2024  

Exercise 1:   
 Given an initial amount and a yearly interest rate calculate the amount of money expected at the end of a period of not less than 5 years. Assume n = 1.   
Hint: 𝐴 = 𝑃 (1 + 𝑟 𝑛 ) 𝑛𝑡   
A = Amount of money expected   
P = Initial amount   
R = Interest Rate in percent   
r = R/100   
t = Time Involved in years, 0.5 years is calculated as 6 months, etc.   

## PROBLEM STATEMENT

The goal of this practical Project is to develop a simple application that calculates the future amount of an investment by using compound Interest . Users will enter the Amount(Principal), interest rate and time in years.The application will validate the Inputs to make sure the year is not less than 5 years and not greater than 100 years. 

If the input is valid the application will calculate with the exercise 1 formula:

<strong> A = P (1+n / r​)nt </strong>

## PROBLEM SPECIFICATION

<strong>Inputs: </strong>
- Principal Amount (Double)  
- Interest Rate (Double)  
- Time (Double, it must be between 5 and 100 years)
  
<strong>Outputs:</strong>

- Final amount displayed in currency form.  
- A prompt response for users  to try another Value.
  
<strong>Validation Rules:</strong>   


- Every input entered must be in numeric format.  
- Time must be 5 years at least and not more than 100 years.  

<strong>Formula:  </strong>
  
<strong> Hint: 𝐴 = 𝑃 (1 + 𝑟 / 𝑛 ) 𝑛t   </strong>
Where:  
P is the principal/ initial amount to invest  
A is the total amount expected  
r is the interest rate divided by 100 (i.e., r = Rate / 100).  
n = 1 since the interest is compounded annually.  
t is the time in years.  


## ANALYSIS

The application must provide user-friendly interactiveness while performing its basic function of calculating the Interest or Loan. 
It should smoothly handle invalid user inputs through Validation check and provide informative messages to the user.

The flow of data involves  
- Accepting user inputs 
- Validating user inputs
- Calculating using the formula provided on the exercise
- providing an accurate response in currency.

## DESIGN

<strong>User Interface:</strong>  

The application will be designed using the following formats  
User Interface:   
- Labels for Principal, Rate and Time and the Output.  
- Textboxes to accept user inputs  
- A command buttons to calculate  
- A command  button to clear input fields and Label as well  

<strong>Logic:</strong>

- Accepts user inputs  
- Validation inputs to make sure it's the only number entered as inputs.  
- Calculation logic (It uses the formula to calculate ).  
- Result is displayed in a label.  

<b>Flowchart: </b>

![FlowChart](https://github.com/marveeygoodlife/Interest-rate-converter/blob/main/images/Copy%20of%20Project%20proposal.jpg)

  <b> A simple flowchart Algorithm</b>


- User Interface:
  
The user input values into the text boxes values for Principal, Rate and Time.
 
- Code Execution:
  
When Users click the Calculate button, the inputs are validated and if it's correct the formula is applied using the provided formula.    
- The output is calculated and displayed in the label.  

## IMPLEMENTATION

<strong>CODE STARTS HERE! </strong>
```vb6
Private Sub Form_Load()

    ' Set default values for Textboxes
    txtPrincipal.Text = "0"     ' Default principal amount to 0
   txtRate.Text = "0"          ' Default interest rate to 0
    txtTime.Text = "1"          ' Default time to 1 year

    ' Set a default message in the Output Label (optional)
    lblOutput.Caption = "Please enter values and click Calculate."
End Sub ```

Private Sub cmdCalculate_Click()
    ' Declare variables for the inputs
    Dim principal As Double
    Dim rate As Double
    Dim time As Double
    Dim amount As Double
    Dim r As Double ' This will be rate/100 for calculation

    ' Input validation: To check if the inputs are numeric


    If IsNumeric(txtPrincipal.Text) And IsNumeric(txtRate.Text) And  
' Continuation of Code 🔻
       IsNumeric(txtTime.Text) Then
  ' Get user inputs
        principal = CDbl(txtPrincipal.Text) ' Convert input text to a number
        rate = CDbl(txtRate.Text) ' Convert rate to number
        time = CDbl(txtTime.Text) ' Convert time to number

        ' Check if time period is at least 5 years
        If time < 5 Then
            MsgBox "Please enter a time period of at least 5 years."
            Exit Sub
            ElseIf time > 100 Then
            
MsgBox "Please enter a time period of no more than 100 years."
Exit Sub
 End If
        ' Calculate the interest rate in decimal
        r = rate / 100

        ' Apply the formula A = P(1 + r)^t, with n = 1 (compounded annually)
        amount = principal * (1 + r) ^ time

        ' Display the result in the Output Label
        lblOutput.Caption = "The final amount is: " & Format(amount, "Currency")
        MsgBox "Try Another Value " & Me.Caption ' Show message with form caption
    Else
        ' Display error message if inputs are invalid
        MsgBox "Please enter valid numbers for Principal, Rate, and Time."
    End If
End Sub
Private Sub cmdClear_Click()

    ' Clear all the input textboxes
    txtPrincipal.Text = ""
    txtRate.Text = ""
    txtTime.Text = ""
    
    ' Optionally clear the output label
    lblOutput.Caption = ""
    
    ' Return focus to the first input field
    txtPrincipal.SetFocus
End Sub
```
CODE ENDS HERE.


## RESULT

When valid inputs  are provided and the application is run, the user receives a message that displays the future amount of their investment based on the Principal, the Rate as well as the Time.   

<b>For Instance: </b>  

If a user inputs a principal of $1000, at a rate of 5%, and a time period of 9 years, the application will display approximately $1, 551.33.

 
<strong>Error Handling:</strong>

If the user enters a wrong input, or a Time less than 5  years or greater than 100 years, the right error message will display to guide them to  input appropriate Numeric format.
This application serves as an effective tool to calculate future interest,provide validation to ensure users enter valid numeric inputs only, greatly enhancing user experience on the application.


![🔝Top ](#table-of-contents)
