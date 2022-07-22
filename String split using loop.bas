Attribute VB_Name = "Module1"
'Ashok Reddy Tangirala dated 15/07/22
'This is the macro example to split the multiple strings using loop
'
'Declare all variables

Sub StringToArray()
'Option Explicit

' Dim txt As String
 Dim hole_1 As String
 Dim hole_2 As String
 Dim hole_3 As String
 Dim myarray As Variant
 Dim hole_Type As String
 Dim standard As String
 Dim sub_Type As String
 Dim size As String
 Dim hole_split() As String
 Dim str() As String
 Dim i As Integer
 
 ' Input
 hole_1 = "ST|ASME|Blind|M16"
 hole_2 = "TH|DIN|Blind|M20"
 hole_3 = "Tk|DIm|Blinds|M24"
 
 ' Declaration 
 Std_text = "Hole_Type|Standard|Sub_Type|Size"    'Comment - This code is doing nothing
 
 ' Get the string and store in array
  myarray = Array(hole_1, hole_2, hole_3)
 
 ' Split the string and assign variables  - Created a array(myarray) to pass the strings which defined above
str = VBA.Split(Std_text, "|")
hole_Type = str(0)
standard = str(1)
sub_Type = str(2)
size = str(3)
 
'created a loop which allows to pass multiple strings
 i = 2
 For x = 0 To i
    hole_array = Split(myarray(x))
    For j = LBound(hole_array) To UBound(hole_array)
        'hole_split = VBA.Split(txt, "|" & i)
         
         Hole = VBA.Split(myarray(x), "|")
          hole_Type = Hole(0)
          standard = Hole(1)
          sub_Type = Hole(2)
          size = Hole(3)
    Next j
 Next x

End Sub

' ---------------------------------------- NEW
Sub StringToArray()
'Option Explicit

' Dim txt As String
 Dim hole_1 As String
 Dim hole_2 As String
 Dim hole_3 As String
 Dim myarray As Variant
 Dim hole_Type As String
 Dim standard As String
 Dim sub_Type As String
 Dim size As String
 Dim hole_split() As String
 Dim str() As String
 Dim i As Integer
 
 ' Input
 hole_1 = "ST|ASME|Blind|M16"
 hole_2 = "TH|DIN|Blind|M20"
 hole_3 = "CB|DIm|Blinds|M24"
 
 ' Declaration
 Std_text = "Hole_Type|Standard|Sub_Type|Size"
 ' Get the string and store in array
  myarray = Array(hole_1, hole_2, hole_3)
 
 ' Split the string and assign variables  - Created a array(myarray) to pass the strings which defined above
str = VBA.Split(Std_text, "|")
A = str(0)
B = str(1)
C = str(2)
D = str(3)
 
'created a loop which allows to pass multiple strings
 i = 2
 For x = 0 To i
    hole_array = Split(myarray(x))
    For j = LBound(hole_array) To UBound(hole_array)
        'hole_split = VBA.Split(txt, "|" & i)
         
         Hole = VBA.Split(myarray(x), "|")
         '
          AA = A & " : " & Hole(0)
          BB = B & " : " & Hole(1)
          CC = C & " : " & Hole(2)
          DD = D & " : " & Hole(3)
          
    Next j
 Next x

End Sub
