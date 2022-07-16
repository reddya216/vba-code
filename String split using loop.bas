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
' Dim x As Integer
' These are the strings needed to be splitted

 hole_1 = "ST|ASME|Blind|M16"
 hole_2 = "TH|DIN|Blind|M20"
 hole_3 = "Tk|DIm|Blinds|M24"
 
' here I have created a std_text to store the splitted strings in the format described below

 Std_text = "Hole_Type|Standard|Sub_Type|Size"
 myarray = Array(hole_1, hole_2, hole_3)
 
'Created a array(myarray) to pass the strings which defined above

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
