Attribute VB_Name = "Module1"
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
' Dim C As String
 Dim str() As String
 Dim i As Integer
' Dim x As Integer
 
 hole_1 = "ST|ASME|Blind|M16"
 hole_2 = "TH|DIN|Blind|M20"
 hole_3 = "Tk|DIm|Blinds|M24"
 
 Std_text = "Hole_Type|Standard|Sub_Type|Size"
 myarray = Array(hole_1, hole_2, hole_3)
 
str = VBA.Split(Std_text, "|")
hole_Type = str(0)
standard = str(1)
sub_Type = str(2)
size = str(3)
 
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
    
'    If x = 1 Then
'        hole_1 = Split(hole_1, "|")
'        str = VBA.Split(Std_text, "|")
'    Else:
'        hole_x = Split(hole_x, "|")
'        str = VBA.Split(Std_text, "|")
'    End If
'
'            hole_Type = str(0)
'            standard = str(1)
'            sub_Type = str(2)
'            size = str(3)
'Next x
'
End Sub
