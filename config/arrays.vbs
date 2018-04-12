'Using the Option Base Statement'
Option Base 1 '<- All arrays declared in the module start counting at 1
              ' Option Base only takes 0 and 1


'Arrays let you hold more than one piece
'of data into the same variable

Sub fixedSizeArray()

  'You declare the size of the array upfront'
  Dim AnArray(2) As String

  'You can also define the upper and lower bounds
  'Of the array at declaration to avoid using Option Base
  Dim AnotherArray(1 to 5) As String

  'Populating an Array'

' You need to declare which element of the array
' You want to populate

  AnArray(1) = Range("A1")
  AnArray(2) = Range("A2")

  'Erasing an Array'

  Erase AnArray '<- clears array contents, does not eliminate variable'

End sub

'Looping over an Array'

Sub LoopOverArray ()

    Dim FirstTenNumbers(1 to 10) as Integer
    Dim Element as Integer

    For Each Element in TopTenScores
      Range("A" & Element).Value = Element
    Next Element

    'You can also refer to the upperbound and the lowerbound
    'Of the array to avoid hardcoded constants

    Dim FirstTenNumbers(1 to 10) as Integer
    Dim Counter as Integer

    For Counter = lbound(FirstTenNumbers) to ubound(FirstTenNumbers)
      Range("A" & Counter).Value = Counter
    Next Counter

End Sub


'Declaring MultiDimensional Arrays'

Sub MultiDimensionArray ()

  'We use Variant as we might store different data types'
  Dim TopTenFilms(0 to 9, 0 to 4) as Variant
    Dim(0, 0 ) = Range("A1").Value
    Dim(0, 1 ) = Range("A2").Value
    Dim(0, 2 ) = Range("A3").Value

End Sub


'Looping over MultiDimensional Array'


Sub MultiDimensionalArrayLoop()

  Dim TopTenFilms(0 to 9, 0 to 4) as Variant
  Dim Dimension1 as Long, Dim Dimension2 as Long

  for Dimension1 = 0 to 9
    for Dimension2 = 0 to 4
      TopTenFilms(Dimension1, Dimension2) = Range("A1").Offset(Dimension1, Dimension2).Value
    Next Dimension2
  Next Dimension1

  'or even better, with lbound and ubound references'

  for Dimension1 = lbound(TopTenFilms, 1) to ubound(TopTenFilms, 1)
    for Dimension2 =  lbound(TopTenFilms, 2) to ubound(TopTenFilms, 2)
       TopTenFilms(Dimension1, Dimension2) = Range("A1").Offset(Dimension1, Dimension2).Value
    Next Dimension2
  Next Dimension1
End Sub

  'Or Even better, using sheet references
  'To avoid hardcoding constants altogether

Sub MultiDimensionalArrayLoop()

  Dim FlexibleArray() As Variant
  Dim Dimension1 as Long
  Dim Dimension2 as Long

  Dimension1 = Range("A1", Range("A1").End(Xldown)).Cells.Count
  Dimension2 = Range("A1", Range("A1").End(XlToRight)).Cells.Count

  Redim FlexibleArray(1 to Dimension1, 1 to Dimension2)

  For Dimension1 = FlexibleArray(lbound, 1) to FlexibleArray(ubound,1)
    For Dimension2 = FlexibleArray(lbound, 2) to FlexibleArray(ubound, 2)

    debug.print(FlexibleArray(Dimension1, Dimension2).Value)

    Next Dimension2
  Next Dimension1

End Sub

'Writing a range to a multidimensional array'

Sub QuickMultiDimensionalArrayLoop()

  Dim  QuickFlexibleArray() As Variant

  'Watch out, this method starts counting at 1, not 0'

  QuickFlexibleArray = Range("A1", Range("A1").End(XlDown).End(XlToRight))

  Erase QuickFlexibleArray '<- All spaces get deallocated
                           '   when erasing dynamic arrays like this one
End Sub


'Resizing Arrays Dynamically'

Sub ResizeDynamicArray()

  Dim ActionFilms() as Variant
  Dim Cell as Range
  Dim ActionCounter as Long, LoopCounter as Long

  For Each Cell in Range("A1",Range("A1").End(Xldown)
    If Cell.Offset(0,3).value = "Action" Then

    ActionCounter = ActionCounter + 1

    'Preserve allows you to retain the old value of the array
    'But you can only redimension the last dimension of the array
    'when you use it, forcing you to transpose the array'

    Redim Preserve ActionFilms(1 to 5, 1 to ActionCounter)

      For LoopCounter = 1 to 5

        ActionFilms(ActionCounter, LoopCounter) = cell.offset(0,LoopCounter - 1).Value

      Next LoopCounter

    End If
  Next Cell

  'Transposing the Array when adding the cells into a new sheet'

  Application.transpose.ActionFilms


End Sub







End Sub





