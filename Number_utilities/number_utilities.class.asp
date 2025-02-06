<%
Class number_utilities
    ' Initialization and destruction'
	sub class_initialize()

	end sub
	
	sub class_terminate()

	end sub

    'Function to split a number when is not possible know how the interpreter works
    Private Function my_split(number)
        If InStr(number, ",") <> 0 Then 
		my_split = Split(number,",")
		Exit Function 
	End If 
	If InStr(number, ".") <> 0 Then 
		my_split = Split(number,".")
		Exit Function 
	End If 
	If is_integer(number) Then
	 	my_split = number
		Exit Function
	End If 
	Call Err.Raise(vbObjectError + 10, "free_round", "The number: " & number & " is not regular ")
    End Function

    'Function to convert a number in a array
    Private Function string_to_array(text)
        Dim length
        length = Len(text)
        Dim outArray() 
        Redim outArray(length)
        Dim index 
        For index = 0 to length - 1
            outArray(index) = Left(Right(text,(length - index)), (1))
        Next 
        string_to_array = outArray
    End Function

    'Function to split a number as a string 
    Public Function split_number(number, splitting_position)
        Dim digits
        digits = count_number_digits(number)
        Dim my_array(1)
        Dim my_number
        If splitting_position = digits Then 
            my_array(0) = number
            my_array(1) = null
            split_number = my_array
            Exit Function
        End If 
        If splitting_position < digits Then 
            If is_integer(number) Then 
                my_number = number / 10 ^ (Len(number) - splitting_position)
                split_number = my_split(my_number)
                Exit Function
            Else 
                my_number = string_to_array(number) 
                Dim index
                Dim detect 
                detect = false 
                index = splitting_position' - 1 
                If my_number(index) = "," or my_number(index) = "." Then 
                    index = splitting_position - 1
                    detect = true 
                End If 
                Dim temp 
                For temp = 0 To index
                    my_array(0) = my_array(0) & my_number(temp)
                Next
                If detect Then 
                    index = index + 1
                End If 
                For temp = index + 1 To UBound(my_number)
                    my_array(1) = my_array(1) & my_number(temp)
                Next
                split_number = my_array
                Exit Function
            End If 
        Else
            Call Err.Raise(vbObjectError + 10, "split_number", "Splitting position is not valid")
        End If 
    End Function

    'Function to check if a number is an integer
    Public Function is_integer(number)
        If InStr(number, ",") <> 0 or InStr(number, ".") <> 0 Then 
            is_integer = false
        Else
            is_integer = true
        End If
    End Function

    'Function to count number's digits
    Public Function count_number_digits(number)
        If is_integer(number) Then 
            count_number_digits = Len(number)
            Exit Function 
        Else
            If InStr(number, ",") <> 0 Then 
                count_number_digits = Len(Replace(number, ",", ""))
                Exit Function 
            End If 
            If InStr(number, ".") <> 0 Then 
                count_number_digits = Len(Replace(number, ".", ""))
                Exit Function 
            End If 
        End If 
        Call Err.Raise(vbObjectError + 10, "free_round", "The number: " & number & " is not contable")
    End Function

    'Function to free round number
    Public Function free_round(number, deciaml_to_round, number_from_starting_round)
        If Not is_integer(number) Then  
        If deciaml_to_round < count_number_digits(my_split(number)(1)) Then
            Dim my_number
            my_number = number * (10 ^ deciaml_to_round)
			Dim temp_number
			If is_integer(my_number) Then 
        			temp_number = my_number
			Else
				temp_number = my_split(my_number)(0)
				If Int(stringToArray(my_split(my_number)(1))(0)) >= number_from_starting_round Then
                			temp_number = temp_number + 1
            			End If 
			End If 
            free_round = temp_number / (10 ^ deciaml_to_round)
            Exit Function
        Else
            'Call Err.Raise(vbObjectError + 10, "free_round", "There is no decimal to round")
            free_round = number
        End If 
    Else
        'Call Err.Raise(vbObjectError + 10, "free_round", "The number is a integer") 
        free_round = number
    End If 
    End Function
End Class  
%> 
