<%
Class number_utilities
    ' Initialization and destruction'
	sub class_initialize()
        my_password = Null 
        Set my_dictionary = new dictionary
	end sub
	
	sub class_terminate()
		my_password = Null 
        my_dictionary = Null 
	end sub

    'Function to convert a string into a number
    Public Function string_to_number(str)
        Dim length
        length = Len(str)
        Dim strTemp
        strTemp = ""
        Dim index 
        Dim characters
        characters = 0
        Dim tmpArray
        tmpArray = Array()
        For index = 0 to length - 1
            Dim character
            character = Left(Right(str,(length - index)), (1))
            If character = "." or character = "," Then
                characters = characters + 1
                character = "."
            End If 
            strTemp = strTemp & character
        Next 
        If characters > 1 Then
            tmpArray = Split(strTemp, ".")
            strTemp = ""
            For index = 0 to UBound(tmpArray)
                strTemp = strTemp & tmpArray(index)
                If index = UBound(tmpArray) - 1 Then 
                    strTemp = strTemp & "." & tmpArray(index + 1)
                    Exit For
                End If 
            Next
        End If
        convert_string_to_number = strTemp
    End Function

    'Function to check if a number is an integer
    Public Function is_integer(number)
        If InStr(number, ",") Then 
            is_integer = True
            Exit Function
        Else
            is_integer = False
            Exit Function
        End If
    End Function

    'Function to count number's digits
    Public Function count_number_digits(number)
        Dim count
        count = 0
        Dim my_number
        my_number = number
        If is_integer(number) Then 
            Do While my_number > 1
                count = count + 1
                my_number = my_number / 10
            Loop
        Else
            Do While my_number > 1
                count = count + 1
                my_number = my_number / 10
            Loop
            my_number = Int(Split(number, ",")(1))
            Do While my_number > 1
                count = count + 1
                my_number = my_number / 10
            Loop
        End If 
        count_number_digits = count
    End Function

    'Function to convert a number in a array
    Private Function stringToArray(text)
        Dim length
        length = Len(text)
        Dim outArray() 
        Redim outArray(length)
        Dim index 
        For index = 0 to length - 1
            outArray(index) = Left(Right(text,(length - index)), (1))
        Next 
        stringToArray = outArray
    End Function

    'Function to free round number
    Public Function free_round(number, deciaml_to_round, number_from_starting_round)
        If Not is_integer(number) Then 
            'Response.write "STAMPA DI DEBUG: " & count_number_digits(Split(number,",")(1)) & "<br>"
            If deciaml_to_round < count_number_digits(Int(Split(number,",")(1))) Then 
                Dim my_number
                my_number = number * (10 ^ deciaml_to_round)
                If Ubound(Split(my_number, ",")) = 0 Then 
                    free_round = temp_number / (10 ^ deciaml_to_round)
                    Exit Function
                Else
                    Dim temp_number
                    temp_number = Split(my_number, ",")(0)
                    If Int(stringToArray(Split(my_number, ",")(1))(0)) >= number_from_starting_round Then
                        temp_number = temp_number + 1
                    End If 
                    free_round = temp_number / (10 ^ deciaml_to_round)
                    Exit Function
                End If
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