# Number utilities in Calssic ASP

[![Codacy Badge](https://app.codacy.com/project/badge/Grade/9b8ec61ecfc142bbbf176a745c4632e5)](https://app.codacy.com/gh/R0mb0/Number_utilities_calssic_asp/dashboard?utm_source=gh&utm_medium=referral&utm_content=&utm_campaign=Badge_grade)

[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)](https://github.com/R0mb0/Number_utilities_calssic_asp)
[![Open Source Love svg3](https://badges.frapsoft.com/os/v3/open-source.svg?v=103)](https://github.com/R0mb0/Number_utilities_calssic_asp)
[![MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/license/mit)

[![Donate](https://img.shields.io/badge/PayPal-Donate%20to%20Author-blue.svg)](http://paypal.me/R0mb0)

## `number_utilities.class.asp`'s avaible functions

- Function to split a number as a string -> `Public Function split_number(number, splitting_position)`
- Function to check if a number is an integer -> `Public Function is_integer(number)`
- Function to count number's digits -> `Public Function count_number_digits(number)`
- Function to free round a number -> `Public Function free_round(number, deciaml_to_round, number_from_starting_round)`
  >
  > - Where the decimal_to_round is the decimal position for the number to round.
  > - Where the number_from_starting_round is the round param -> For example, add 1 if the number afther the number to round is >= 5.

## How to use 

> From `Test.asp`

1. Initialize the class
   ```asp
     <%@LANGUAGE="VBSCRIPT"%>
     <!--#include file="number_utilities.class.asp"-->
     <%
        Dim utilities
        Set utilities = new number_utilities
   ```

2. Use the class
   ```asp
     Response.write("<h3> Test Split Number </h3><br>")
     Response.write("Number: 345 <br>")
     Response.write("Splitting Position: 2 <br>")
     Dim temp 
     temp = utilities.split_number(345, 2)
     Response.write("First element: " & temp(0) & "<br>")
     Response.write("Second element: " & temp(1) & "<br>")
     Response.write("Number: 12.345 <br>")
     Response.write("Splitting Position: 3 <br>")
     temp = utilities.split_number(12.345, 2)
     Response.write("First element: " & temp(0) & "<br>")
     Response.write("Second element: " & temp(1) & "<br>")
     Response.write("Number: 123.0345 <br>")
     Response.write("Splitting Position: 3 <br>")
     temp = utilities.split_number(123.0345, 2)
     Response.write("First element: " & temp(0) & "<br>")
     Response.write("Second element: " & temp(1) & "<br>")
   %>
   ```
