<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="number_utilities.class.asp"-->
<%
    Dim utilities
    Set utilities = new number_utilities

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

    Response.write("<h3> Test Is Integer </h3><br>")
    Response.write("Number: 345 <br>")
    Response.write(utilities.is_integer(345) & "<br>")
    Response.write("Number: 12.345 <br>")
    Response.write(utilities.is_integer(12.345) & "<br>")
    Response.write("Number: 123.0345 <br>")
    Response.write(utilities.is_integer(123.0345) & "<br>")

    Response.write("<h3> Test Count Number Digits </h3><br>")
    Response.write("Number: 345 <br>")
    Response.write(utilities.count_number_digits(345) & "<br>")
    Response.write("Number: 12.345 <br>")
    Response.write(utilities.count_number_digits(12.345) & "<br>")
    Response.write("Number: 123.0345 <br>")
    Response.write(utilities.count_number_digits(123.0345) & "<br>")

    Response.write("<h3> Test Free Round </h3><br>")
    Response.write("<h4> Round at second position with >= 5 as criteria </h4><br>")
    Response.write("Number: 345 <br>")
    Response.write(utilities.free_round(345, 2, 5) & "<br>")
    Response.write("Number: 12.346 <br>")
    Response.write(utilities.free_round(12.346, 2, 5) & "<br>")
    Response.write("Number: 123.0365 <br>")
    Response.write(utilities.free_round(123.0365, 2, 5) & "<br>")
%>