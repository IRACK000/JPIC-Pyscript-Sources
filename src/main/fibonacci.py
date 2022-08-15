"""
Copyright (C) 2014-2016, Zoomer Analytics LLC.
All rights reserved.

License: BSD 3-clause (see LICENSE.txt for details)
"""
import xlwings as xw


async def fibonacci(n):
    """
    Generates the first n Fibonacci numbers.
    Adopted from: https://docs.python.org/3/tutorial/modules.html
    """
    result = []
    a, b = 0, 1
    while len(result) < n:
        result.append(b)
        a, b = b, a + b
    return result


async def xl_fibonacci():
    """
    This is a wrapper around fibonacci() to handle all the Excel stuff
    """
    # Create a reference to the calling Excel Workbook
    sht = xw.Book.caller().sheets[0]

    # Get the input from Excel and turn into integer
    n = sht.range('B1').options(numbers=int).value

    # Call the main function
    seq = fibonacci(n)

    # Clear output
    sht.range('C1').expand('vertical').clear_contents()

    # Return the output to Excel in column orientation
    sht.range('C1').options(transpose=True).value = seq


def main():
    path = r"files:///C:\Users\IRACK\Desktop\JPIC-Pyscript-Sources\src\main"
    xw.Book(path+"/"+"fibonacci.xlsm").set_mock_caller()
    xl_fibonacci()
