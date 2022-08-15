import asyncio
import sys
from js import alert, document, Object, window
from js import DOMParser, document, setInterval
from pyodide import create_proxy, to_js

from pyodide.http import open_url

from openpyxl import load_workbook

save_target = None


async def file_select(event):
    # Note: print() does not work in event handlers
    global save_target

    try:
        options = {
            "startIn": "documents",
            "suggestedName": "new.xlsx"
        }

        save_target = await window.showSaveFilePicker(Object.fromEntries(to_js(options)))
    except Exception as e:
        console.log('Exception: ' + str(e))
        return

    file = await save_target.createWritable()

    content = open_url("./fibonacci.xlsm").read()
    wb = openpyxl.load_workbook('example.xlsx')
    wb.get_sheet_names()
    console.log('Contests: ' + content)
    await file.write(content)
    await file.close()


async def file_save(event):
    global save_target

    if save_target:
        file = await save_target.createWritable()
        content = document.getElementById("content").value
        console.log('Contests: ' + content)
        await file.write(content)
        await file.close()
    else:
        alert("Please select a file first")


def setup_button():
    # Create a Python proxy for the callback function
    file_select_proxy = create_proxy(file_select)
    file_save_proxy = create_proxy(file_save)

    # Set the listener to the callback
    document.getElementById("select_file").addEventListener("click", file_select_proxy, False)
    document.getElementById("file_save").addEventListener("click", file_save_proxy, False)


setup_button()
