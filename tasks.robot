*** Settings ***
Documentation       Use case 2 building.
...                 Read the excel and get the value of conversion from google.

Library             RPA.Browser.Selenium    auto_close=${False}
Library             RPA.Excel.Files
Library             RPA.Robocloud.Items
Library             RPA.RobotLogListener
Library             RPA.Excel.Application
Library             RPA.Tables
Library             Collections
Library             RPA.FileSystem
Library             Py.py
Library             RPA.Salesforce

Task Setup          Open Application
Task Teardown       Quit Application


*** Variables ***
${URL}                          https://www.google.com/
${what_to_search}               Currency Converter
${excel_path}                   C:\\Users\\raavi.padmavathi\\Documents\\CurrencyConverter.xlsx
${counter}                      ${2}
${Col_counter}                  ${2}
@{USD_values}
@{Euro_values}
@{row_counter}                  2
${USD_doller_locator1}
...                             xpath://html/body/div[7]/div/div[10]/div/div[2]/div[2]/div/div/div[1]/div/div/div/div/div/div/div[3]/div[1]/div[3]/div/div[2]/input
${USD_doller_locator2}
...                             xpath://html/body/div[7]/div/div[10]/div/div[2]/div[2]/div/div/div[1]/div/div/div/div/div/div/div[3]/div[1]/div[3]/div/div[2]/div/select
${Indian_ruppes_locator1}
...                             xpath://html/body/div[7]/div/div[10]/div/div[2]/div[2]/div/div/div[1]/div/div/div/div/div/div/div[3]/div[1]/div[3]/div/div[1]/input
${Indian_ruppes_locator2}
...                             xpath://html/body/div[7]/div/div[10]/div/div[2]/div[2]/div/div/div[1]/div/div/div/div/div/div/div[3]/div[1]/div[3]/div/div[1]/div/select


*** Tasks ***
Open Browser and the conversion rate from website
    Open the browser and perform task
    Open the excel and get the values in table
    Write back to excel sheet    ${USD_values}    ${Col_counter}
    ${Col_counter}=    py.increment_var    ${Col_counter}
    Write back to excel sheet    ${Euro_values}    ${Col_counter}

Minimal task
    Log    Done.


*** Keywords ***
Open the browser and perform task
    Create List    @{USD_values}
    Create List    @{Euro_values}
    Open Available Browser    ${URL}
    Input Text    //input[@title='Search']    ${what_to_search}
    Submit Form
    Maximize Browser Window
    Sleep    5

Open the excel and get the values in table
    RPA.Excel.Files.Open Workbook    ${excel_path}
    @{table}=    Read Worksheet As Table    header=True

    Log    ${table}
    Close Workbook
    # FOR    ${row}    IN    @{table}
    #    select the conversion and get value for USD    ${row}
    #    select the conversion and get value for Euro    ${row}
    #    #Write back to excel sheet    ${excel_path}
    #    Log    ${row}[Indian Rupee]
    # END

select the conversion and get value for USD
    [Arguments]    ${row}
    Sleep    5
    Select From List By Label    ${USD_doller_locator2}    United States Dollar
    Sleep    5
    Input Text    ${Indian_ruppes_locator1}    ${row}[Indian Rupee]
    #${USD_value}=    Get Text    //table[@class='qzNNJ']/tbody/tr[3]/td[1]/input
    ${value_USD}=    Get Element Attribute    ${USD_doller_locator1}    value
    Append To List    ${USD_values}    ${value_USD}
    Log    ${value_USD}
    Log    ${USD_values}

select the conversion and get value for Euro
    [Arguments]    ${row}
    Sleep    5
    Select From List By Label    ${USD_doller_locator2}    Euro
    Sleep    5
    Input Text    ${Indian_ruppes_locator1}    ${row}[Indian Rupee]
    ${value_EURO}=    Get Text    ${USD_doller_locator1}
    Append To List    ${Euro_values}    ${value_EURO}
    Log    ${value_EURO}
    Log    ${Euro_values}

Write back to excel sheet
    [Arguments]    ${lists}    ${Col_counter}
    ${decrement_counter}=    set variable    1
    RPA.Excel.Application.Open Workbook    ${excel_path}
    RPA.Excel.Application.Set Active Worksheet    Sheet1

    FOR    ${element}    IN    @{lists}
        RPA.Excel.Application.Write To Cells    row=${counter}    column=${Col_counter}
        ...    value=${element}
        RPA.Excel.Application.Save Excel
        ${counter}=    py.increment_var    ${counter}
        Log    ${element}
    END
