*** Settings ***
Documentation       Template robot main suite.

Library             RPA.Excel.Files
Library             RPA.Excel.Application
Library             String
Library             RPA.Tables
Library             Collections
Library             RPA.Browser
Library             fun.py
Library             RPA.Outlook.Application
Library             RPA.Archive
Library             DateTime
Library             OperatingSystem
Library             RPA.FileSystem



*** Variables ***
@{FileName-List}            StudySample.xlsx    StudySample.xlsx
@{Sheet-List}
@{Input-File-SheetName}     Sheet1    Sheet1
${FILE}                     Study#.xlsx
@{sheet}                    Mini_Avg    Mini_Stdev
${Matched_DOE}              Matched_DOE.xlsx
${ROW}                      Set Variable    2
${COLUMN}                   3
${attachmentFilePath}       D:/RoboCorp Task/MVDA BOT/Input File
${inputFileFolderPath}      D:/RoboCorp Task/Input Folder
${VID_Compare}              VID_Alignment
${stdNumber}
@{ExceptionList}


*** Tasks ***
MVDA Steps
    TRY
        
        ${studyNumberList}=    Get Study Number    ${attachmentFilePath}
        FOR    ${studyNumber}    IN    @{studyNumberList}
            TRY
                ${sheetList}=    Check Input Files in Input Folder    ${studyNumber}
                Log To Console    ${sheetList}[0]
                Create File    ${FILE}    ${sheetList}[0][0]
                Getting Master and Minion Sheets data into Study File
                ...    ${sheetList}[0]
                ...    ${inputFileFolderPath}
                ...    ${studyNumber}
                Create VID_Compare Sheet and Compare Columns    ${VID_Compare}    ${sheetList}[0]
                ${Matched_Data}=    Read Matched_DOE File ${sheetList}[1]
                Create Avg and Stdev Sheets ${Matched_Data}
                Create Matched_Sheets ${Matched_Data}
            EXCEPT    AS    ${exception}
                Append To List    ${ExceptionList}    ${exception}
            END
        END
    EXCEPT    AS    ${exception}
        Append To List    ${ExceptionList}    ${exception}
    END
    Log To Console    ${ExceptionList}
    Create File    Summary Report.xlsx    Summary Report
    Paste WorkSheet Data    Summary Report.xlsx    Summary Report    ${ExceptionList}
    RPA.Excel.Files.Open Workbook    Summary Report.xlsx
    RPA.Excel.Files.Save Workbook    ${inputFileFolderPath}${/}${stdNumber}${/}Summary Report.xlsx 
    Close Workbook

*** Keywords ***
Check Input Files in Input Folder
    [Arguments]    ${std}
    Should Exist    ${inputFileFolderPath}/${std}    Business Exception : Study Number: ${std}: Folder not found
    Should Exist
    ...    ${inputFileFolderPath}/${std}/*Matched*.xlsx
    ...    Business Exception : Study Number: ${std} : Matched_DOE.xlsx File is not Exist
    Should Exist
    ...    ${inputFileFolderPath}/${std}/*Mastered_Centered*.xlsx
    ...    Business Exception : Study Number: ${std} :Mastered_Centered.xlsx File is not Exist
    Should Exist
    ...    ${inputFileFolderPath}/${std}/*Minion_DOE*.xlsx
    ...    Business Exception : Study Number: ${std} : Minion_DOE.xlsx File is not Exist
    @{Sheets}=    Create List
    ${Mastered}=    Find Files    ${inputFileFolderPath}/${std}/*Mastered_Centered*.xlsx
    ${Sheet}=    Validate    ${std}    ${Mastered}[0][0]
    Append To List    ${Sheets}    ${Sheet}
    ${Mastered}=    Find Files    ${inputFileFolderPath}/${std}/*Minion_DOE*.xlsx
    ${Sheet}=    Validate    ${std}    ${Mastered}[0][0]
    Append To List    ${Sheets}    ${Sheet}

    ${Mastered}=    Find Files    ${inputFileFolderPath}/${std}/*Matched*.xlsx
    RPA.Excel.Files.Open Workbook    ${Mastered}[0][0]
    ${data}=    Read Worksheet As Table
    Close Workbook
    ${headers}=    Get Table Row    ${data}    0    True
    Should Be Equal    ${headers}[0]    Matched DOE #    Column A is invalid in ${Mastered}[0][0] file
    Should Be Equal    ${headers}[1]    Trail Combinatons    Column B is invalid in ${Mastered}[0][0] file

    FOR    ${sheet}    IN    @{Sheets}
        Append To List    ${Sheet-List}    ${sheet}
    END
    ${stdNumber}=    Set Variable    ${std}
    RETURN    ${Sheets}    ${Mastered}[0][0]

Validate
    [Arguments]    ${studyNum}    ${WorkBook}
    RPA.Excel.Files.Open Workbook    ${WorkBook}
    ${data}=    Read Worksheet As Table
    Close Workbook
    Pop Table Column    ${data}    A
    Pop Table Column    ${data}    B
    ${headers}=    Get Table Row    ${data}    0    True

    ${Mastered}=    Get File Stem    ${WorkBook}

    FOR    ${col}    IN    @{headers}
        Should Contain
        ...    ${col}
        ...    VID_
        ...    Business Exception : Study Number : ${studyNum} - ${Mastered} file has invalid column
    END
    RETURN    ${Mastered}

Getting Master and Minion Sheets data into Study File
    [Arguments]    ${workBook}    ${inputFileFolderPath}    ${std}
    ${length}=    Evaluate    len(${workBook})
    ${length}=    Evaluate    ${length} - 1

    WHILE    ${length} >=0
        ${WorksheetData}=    Read Input WorkSheet
        ...    ${inputFileFolderPath}/${std}/${workBook}[${length}].xlsx
        ...    ${Input-File-SheetName}[${length}]
        Paste WorkSheet Data    ${FILE}    ${workBook}[${length}]    ${WorksheetData}
        ${length}=    Evaluate    ${length} - 1
    END

Create VID_Compare Sheet and Compare Columns
    [Arguments]    ${VID_Compare}    ${sheetList}
    #${VID_Compare}=    Set Variable    VID_Compare
    ${row}=    Set Variable    1
    @{l1}=    Set Variable    1    3    5
    @{l2}=    Set Variable    2    4    6
    @{l3}=    Set Variable    ${l1}    ${l2}
    ${column_index}=    Set Variable    3
    ${Column-Data}=    Create Work Sheet in File    ${FILE}    ${VID_Compare}

    FOR    ${SheetName}    IN    @{sheetList}
        FOR    ${counter}    IN RANGE    1    3
            ${Column-Data}=    Get Columns Data from ${SheetName} ${FILE} ${counter} ${column_index}
            Write Columns in ${VID_Compare} ${FILE} ${Column-Data} ${row}
            ${row}=    Evaluate    ${row} + 1
        END
    END

    FOR    ${element}    IN    @{l3}
        Compare Column Name    ${VID_Compare}    ${element}
        
    END

Create File
    [Arguments]    ${File-Name}    ${workSheet}
    RPA.Excel.Files.Create Workbook    sheet_name=${workSheet}
    RPA.Excel.Files.Save Workbook    ${File-Name}

Create Work Sheet in File
    [Arguments]    ${WorkBook}    ${WorkSheet}
    RPA.Excel.Files.Open Workbook    ${WorkBook}
    ${WorkSheet-Exist}=    RPA.Excel.Files.Worksheet Exists    ${WorkSheet}
    IF    ${WorkSheet-Exist}
        Log To Console    Worksheet already exist
    ELSE
        Create Worksheet    ${WorkSheet}
    END
    Save Workbook    ${WorkBook}
    Close Workbook

Read Input WorkSheet
    [Arguments]    ${WorkBook_Name}    ${SheetName}
    RPA.Excel.Files.Open Workbook    ${WorkBook_Name}
    ${File1}=    Read Worksheet    ${SheetName}
    Close Workbook
    RETURN    ${File1}

Paste WorkSheet Data
    [Arguments]    ${WorkBook}    ${SheetName}    ${SheetData}
    RPA.Excel.Files.Open Workbook    ${WorkBook}
    ${WorkSheet-Exist}=    Worksheet Exists    ${SheetName}
    IF    ${WorkSheet-Exist}
        RPA.Excel.Files.Set Active Worksheet    ${SheetName}
        Append Rows To Worksheet    ${SheetData}
    ELSE
        Create Worksheet    ${SheetName}    ${SheetData}
        #RPA.Excel.Files.Set Active Worksheet    ${SheetName}
        #Append Rows To Worksheet    ${SheetData}
    END
    Save Workbook    ${WorkBook}
    Close Workbook

Get Columns Data from ${WorkSheet} ${WorkBook} ${rowindex} ${columnindex}
    ${Data}=    Set Variable

    ${Check}=    Convert To Boolean    True
    @{Column-Data}=    Create List
    RPA.Excel.Files.Open Workbook    ${WorkBook}
    RPA.Excel.Files.Set Active Worksheet    ${WorkSheet}

    WHILE    ${Check}
        ${Data}=    Get Cell Value    ${rowindex}    ${columnindex}
        IF    $Data==None
            Append To List    ${Column-Data}    ${EMPTY}
            ${columnindex}=    Evaluate    ${columnindex} + 1
            ${Data}=    Get Cell Value    ${rowindex}    ${columnindex}
            IF    $Data==None
                ${Check}=    Convert To Boolean    False
            ELSE
                Append To List    ${Column-Data}    ${Data}
                ${Check}=    Convert To Boolean    True
            END
        ELSE
            Append To List    ${Column-Data}    ${Data}
            ${Check}=    Convert To Boolean    True
        END
        ${columnindex}=    Evaluate    ${columnindex} + 1
    END
    #Log To Console    Column Data: -${Column-Data}
    Close Workbook
    RETURN    ${Column-Data}

Write Columns in ${WorkSheet} ${WorkBook} ${Col-Name} ${row-index}
    ${column}=    Set Variable    1
    RPA.Excel.Files.Open Workbook    ${WorkBook}
    RPA.Excel.Files.Set Active Worksheet    ${WorkSheet}
    FOR    ${Col}    IN    @{Col-Name}
        Set Cell Value    ${row-index}    ${column}    ${Col}
        ${column}=    Evaluate    ${column} + 1
    END
    Save Workbook    ${WorkBook}
    Close Workbook

Compare Column Name
    [Arguments]    ${WorkSheet}    ${row}
    RPA.Excel.Files.Open Workbook    ${FILE}
    RPA.Excel.Files.Set Active Worksheet    ${WorkSheet}
    ${colCounter}=    Set Variable    1
    ${Check}=    Set Variable    ${True}
    WHILE    ${Check}
        ${row1}=    Get Cell Value    ${row}[0]    ${colCounter}
        ${row2}=    Get Cell Value    ${row}[1]    ${colCounter}
        IF    $row1==None and $row2==None
            ${Check}=    Convert To Boolean    False
        ELSE
            Should Be Equal
            ...    ${row1}
            ...    ${row2}
            ...    Business Exception: Study Number : ${stdNumber} : Column is not equal
            Set Cell Value    ${row}[2]    ${colCounter}    True
            Save Workbook    ${FILE}

            ${Check}=    Convert To Boolean    True
        END
        ${colCounter}=    Evaluate    ${colCounter} + 1
    END
    Save Workbook    ${FILE}
    Close Workbook

Create Avg and Stdev Sheet
    [Arguments]    ${SheetList}    ${Matched_DOE_Data}
    ${column_index}=    Set Variable    1
    ${DataSet}=    Set Variable    Dataset
    ${Statistics}=    Set Variable    Statistics

    FOR    ${Worksheet}    IN    @{SheetList}
        Create Work Sheet in File    ${FILE}    ${Worksheet}
    END

    FOR    ${SheetName}    IN    @{SheetList}
        FOR    ${counter}    IN RANGE    1    3
            ${Column-Data}=    Get Columns Data from ${Sheet-List}[1] ${FILE} ${counter} ${column_index}
            Write Columns in ${SheetName} ${FILE} ${Column-Data} ${counter}
        END
    END
    RPA.Excel.Files.Open Workbook    ${FILE}
    FOR    ${SheetName1}    IN    @{SheetList}
        RPA.Excel.Files.Set Active Worksheet    ${SheetName1}
        FOR    ${row}    IN RANGE    1    3
            FOR    ${column}    IN RANGE    1    3
                IF    $column == 2
                    Set Cell Value    ${row}    ${column}    ${Statistics}
                ELSE
                    Set Cell Value    ${row}    ${column}    ${DataSet}
                END
            END
        END
    END

    ${Row}=    Set Variable    3
    ${Average}=    Set Variable    Average
    ${Standard Deviation}=    Set Variable    Standard Deviation
    ${length}=    Evaluate    len(${Matched_DOE_Data})

    RPA.Excel.Files.Set Active Worksheet    ${SheetList}[0]
    FOR    ${element}    IN    @{Matched_DOE_Data}
        FOR    ${column}    IN RANGE    1    3
            IF    $column == 2
                Set Cell Value    ${Row}    ${column}    ${Average}
            ELSE
                Set Cell Value    ${Row}    ${column}    ${element}
            END
        END
        ${Row}=    Evaluate    ${Row} + 1
    END

    ${Row}=    Set Variable    3
    RPA.Excel.Files.Set Active Worksheet    ${SheetList}[1]
    FOR    ${element}    IN    @{Matched_DOE_Data}
        FOR    ${column}    IN RANGE    1    3
            IF    $column == 2
                Set Cell Value    ${Row}    ${column}    ${Standard Deviation}
            ELSE
                Set Cell Value    ${Row}    ${column}    ${element}
            END
        END
        ${Row}=    Evaluate    ${Row} + 1
    END

    Save Workbook
    Close Workbook

Read Matched_DOE File ${FileName}
    @{Matched DOE #}=    Create List
    @{Trail Combinaton}=    Create List

    RPA.Excel.Files.Open Workbook    ${FileName}
    ${Data}=    Read Worksheet    start=2
    Close Workbook
    FOR    ${element}    IN    @{Data}
        Append To List    ${Matched DOE #}    ${element}[A]
    END
    FOR    ${element}    IN    @{Data}
        Append To List    ${Trail Combinaton}    ${element}[B]
    END
    RETURN    ${Matched DOE #}    ${Trail Combinaton}

Calculate Average
    [Arguments]    ${Trail_Combination}    ${WorkSheet}    ${Column}

    RPA.Excel.Application.Open Application    
    RPA.Excel.Application.Open Workbook    ${FILE}
    RPA.Excel.Application.Set Active Worksheet    ${WorkSheet}
    ${Check}=    Convert To Boolean    True
    ${Row}=    Set Variable    2
    ${Col}=    Set Variable    B
    WHILE    ${Check}
        Log To Console    ${Row}
        Log To Console    ${Column}
        ${Data}=    Read From Cells    ${WorkSheet}    ${Row}    ${Column}
        Log To Console    Data:-${Data}
        IF    $Data == None
            ${Check}=    Convert To Boolean    False
        ELSE
            ${Row}=    Evaluate    ${Row} + 1
            Log To Console    Row:-${Row}
            ${Check}=    Convert To Boolean    True
            ${Col}=    Get Excel Col    ${Col}
            FOR    ${combination}    IN    @{Trail_Combination}
                Log To Console    ${combination}

                RPA.Excel.Application.Write To Cells
                ...    ${WorkSheet}
                ...    ${Row}
                ...    ${Column}
                ...    formula==AVERAGE(IF(${Sheet-List}[1]!$B:$B={${combination}},${Sheet-List}[1]!${Col}:${Col}))
                Log To Console
                ...    AVERAGE(IF(${Sheet-List}[1]!$B:$B={${combination}},${Sheet-List}[1]!${Col}:${Col}))
                ${Row}=    Evaluate    ${Row} + 1
            END
        END
        ${Column}=    Evaluate    ${Column} + 1
        ${Col}=    Set Variable    ${Col}
        ${Row}=    Set Variable    2
        Save Excel
    END
    Save Excel
    RPA.Excel.Application.Close Document    True

Calculate StDev
    [Arguments]    ${Trail_Combination}    ${WorkSheet}    ${Column}

    RPA.Excel.Application.Open Application    
    RPA.Excel.Application.Open Workbook    ${FILE}
    RPA.Excel.Application.Set Active Worksheet    ${WorkSheet}
    ${Check}=    Convert To Boolean    True
    ${Row}=    Set Variable    2
    ${Col}=    Set Variable    B
    WHILE    ${Check}
        Log To Console    ${Row}
        Log To Console    ${Column}
        ${Data}=    Read From Cells    ${WorkSheet}    ${Row}    ${Column}
        Log To Console    Data:-${Data}
        IF    $Data == None
            ${Check}=    Convert To Boolean    False
        ELSE
            ${Row}=    Evaluate    ${Row} + 1
            Log To Console    Row:-${Row}
            ${Check}=    Convert To Boolean    True
            ${Col}=    Get Excel Col    ${Col}
            FOR    ${combination}    IN    @{Trail_Combination}
                Log To Console    ${combination}

                RPA.Excel.Application.Write To Cells
                ...    ${WorkSheet}
                ...    ${Row}
                ...    ${Column}
                ...    formula==STDEV(IF(${Sheet-List}[1]!$B:$B={${combination}},${Sheet-List}[1]!${Col}:${Col}))
                Log To Console
                ...    STDEV(IF(${Sheet-List}[1]!$B:$B={${combination}},${Sheet-List}[1]!${Col}:${Col}))
                ${Row}=    Evaluate    ${Row} + 1
            END
        END
        ${Column}=    Evaluate    ${Column} + 1
        ${Col}=    Set Variable    ${Col}
        ${Row}=    Set Variable    2
        Save Excel
    END
    Save Excel
    RPA.Excel.Application.Close Document    True

Create Avg and Stdev Sheets ${Matched_DOE}
    Create Avg and Stdev Sheet    ${sheet}    ${Matched_DOE}[0]
    Calculate Average    ${Matched_DOE}[1]    ${sheet}[0]    ${COLUMN}
    Calculate StDev    ${Matched_DOE}[1]    ${sheet}[1]    ${COLUMN}


Create Matched_Sheets ${Matched_DOE}
    @{Matched_Sheets}=    Create List
    FOR    ${Worksheet}    IN    @{Matched_DOE}[0]
        Create Work Sheet in File    ${FILE}    Matched_DOE${Worksheet}
        Append To List    ${Matched_Sheets}    Matched_DOE${Worksheet}
    END
    # Log To Console    ${Matched_Sheets}
    RPA.Excel.Files.Open Workbook    ${FILE}
    RPA.Excel.Files.Set Active Worksheet    ${Sheet-List}[0]

    ${Master_Table}=    Read Worksheet As Table    ${Sheet-List}[0]
    ${Col1}=    Get Table Column    ${Master_Table}    0
    ${Col2}=    Get Table Column    ${Master_Table}    1

    ${Row1}=    Get Table Row    ${Master_Table}    0    True
    ${Row2}=    Get Table Row    ${Master_Table}    1    True

    FOR    ${Worksheet}    IN    @{Matched_Sheets}
        Write Row    ${FILE}    ${Worksheet}    ${Row1}
        Write Row    ${FILE}    ${Worksheet}    ${Row2}
        Write Verticle    ${FILE}    ${Worksheet}    1    ${Col1}
        Write Verticle    ${FILE}    ${Worksheet}    2    ${Col2}
    END

    ${len}=    Evaluate    len(${Row1})
    Log To Console    ${len}
    ${Val}=    Set Variable    3
    FOR    ${sheet}    IN    @{Matched_Sheets}
        ${colHeader}=    Get Excel Col    A
        FOR    ${row}    IN RANGE    3    ${len}+1
            ${colHeader}=    Get Excel Col    ${colHeader}
            #Log To Console    ${row}:${colHeader}
            Apply Formula On Matched Sheets    ${FILE}    ${sheet}    ${colHeader}    ${Sheet-List}[0]    ${Val}
        END
        ${Val}=    Evaluate    ${Val}+1
    END
    Close Workbook   
    RPA.Excel.Files.Save Workbook    ${inputFileFolderPath}${/}${stdNumber}${/}${FILE}

Get Study Number
    [Arguments]    ${attachmentFilePath}
    ${time}=    Get Current Date    result_format=%d%m%Y
    ${fileName}=    Set Variable    ${attachmentFilePath}/MVDA Input - ${time}.xlsx

    Should Exist    ${fileName}    ${fileName} is not found

    RPA.Excel.Files.Open Workbook    ${fileName}
    ${data}=    Read Worksheet As Table
    Close Workbook
    ${headers}=    Get Table Row    ${data}    0    True
    Should Be Equal    ${headers}[0]    Study Number    Business Exception :Column A is invalid in Attachment file
    Should Be Equal    ${headers}[1]    Priority    Business Exception :Column B is invalid in Attachment file
    Pop Table Row    ${data}
    Sort Table By Column    ${data}    1
    ${studyNumbersList}=    Get Table Column    ${data}    0
    Should Not Be Empty
    ...    ${studyNumbersList}
    ...    Business Exception : Attachment file named MVDA Input - ${time}.xlsx has No Study Number
    Log To Console    ${studyNumbersList}

    RETURN    ${studyNumbersList}
