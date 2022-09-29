*** Settings ***
Documentation       Template robot main suite.

Library    RPA.Windows    WITH NAME    w
Library    RPA.Desktop    WITH NAME    d
Library    RPA.Excel.Files    WITH NAME    e
Library    RPA.FileSystem    WITH NAME    f
Library    RPA.Tables    WITH NAME    t
Library    Collections
Library    String
Library    DateTime
Library    Process
Library    Dialogs
Library    RPA.FTP

Suite Teardown    Run Keyword If Any Tests Failed    close

*** Tasks ***
Tally Purchase
    Start Tally
    Sales Entry driver

*** Keywords ***
start tally
    ${a}=    d.Open Application    C:\\Program Files\\TallyPrime\\tally.exe
    Set Suite Variable    ${app}    ${a}
    Wait For Element    ocr:(10003)    30
    d.Click    ocr:10002,81    double click
    # Wait For Element    ocr:Vouchers,70    10
    sleep    2   
    Press Keys    alt    f2
    Send Keys    keys={END}{BACK}0    send_enter=${True}
    Send Keys    keys={END}{BACK}3    send_enter=${True}
    Send Keys    keys=V

close
    Take Screenshot
    Close Application    ${app}

name format
    [Arguments]    ${name}
    ${name}=    Strip String    ${name}
    ${name}=    Convert To Lower Case    ${name}
    ${name}=    Convert To Title Case    ${name}
    ${name_list}=    Split String    ${name}
    ${name}=    Catenate    SEPARATOR={SPACE}    @{name_list}
    [Return]    ${name}

Company Validation
    [Arguments]    ${row}
    ${company}=    Get From Dictionary    ${row}    Billing Name
    ${company}=    name format    ${company}
    Send Keys    keys=${company}
    Set Local Variable    ${error}    0
    TRY
        Find Element    image:oops_error.png,60
    EXCEPT 
        Set Local Variable    ${error}    1
    END
    IF    ${error}==0    
        Press Keys    alt    c
        Send Keys    keys=${company}    send_enter=${True}
        Send Keys    keys={ENTER}
        Send Keys    keys=Sundry{SPACE}Creditors    send_enter=${True}
        Send Keys    keys=yes    send_enter=${True}
        Send Keys    keys={ENTER}{ENTER}{ENTER}    interval=0.3
        ${address}=    Get From Dictionary    ${row}    Billing Street
        ${address}=    name format    ${address}
        Send Keys    keys=${address}    send_enter=${True}
        Send Keys    keys={ENTER}
        ${state}=    Get From Dictionary    ${row}    Shipping Province Name
        Send Keys    keys=${state}    send_enter=${True}
        ${pincode}=    Get From Dictionary    ${row}    Shipping Zip 
        ${pincode}=    Fetch From Right    ${pincode}    '
        Send Keys    keys=${pincode}    send_enter=${True}
        Send Keys    keys={ENTER}
        Send Keys    keys=yes    send_enter=${True}
        Send Keys    keys={ENTER}{ENTER}{ENTER}
        ${email}=    Get From Dictionary    ${row}    Email
        Send Keys    keys=${email}    send_enter=${True}
        Send Keys    keys={ENTER}{ENTER}{ENTER}{ENTER}
        Send Keys    keys=Unregistered    send_enter=${True}
        Send Keys    keys={ENTER}
    END
    Send Keys    keys={ENTER}

Item Validation
    [Arguments]    ${row}    ${totalrate}
    ${item}=    Get From Dictionary    ${row}    Lineitem name
    ${item}=    name format    ${item}
    Send Keys    keys=${item}
    Set Local Variable    ${error}    0
    TRY
        Find Element    image:oops_error.png,60
    EXCEPT 
        Set Local Variable    ${error}    1
    END
    IF    ${error}==0    
        Press Keys    alt    c
        Send Keys    keys=${item}    send_enter=${True}
        Send Keys    keys={ENTER}{ENTER}
        Send Keys    keys=nos    send_enter=${True}
        Send Keys    keys={ENTER}{ENTER}{ENTER}{ENTER}    interval=0.3
        Send Keys    keys=yes    send_enter=${True}
        Send Keys    keys=${totalrate}    send_enter=${True}
    END
    Send Keys    keys={ENTER}{ENTER}    interval=0.3

Close Party Details
    Sleep    1
    FOR    ${i}    IN RANGE    0    10
        Set Local Variable    ${open}    0
        TRY
            Find Element    image:Party_Details.png,60
        EXCEPT
            Set Local Variable    ${open}    1
        END
        IF    ${open}==0
            Send Keys    keys={ENTER}
        ELSE
            Exit For Loop
        END
    END

Reopen Sales Entry
    Set Local Variable    ${flag}    0
    WHILE    ${flag}==0
        TRY
            Find Element    image:quit.png
            Send Keys    keys={ESC}v{F8}    interval=0.3
            Set Local Variable    ${flag}    1
        EXCEPT
            Send Keys    keys={ESC}
        END
    END

Sales Entry driver
    Send Keys    keys={F8}
    ${files}=    List Files In Directory    B:\\Shopify to Tally\\csv file
    FOR    ${file}    IN    @{files}         
        ${ext}=    Get File Extension    ${file}
        ${name}=    Get File Name    ${file}
        Continue For Loop If    '${ext}'!='.xlsx'
        Open Workbook    ${file}    data_only=${True}
        ${table_temp}=    Read Worksheet As Table    name=SALE    header=${True}    trim=${True}
        Set Global Variable    ${table}    ${table_temp}
        Close Workbook
        ${rows}    ${columns}    Get Table Dimensions    ${table}
        Set Global Variable    ${index}    0
        WHILE    ${index}<${rows}
            ${row}=    Get Table Row    ${table}    ${index}
            ${InvNum}=    Get From Dictionary    ${row}    Name
            IF    '${InvNum}'=='None' or '${InvNum}'==''     Exit For Loop
            TRY
                Wait Until Keyword Succeeds    3 times    1    Sales Entry worker    ${row}    ${name}    ${index}    ${rows}
            EXCEPT
                Open Workbook    error_log.xlsx
                Append Rows To Worksheet    content=&{row}    header=${True}
                Save Workbook
                Close Workbook
            END
            #Exit For Loop
        END
        #Exit For Loop
    END
    Close Application    ${app}

Sales Entry worker
    [Arguments]    ${row}    ${name}    ${index}    ${rows}
    ${date1}=    Get From Dictionary    ${row}    Created at
    TRY
        ${date}=    Convert Date    ${date1}    %d-%m-%Y
    EXCEPT
        ${date}=    Fetch From Left    ${date1}    ${SPACE}
        ${date}=    Convert Date    ${date1}    %d-%m-%Y
    END
    ${InvNum}=    Get From Dictionary    ${row}    Name
    ${InvNum}=    Fetch From Right    ${InvNum}    \#
    Send Keys    keys=${InvNum}    send_enter=${True}
    Send Keys    keys=${date}    send_enter=${True}
    Run Keyword    Company Validation    ${row}
    Run Keyword    Close Party Details
    ${tax}=    Get From Dictionary    ${row}    Tax 1 Name
    ${taxname}    ${taxrate}=    Split String    ${tax}
    ${taxrate}=    Fetch From Left    ${taxrate}    %
    ${taxrate}=    Convert To Integer    ${taxrate}
    IF    '${taxname}' == 'CGST' or '${taxname}' == 'SGST'
        ${totalrate}=    Evaluate    ${taxrate}*2
        ${salestype}=    Catenate    Sales    ${totalrate}
        ${salestype}=    name format    ${salestype}
    ELSE
        Set Local Variable    ${totalrate}    ${taxrate}
        ${salestype}=    Catenate    Interstate Sales    ${totalrate}
        ${salestype}=    name format    ${salestype}
    END
    sleep    0.5
    Send Keys    keys=${salestype}    send_enter=${True}
    sleep    0.5
    Run Keyword    Item Validation    ${row}    ${totalrate}
    sleep    0.5
    ${qty}=    Get From Dictionary    ${row}    Lineitem quantity
    Send Keys    keys=${qty}    send_enter=${True}
    Send Keys    keys={ENTER}
    ${amt}=    Get From Dictionary    ${row}    Lineitem price
    sleep    0.5
    Send Keys    keys=${amt}    send_enter=${True}
    sleep    0.5
    Send Keys    keys={ENTER}{ENTER}{ENTER}    interval=0.3
    ${index1}=    Evaluate    ${index}+1
    IF    ${index1}<${rows}
        ${row1}=    Get Table Row    ${table}    ${index1}
        ${InvNum1}=    Get From Dictionary    ${row1}    Name
        ${InvNum1}=    Fetch From Right    ${InvNum1}    \#
    END
    WHILE    '${InvNum}' == '${InvNum1}'
        Run Keyword    Item Validation    ${row1}    ${totalrate}
        ${qty}=    Get From Dictionary    ${row1}    Lineitem quantity
        Send Keys    keys=${qty}    send_enter=${True}
        Send Keys    keys={ENTER}
        ${amt}=    Get From Dictionary    ${row1}    Lineitem price
        sleep    0.5
        Send Keys    keys=${amt}    send_enter=${True}
        sleep    0.5
        Send Keys    keys={ENTER}{ENTER}{ENTER}    interval=0.3
        sleep    0.5
        ${index1}=    Evaluate    ${index1}+1
        IF    ${index1}<${rows}
            ${row1}=    Get Table Row    ${table}    ${index1}
            ${InvNum1}=    Get From Dictionary    ${row1}    Name
            ${InvNum1}=    Fetch From Right    ${InvNum1}    \#
        ELSE
            BREAK
        END
    END
    Set Global Variable    ${index}    ${index1}
    Send Keys    keys={ENTER}
    ${tax}=    name format    ${tax}
    IF    '${taxname}' == 'CGST' or '${taxname}' == 'SGST'
        Send Keys    keys=${tax}    send_enter=${True}
        Send Keys    keys={ENTER}{ENTER}    interval=0.3
        ${tax}=    Get From Dictionary    ${row}    Tax 2 Name
        ${tax}=    name format    ${tax}
        Send Keys    keys=${tax}    send_enter=${True}
        Send Keys    keys={ENTER}{ENTER}    interval=0.3
    ELSE
        Send Keys    keys=${tax}    send_enter=${True}
        Send Keys    keys={ENTER}{ENTER}{ENTER}    interval=0.3
    END
    Send Keys    keys=end    send_enter=${True}
    Send Keys    keys={ENTER}{ENTER}    interval=0.3
    TRY
        Wait For Element    image:accept.png,60
        Send Keys    keys={ENTER}
    EXCEPT
        TRY
            Send Keys    keys={ENTER}
            Wait For Element    image:accept.png,60
            Send Keys    keys={ENTER}
        EXCEPT
            Take Screenshot
            Reopen Sales Entry
            ${row_num}=    Evaluate    ${index}
            Fail    Accept not found. Last row was from ${name}'s sale sheet row number ${row_num}.
        END
    END
