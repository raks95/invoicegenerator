*** Settings ***
Documentation      Invoice Downloader
# InvoiceQuery Application v1.0
# Copyright (C) 2023 R3 Tax Solutions, Inc.
 

Library    RPA.Browser.Selenium    auto_close=${False}
Library    RPA.HTTP
Library    OperatingSystem
Library    RPA.Desktop
Library    RPA.FTP
Library    RPA.Robocorp.WorkItems
Library    RPA.RobotLogListener

*** Keywords ***
xpath=//ul[@class='ng-tns-c7-2 sub-menu ng-star-inserted']/li[@id='navbarMenuItem-Invoice']/a
    

*** Variables ***
${Cancel_path}   //div[@class='text-core']/div[@class='text-wrap']//a[@class='text-remove']
${OUTPUT_DIR}      C:\\Users\\saira\\OneDrive\\Desktop\\DownloadedInvoice
${URL}      https://www.openinvoice.com/docp/public/OILogin.xhtml


*** Tasks ***
Login
    Set Download Directory    ${OUTPUT_DIR}
    Open Chrome Browser     ${URL}
    Maximize Browser Window
    Input Text    //input[@id='j_username']    salestax@saltapllc.com
    Input Password    //input[@name = 'j_password']    ToBotOrNotR3
    Click Button    //button[@id = 'loginBtn']
    Sleep    10
    Mouse Over    navbarMenuItem-Invoice
    Click Element When Visible    navbarSubItem-InvoiceSearch
    Sleep    5
Removing Data
     # Find all 'Cancel' buttons using the specified XPath
    ${cancel_buttons}    Get WebElements    ${Cancel_path}

    # Loop through each cancel button and click on it
    FOR    ${button}    IN    @{cancel_buttons}
        Click Element    ${button}
    END

Entering Invoice Details
    Input Text When Element Is Visible    documentNumber     523812
    Input Text When Element Is Visible    supplierNumber     A00083
       
    Click Element When Visible    DetailedSearch1_label
    #Click Element When Visible    css:.centerAlign a

Downloading Invoice
    Click Element When Visible    xpath=//a[contains(text(), '523812')]
    Sleep    5
    Wait Until Element Contains    xpath=//a[text()='523812']    523812
    ${url}    Get Element Attribute    xpath=//a[text()='523812']    href
    Execute Javascript    window.open('', '_blank');
    ${handles} =    Get Window Handles
    Switch Window    ${handles}[1]
    Go To    ${URL}    # Load a URL in the new tab
    Close Window    
    Switch Window    ${handles}[0]
    Sleep    10
    Click Element When Visible    sign-out
    Click Element When Visible    revit_form_Button_0_label
    Sleep    20
    
Renaming Invoice
    OperatingSystem.Move File    ${OUTPUT_DIR}\\viewAttachment.pdf  ${OUTPUT_DIR}\\523812.pdf
    Log    Task Completed
