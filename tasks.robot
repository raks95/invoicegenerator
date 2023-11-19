*** Settings ***
Documentation       Template robot main suite.

Library    RPA.Browser.Selenium    auto_close=${False}
Library    RPA.HTTP
Library    OperatingSystem
Library    RPA.Desktop
Library    RPA.FTP
Library    RPA.PDF
Library    RPA.FileSystem
Library    RPA.Word.Application
Library    RPA.Email.Exchange

*** Keywords ***
xpath=//ul[@class='ng-tns-c7-2 sub-menu ng-star-inserted']/li[@id='navbarMenuItem-Invoice']/a
    

*** Variables ***

#${PDF_URL}         https://www.openinvoice.com/docp/openInvoice/main/statusListAttachmentsAction?printFriendlyPreview=true&documentId=213036098&type=Invoice
${OUTPUT_DIR}      C:\\Users\\saira\\OneDrive\\Desktop\\DownloadedInvoice\\
*** Keywords ***

*** Tasks ***
Store WebPage Content
    Set Download Directory    ${OUTPUT_DIR}
    Open Chrome Browser    https://www.openinvoice.com/docp/public/OILogin.xhtml
    Maximize Browser Window
    Input Text    //input[@id='j_username']    salestax@saltapllc.com
    Input Password    //input[@name = 'j_password']    ToBotOrNotR3
    Click Button    //button[@id = 'loginBtn']
    Sleep    10
    Mouse Over    navbarMenuItem-Invoice
    Click Element When Visible    navbarSubItem-InvoiceSearch
    Sleep    5

     # Find all 'Cancel' buttons using the specified XPath
    ${cancel_buttons}    Get WebElements    //div[@class='text-core']/div[@class='text-wrap']//a[@class='text-remove']

    # Loop through each cancel button and click on it
    FOR    ${button}    IN    @{cancel_buttons}
        Click Element    ${button}
    END

    Input Text When Element Is Visible    documentNumber    523812
    Input Text When Element Is Visible    supplierNumber    A00083
    Click Element When Visible    DetailedSearch1_label
    #Click Element When Visible    css:.centerAlign a

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
    #Rename    viewAttachment , Invoice
    Click Element When Visible    sign-out
    Click Element When Visible    revit_form_Button_0_label
    Sleep    20
    OperatingSystem.Move File    ${OUTPUT_DIR}\\viewAttachment.pdf  ${OUTPUT_DIR}\\523812.pdf
    Log    Task Completed