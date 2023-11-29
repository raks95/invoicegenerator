*** Settings ***
Documentation      Invoice Downloader
Library    RPA.Browser.Selenium    auto_close=${False}
Library    RPA.HTTP
Library    OperatingSystem
Library    RPA.PDF
Library    RPA.Excel.Files
Library    Collections
Library    RPA.Tables
Library    RPA.FTP
Library    XML
Library    String
Library    RPA.Windows
Library    RPA.Excel.Application
Library    RPA.PDF
Library    RPA.Excel.Files
Library    Collections
Library    RPA.Tables
Library    RPA.FTP
Library    XML
Library    String
Library    RPA.Windows
Library    RPA.Excel.Application
Library    RPA.Desktop

*** Variables ***
${OUTPUT_DIR}    C:\\Users\\saira\\Desktop\\DownloadedInvoice
${Renamed_Invoice}  C:\\Users\\saira\\Desktop\\RenamedInvoices
${Excel_path}   C:\\Users\\saira\\Desktop\\InvoicePullData.xlsx

*** Variables ***
${URL}    https://www.openinvoice.com/docp/public/OILogin.xhtml
${Cancel_path}   //div[@class='text-core']/div[@class='text-wrap']//a[@class='text-remove']
${my_integer}    1
${current_row_index}    1  # Initialize with the starting row index
${filename}    ""
${EXCEPTION}    ""

*** Variables ***
${1000}    Success, Single
${1100}    Single Invoice Found - No Download
${1001}    Success, Multi, Attachment 1 found
${1103}    Multiple Invoices Found - No Download for Either
${2001}    Record Not Found
${1000_Long}    Successful retrieval, only one attachment present
${1100_Long}    Invoice was found, but the download took longer than the maximum time allowed (400 seconds)
${1001_Long}    Successful retrieval, but there were multiple attachments and one of them contained a search string hit for Attachment 1
${1103_Long}    Invoice was found, but the download for both files took longer than the maximum time allowed each (400 seconds)
${2001_Long}    Failed retrieval, based on the invoice not being found in the search



*** Tasks ***
Logging into Application
    Login

Removing Data From Cancel buttons
    Removing Data

Fetching The Excel records, Downloading Invoice and Updating Excel
    Fetch Excel Records  



*** Keywords ***
Login
    Set Download Directory    ${OUTPUT_DIR} 
    Open Chrome Browser     ${URL}
    Maximize Browser Window
    Input Text    //input[@id='j_username']    salestax@saltapllc.com
    Input Password    //input[@name='j_password']    ToBotOrNotR3
    Click Button    //button[@id='loginBtn']
    Wait Until Page Contains Element    navbarMenuItem-Invoice    50
    Mouse Over    navbarMenuItem-Invoice
    Wait Until Page Contains Element    navbarSubItem-InvoiceSearch    50
    Click Element When Visible    navbarSubItem-InvoiceSearch
    Wait Until Page Contains Element    ${Cancel_path}     50


Removing Data
    ${cancel_buttons}    Get WebElements    ${Cancel_path}
    FOR    ${button}    IN    @{cancel_buttons}
        Click Element    ${button}
    END

    
Fetch Excel Records
        #TODO: close all the open excel
        RPA.Excel.Files.Open Workbook    ${Excel_path}
        ${InvoiceDetails} =    Read Worksheet As Table    header=${True}     
    FOR      
    ...  ${invoiceNumbers}    IN    @{InvoiceDetails}
        ${current_row_index}     Evaluate     ${current_row_index} + 1
        ${result_value}    Set Variable    ${invoiceNumbers}[Result]
        #Log    Current Row Index: ${current_row_index}, Result Value: ${result_value}
        # Check if the "Result" column is empty for the current row
        IF     ${result_value} == $None
            Downloading Invoice    ${invoiceNumbers}    ${current_row_index}
            Wait Until Page Contains Element    ${Cancel_path}  50
            Run Keyword    Removing Data            
        END        
    END
    Wait Until Page Contains Element    sign-out
    Click Element When Visible    sign-out
    Wait Until Page Contains Element    revit_form_Button_0_label
    Click Element When Visible    revit_form_Button_0_label
    

Downloading Invoice
    [Arguments]    ${invoiceNumbers}    ${row_number}
        Input Text When Element Is Visible    documentNumber    ${invoiceNumbers}[Inv No]
        Input Text When Element Is Visible    supplierNumber    ${invoiceNumbers}[Client Vendor No]
    # Check if the element visible is enough - TODO (**)
        Wait Until Page Contains Element    DetailedSearch1_label   50
        Click Element    DetailedSearch1_label
        Wait Until Page Contains     Invoice Search Results    50
        ${elementExists} =   Run Keyword    Is Invoice Present  ${invoiceNumbers}
    
        IF    ${elementExists}
            ${Single_Invoice} =   Run Keyword    single Invoice
            IF     ${Single_Invoice}
            Click Element When Visible    css:.centerAlign a     
            Sleep    5
            ${handles} =    Get Window Handles
            Switch Window    ${handles}[1]
                    TRY
                        OperatingSystem.Wait Until Created    ${OUTPUT_DIR}\\statusListAttachmentsAction.pdf     400
                        ${filename} =   Set Variable   ${invoiceNumbers}[Vendor Name]
                        ${filename}   Evaluate    "${filename}".rstrip()
                        ${filename} =   Set Variable   ${filename}-${invoiceNumbers}[Inv No](1).pdf
                        OperatingSystem.Move File    ${OUTPUT_DIR}\\statusListAttachmentsAction.pdf    ${Renamed_Invoice}\\${filename}
                        Run Keyword    Update Excel Status     ${invoiceNumbers}    1000    ${row_number}
                        Run Keyword    Write Filename To Excel    ${invoiceNumbers}   ${filename}     ${row_number}
                        Run Keyword    Write Short_Desc To Excel    ${invoiceNumbers}    ${1000}    ${row_Number}
                        Run Keyword    Write Long_Desc To Excel     ${invoiceNumbers}    ${1000_Long}    ${row_Number}
                        Switch Window    ${handles}[0]
                        Wait Until Page Contains Element    Invoice_navItem_label    60
                        Mouse Over     Invoice_navItem_label
                        Wait Until Page Contains Element    Invoice_ttd_Invoice_InvoiceSearch     60
                        Click Element When Visible        Invoice_ttd_Invoice_InvoiceSearch
                        
                    EXCEPT  
                        Close Browser
                        Run Keyword    Login    
                        Run Keyword    Update Excel Status     ${invoiceNumbers}    1100    ${row_number}
                        Run Keyword    Write Short_Desc To Excel    ${invoiceNumbers}    ${1100}    ${row_Number}
                        Run Keyword    Write Long_Desc To Excel     ${invoiceNumbers}    ${1100_Long}    ${row_Number}
                        OperatingSystem.Remove Directory    ${OUTPUT_DIR}     overwrite=True  
                    END
            # Update status for single invoice - Modify this according to your needs
            ELSE
            # Get all the PDF in the page which contains PDF attachment
            # Get all the elements with the text    
            # Loop through all the elements and download
                    TRY
                        Click Element When Visible    css:.centerAlign a
                        Sleep    5
                        ${handles} =  Get Window Handles
                        Switch Window    ${handles}[1]
                        Click Element When Visible    css:.journalComment a
                        # Update status for multiple invoices - Modify this according to your needs
                        Run Keyword    Update Excel Status     ${invoiceNumbers}   1001     ${row_number}
                        Run Keyword    Write Short_Desc To Excel    ${invoiceNumbers}    ${1001}    ${row_Number}
                        Run Keyword    Write Long_Desc To Excel     ${invoiceNumbers}    ${1001_Long}    ${row_Number}
                        RPA.Browser.Selenium.Close Window
                        Switch Window    ${handles}[0]
                        Wait Until Page Contains Element    Invoice_navItem_label    60
                        Mouse Over     Invoice_navItem_label
                        Wait Until Page Contains Element    Invoice_ttd_Invoice_InvoiceSearch     60
                        Click Element When Visible        Invoice_ttd_Invoice_InvoiceSearch
                    EXCEPT    
                        Run Keyword    Update Excel Status     ${invoiceNumbers}    1103    ${row_number}
                        Run Keyword    Write Short_Desc To Excel    ${invoiceNumbers}   ${1103}     ${row_Number}
                        Run Keyword    Write Long_Desc To Excel     ${invoiceNumbers}    ${1103_Long}    ${row_Number}
                    #logout
                    #close the
                    END
            END
            
        ELSE
            Run Keyword    Update Excel Status     ${invoiceNumbers}   2001    ${row_number}
            Run Keyword    Write Short_Desc To Excel    ${invoiceNumbers}    ${2001}    ${row_Number}
            Run Keyword    Write Long_Desc To Excel     ${invoiceNumbers}    ${2001_Long}    ${row_Number}
        # Update status when invoice is not found - Modify this according to your needs
            Wait Until Page Contains Element    Invoice_navItem_label    60
            Mouse Over     Invoice_navItem_label
            Wait Until Page Contains Element    Invoice_ttd_Invoice_InvoiceSearch     60
            Click Element When Visible     Invoice_ttd_Invoice_InvoiceSearch
            Sleep    10
        END 
    
single Invoice
    ${SingleInvoiceText} =  RPA.Browser.Selenium.Get Text    css:.centerAlign a
    ${IsSingleInvoice} =  Evaluate  "${SingleInvoiceText}" == "1"
    Log  One Invoice: ${IsSingleInvoice}
    [Return]  ${IsSingleInvoice}

Is Invoice Present
    [Arguments]    ${invoiceNumbers}
    ${elementExists} =    Does Page Contain Element      //a[contains(text(), '${invoiceNumbers}[Inv No]')]
    Log    Element exists: ${elementExists}
    IF    ${elementExists}
       RETURN     ${True}
    ELSE
        RETURN    ${False}
    END

Update Excel Status
    [Arguments]    ${invoiceNumbers}    ${status}    ${row_Number}
    RPA.Excel.Files.Open Workbook    ${Excel_path}
    Set Cell Value    ${row_Number}    7    ${status}
    Save Workbook

Write Filename To Excel
    [Arguments]    ${invoiceNumbers}    ${status}    ${row_Number}
    RPA.Excel.Files.Open Workbook    ${Excel_path}
    Set Cell Value    ${row_Number}    5    ${status}
    Save Workbook

Write Short_Desc To Excel
    [Arguments]    ${invoiceNumbers}    ${Short_Desc}    ${row_Number}
    RPA.Excel.Files.Open Workbook    ${Excel_path}
    Set Cell Value    ${row_Number}    8     ${Short_Desc}
    Save Workbook

Write Long_Desc To Excel
    [Arguments]    ${invoiceNumbers}    ${Long_Desc}    ${row_Number}
    RPA.Excel.Files.Open Workbook    ${Excel_path}
    Set Cell Value    ${row_Number}    9    ${Long_Desc}
    Save Workbook


