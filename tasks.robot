
*** Settings ***

Documentation      Invoice Downloader

Library    RPA.Browser.Selenium    auto_close=${False}

Library    OperatingSystem

Library    RPA.PDF

Library    RPA.Excel.Files

Library    RPA.Desktop
Library    String

 

*** Variables ***
${OUTPUT_DIR}    C:\\Users\\Yeswanth\\Downloads\\DownloadedInvoice
${Renamed_Invoice}  C:\\Users\\Yeswanth\\Desktop\\RenamedInvoices1
${Excel_path}    C:\\Users\\Yeswanth\\Downloads\\InvoicePullData.xlsx
 

*** Variables ***

${URL}    https://www.openinvoice.com/docp/public/OILogin.xhtml
${my_integer}    1
${donwloadCount}    0
${current_row_index}    1  # Initialize with the starting row index
${filename}    ""
${EXCEPTION}    ""
${DownloadPath}    xpath://a[contains(@href, '/docp/openInvoice/main/viewAttachment.pdf?attachmentId=')]
${Cancel_path}    //div[@class='text-core']/div[@class='text-wrap']//a[@class='text-remove']
@{elements_invoice}    Invoice    inv

 

*** Variables ***

${1000}    Success, Single
${1100}    Single Invoice Found - No Download
${1001}    Success, Multi, Attachment 1 found
${1103}    Multiple Invoices Found – Download Failure
${1104}    Multiple Invoices Found - Bothe the Invoices are downloaded
${1900}    Success, Single – No PDF File Found
${1901}    Success, Multi – No PDF File Found
${2001}    Record Not Found
${1900_Long}    Invoice was found but there were no PDF attachments associated with it – download was skipped
${1901_Long}    Invoice was found and there were attachments associated with it – but none was a PDF so download was skipped
${1000_Long}    Successful retrieval, only one attachment present
${1100_Long}    Invoice was found, but the download took longer than the maximum time allowed (400 seconds)
${1001_Long}    Successful retrieval, but there were multiple attachments and one of them contained a search string hit for Attachment 1
${1103_Long}    Invoice was found, but the download for at least one file took longer than the maximum time allowed each (400 seconds)
${1104_Long}    Multiple Invoices are found two invoices contained a search string & both are downloaded. 
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
    Input Text    //input[@id='j_username']    hwillis@saltapllc.com
    Input Password    //input[@name='j_password']    Blackbeardguest8
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
            IF  ${Single_Invoice}
                    TRY
                    ${anchor_element}  Get WebElements    xpath=//table[@id='documentListTableId']//a[contains(@href,'invoiceInput?documentId=')]
                    Click Element      ${anchor_element}[0]
                    ${has_open} =  Run Keyword And Return Status  Wait Until Page Contains Element  xpath=//a[contains(@onclick,'openTicketAttachmentWindow')]  5
                    IF  not ${has_open}
                        Wait Until Page Contains Element    xpath=//td[contains(@class,'journalComment')]//a    10
                    END
                    ${html} =    Get Source
                    ${matches}=   Get Regexp Matches    ${html}    openTicketAttachmentWindow\\s*\\(\\s*'([^']+)'    1
                    ${length}       Get Length    ${matches}
                    ${all_links} =    Get WebElements    xpath=//td[contains(@class,'journalComment')]//a
                    ${all_count} =    Get Length    ${all_links}
                    IF   ${all_count} >= 1
                        ${pdf_links} =    Get WebElements    xpath=//td[contains(@class,'journalComment')]//a[contains(translate(normalize-space(.),'PDF','pdf'),'.pdf') or contains(translate(@href,'PDF','pdf'),'.pdf') or contains(translate(@onclick,'PDF','pdf'),'.pdf')]
                        ${pdf_count}       Get Length    ${pdf_links}
                        IF   ${pdf_count} == 0
                            Run Keyword    Update Excel Status     ${invoiceNumbers}   1900    ${row_number}
                            Run Keyword    Write Short_Desc To Excel    ${invoiceNumbers}    ${1900}    ${row_Number}
                            Run Keyword    Write Long_Desc To Excel     ${invoiceNumbers}    ${1900_Long}    ${row_Number}
                        ELSE
                            ${pdf_link}=    Set Variable    ${pdf_links}[0]
                            ${pdf_href}=    Get Element Attribute    ${pdf_link}    href
                            ${pdf_name}=    Get Text    ${pdf_link}
                            ${pdf_name}=    Evaluate    "${pdf_name}".strip()
                            IF    "${pdf_name}" == ""
                                ${pdf_name}=    Set Variable    viewAttachment.pdf
                            END
                            ${full_url}=    Evaluate    "${pdf_href}" if "${pdf_href}".startswith("http") else "https://www.openinvoice.com${pdf_href}"
                            Execute Javascript    window.open('${full_url}', '_blank')
                            Wait For Download Complete    ${OUTPUT_DIR}\\viewAttachment.pdf     400
                            ${filename} =   Set Variable    ${invoiceNumbers}[Vendor Name]
                            ${filename}   Evaluate    "${filename}".rstrip()                  
                            ${filename} =   Set Variable   ${filename}-${invoiceNumbers}[Inv No](1).pdf
                            OperatingSystem.Move File      ${OUTPUT_DIR}\\viewAttachment.pdf      ${Renamed_Invoice}\\${filename}
                            Run Keyword    Update Excel Status     ${invoiceNumbers}    1000    ${row_number}
                            Run Keyword    Write Filename To Excel    ${invoiceNumbers}   ${filename}     ${row_number}
                            Run Keyword    Write Short_Desc To Excel    ${invoiceNumbers}    ${1000}    ${row_Number}
                            Run Keyword    Write Long_Desc To Excel     ${invoiceNumbers}    ${1000_Long}    ${row_Number}
                        END
                    ELSE
                        IF   ${length} == 0
                            Run Keyword    Update Excel Status     ${invoiceNumbers}   1900    ${row_number}
                            Run Keyword    Write Short_Desc To Excel    ${invoiceNumbers}    ${1900}    ${row_Number}
                            Run Keyword    Write Long_Desc To Excel     ${invoiceNumbers}    ${1900_Long}    ${row_Number}
                        ELSE
                            ${attachment_url}=    Set Variable    ${matches}[0]
                            ${attachment_id}=     Get Regexp Matches    ${attachment_url}    /attachments/(\\d+)    1
                            ${attachment_id}=     Set Variable    ${attachment_id}[0]
                            Log    Attachment ID: ${attachment_id}
                            ${full_url}=    Set Variable    https://www.openinvoice.com${attachment_url}
                            Execute Javascript    window.open('${full_url}', '_blank')
                            Wait For Download Complete    ${OUTPUT_DIR}\\${attachment_id}.pdf     400
                            ${filename} =   Set Variable    ${invoiceNumbers}[Vendor Name]
                            ${filename}   Evaluate    "${filename}".rstrip()                  
                            ${filename} =   Set Variable   ${filename}-${invoiceNumbers}[Inv No](1).pdf
                            OperatingSystem.Move File      ${OUTPUT_DIR}\\${attachment_id}.pdf      ${Renamed_Invoice}\\${filename}
                            Run Keyword    Update Excel Status     ${invoiceNumbers}    1000    ${row_number}
                            Run Keyword    Write Filename To Excel    ${invoiceNumbers}   ${filename}     ${row_number}
                            Run Keyword    Write Short_Desc To Excel    ${invoiceNumbers}    ${1000}    ${row_Number}
                            Run Keyword    Write Long_Desc To Excel     ${invoiceNumbers}    ${1000_Long}    ${row_Number}
                        END
                    END
                        Wait Until Page Contains Element    Invoice_navItem_label    60
                        Mouse Over     Invoice_navItem_label
                        Wait Until Page Contains Element    Invoice_ttd_ManageInvoices_InvoiceSearch    50
                        Click Element When Visible    Invoice_ttd_ManageInvoices_InvoiceSearch
                    EXCEPT  (msg=${EXCEPTION})
                        Close Browser
                        Run Keyword    Login    
                        Run Keyword    Update Excel Status     ${invoiceNumbers}    1100    ${row_number}
                        Run Keyword    Write Short_Desc To Excel    ${invoiceNumbers}    ${1100}    ${row_Number}
                        Run Keyword    Write Long_Desc To Excel     ${invoiceNumbers}    ${1100_Long}    ${row_Number}
                        OperatingSystem.Remove Directory    ${OUTPUT_DIR}     overwrite=True  
                    END

            # Update status for single invoice - Modify this according to your needs
            ELSE
                TRY
                    ${anchor_element}  Get WebElements  xpath=//a[contains(@onclick, 'return buttonClickForSubmit(event)')]
                    ${href_value}  Get Element Attribute  ${anchor_element}  href
                    Execute Javascript    window.open('', '_blank');
                    ${handles} =    Get Window Handles
                    Switch Window    ${handles}[1]
                    Go To    ${href_value}    # Load a URL in the new tab
                    #${elements}   Get WebElements    .journalComment a
                    ${element_s}    Get WebElements   ${DownloadPath}
                    ${length}    Get Length    ${element_s}
                    #If condition check if array is null writ excel and close.[]
                    #Create a new int with 0  
                    #${filtered_links}=     Create List
                        IF   ${length} == 0
                            Run Keyword    Update Excel Status     ${invoiceNumbers}   1901    ${row_number}
                            Run Keyword    Write Short_Desc To Excel    ${invoiceNumbers}    ${1901}    ${row_Number}
                            Run Keyword    Write Long_Desc To Excel     ${invoiceNumbers}    ${1901_Long}    ${row_Number}
                        END
                        IF   ${length} == 1
                                ${donwloadCount}   Evaluate     ${donwloadCount} + 2
                                Click Element    ${element_s}
                                ${filename} =   Set Variable    ${invoiceNumbers}[Vendor Name]
                                ${filename}   Evaluate    "${filename}".rstrip()
                                Wait For Download Complete    ${OUTPUT_DIR}\\viewAttachment.pdf     400
                                ${filename} =   Set Variable   ${filename}-${invoiceNumbers}[Inv No](1).pdf
                                OperatingSystem.Move File      ${OUTPUT_DIR}\\viewAttachment.pdf      ${Renamed_Invoice}\\${filename}
                                IF  ${donwloadCount} == 2
                                    Run Keyword    Update Excel Status     ${invoiceNumbers}    1000    ${row_number}
                                    Run Keyword    Write Filename To Excel    ${invoiceNumbers}   ${filename}     ${row_number}
                                    Run Keyword    Write Short_Desc To Excel    ${invoiceNumbers}    ${1000}    ${row_Number}
                                    Run Keyword    Write Long_Desc To Excel     ${invoiceNumbers}    ${1000_Long}    ${row_Number}
                                END
                        END
                        FOR    ${index}    ${element}    IN ENUMERATE    @{element_s}
                                IF  ${donwloadCount} == 2
                                    BREAK
                                END
                               
                                IF   ${length} == 2
                                    ${donwloadCount}   Evaluate     ${donwloadCount} + 1
                                    Click Element    ${element}
                                    ${filename} =   Set Variable    ${invoiceNumbers}[Vendor Name]
                                    ${filename}   Evaluate    "${filename}".rstrip()
                                    Wait For Download Complete    ${OUTPUT_DIR}\\viewAttachment.pdf     400
                                    ${filename} =   Set Variable   ${filename}-${invoiceNumbers}[Inv No](${donwloadCount}).pdf
                                    OperatingSystem.Move File      ${OUTPUT_DIR}\\viewAttachment.pdf      ${Renamed_Invoice}\\${filename}
                                            IF  ${donwloadCount} == 1
                                                Run Keyword    Write Filename To Excel    ${invoiceNumbers}   ${filename}     ${row_number}
                                                Run Keyword    Update Excel Status     ${invoiceNumbers}   1001     ${row_number}
                                                Run Keyword    Write Short_Desc To Excel    ${invoiceNumbers}    ${1001}    ${row_Number}
                                                Run Keyword    Write Long_Desc To Excel     ${invoiceNumbers}    ${1001_Long}    ${row_Number}
                                            ELSE IF  ${donwloadCount} == 2
                                                Run Keyword    Write Second Filename To Excel    ${invoiceNumbers}   ${filename}     ${row_number}
                                                Run Keyword    Update Excel Status     ${invoiceNumbers}   1004     ${row_number}
                                                Run Keyword    Write Short_Desc To Excel    ${invoiceNumbers}    ${1104}    ${row_Number}
                                                Run Keyword    Write Long_Desc To Excel     ${invoiceNumbers}    ${1104_Long}    ${row_Number}
                                           END
                                END
                                    IF   ${length} >= 3
                                       
                                        FOR    ${element}    IN    @{elements}
                                            ${element_text} =    Get Text    ${element}
                                            ${contains_invoice} =    Check If Partial String Contains Invoice    ${element_text}
                                            IF  ${contains_invoice}
                                                ${donwloadCount}   Evaluate     ${donwloadCount} + 1
                                                Click Element    ${element}
                                                ${filename} =   Set Variable    ${invoiceNumbers}[Vendor Name]
                                                ${filename}   Evaluate    "${filename}".rstrip()
                                                Wait For Download Complete    ${OUTPUT_DIR}\\viewAttachment.pdf     400
                                                ${filename} =   Set Variable   ${filename}-${invoiceNumbers}[Inv No](${donwloadCount}).pdf
                                                OperatingSystem.Move File      ${OUTPUT_DIR}\\viewAttachment.pdf      ${Renamed_Invoice}\\${filename}
                                                END
                                            IF  ${donwloadCount} == 1
                                                Run Keyword    Write Filename To Excel    ${invoiceNumbers}   ${filename}     ${row_number}
                                                Run Keyword    Update Excel Status     ${invoiceNumbers}   1001     ${row_number}
                                                Run Keyword    Write Short_Desc To Excel    ${invoiceNumbers}    ${1001}    ${row_Number}
                                                Run Keyword    Write Long_Desc To Excel     ${invoiceNumbers}    ${1001_Long}    ${row_Number}
                                            ELSE IF  ${donwloadCount} == 2
                                                Run Keyword    Write Second Filename To Excel    ${invoiceNumbers}   ${filename}     ${row_number}
                                                Run Keyword    Update Excel Status     ${invoiceNumbers}   1004     ${row_number}
                                                Run Keyword    Write Short_Desc To Excel    ${invoiceNumbers}    ${1104}    ${row_Number}
                                                Run Keyword    Write Long_Desc To Excel     ${invoiceNumbers}    ${1104_Long}    ${row_Number}
                                           END
                                            IF  ${donwloadCount} == 2
                                                BREAK
                                            END
                                        END
                                    IF  ${donwloadCount} == 0
                                        FOR    ${element}    IN    Slice List    ${elements}    0    1
                                            ${donwloadCount}   Evaluate     ${donwloadCount} + 1
                                            Click Element    ${element}
                                            ${filename} =   Set Variable    ${invoiceNumbers}[Vendor Name]
                                            ${filename}   Evaluate    "${filename}".rstrip()
                                            Wait For Download Complete    ${OUTPUT_DIR}\\viewAttachment.pdf     400
                                            ${filename} =   Set Variable   ${filename}-${invoiceNumbers}[Inv No](${donwloadCount}).pdf
                                            OperatingSystem.Move File      ${OUTPUT_DIR}\\viewAttachment.pdf      ${Renamed_Invoice}\\${filename}
                                            IF  ${donwloadCount} == 1
                                                Run Keyword    Write Filename To Excel    ${invoiceNumbers}   ${filename}     ${row_number}
                                                Run Keyword    Update Excel Status     ${invoiceNumbers}   1001     ${row_number}
                                                Run Keyword    Write Short_Desc To Excel    ${invoiceNumbers}    ${1001}    ${row_Number}
                                                Run Keyword    Write Long_Desc To Excel     ${invoiceNumbers}    ${1001_Long}    ${row_Number}
                                            ELSE IF  ${donwloadCount} == 2
                                                Run Keyword    Write Second Filename To Excel    ${invoiceNumbers}   ${filename}     ${row_number}
                                                Run Keyword    Update Excel Status     ${invoiceNumbers}   1004     ${row_number}
                                                Run Keyword    Write Short_Desc To Excel    ${invoiceNumbers}    ${1104}    ${row_Number}
                                                Run Keyword    Write Long_Desc To Excel     ${invoiceNumbers}    ${1104_Long}    ${row_Number}
                                           END
                                        END
                                    ELSE IF  ${donwloadCount} == 1
                                        ${donwloadCount}   Evaluate     ${donwloadCount} + 1
                                        Click Element    ${elements}[1]
                                        ${filename} =   Set Variable    ${invoiceNumbers}[Vendor Name]
                                        ${filename}   Evaluate    "${filename}".rstrip()
                                        Wait For Download Complete    ${OUTPUT_DIR}\\viewAttachment.pdf     400
                                        ${filename} =   Set Variable   ${filename}-${invoiceNumbers}[Inv No](${donwloadCount}).pdf
                                        OperatingSystem.Move File      ${OUTPUT_DIR}\\viewAttachment.pdf      ${Renamed_Invoice}\\${filename}
                                            IF  ${donwloadCount} == 2
                                                Run Keyword    Write Second Filename To Excel    ${invoiceNumbers}   ${filename}     ${row_number}
                                                Run Keyword    Update Excel Status     ${invoiceNumbers}   1004     ${row_number}
                                                Run Keyword    Write Short_Desc To Excel    ${invoiceNumbers}    ${1104}    ${row_Number}
                                                Run Keyword    Write Long_Desc To Excel     ${invoiceNumbers}    ${1104_Long}    ${row_Number}
                                           END
                                        END
                                    END
                                       
                           
                        END
                        # Update status for multiple invoices - Modify this according to your needs
                        Close Window    
                        Switch Window    ${handles}[0]
                        Wait Until Page Contains Element    Invoice_navItem_label    60
                        Mouse Over     Invoice_navItem_label
                        Wait Until Page Contains Element     Invoice_ttd_ManageInvoices_InvoiceSearch    50
                        Click Element When Visible     Invoice_ttd_ManageInvoices_InvoiceSearch
                       
                EXCEPT  
                    Run Keyword    Update Excel Status     ${invoiceNumbers}    1103    ${row_number}
                    Run Keyword    Write Short_Desc To Excel    ${invoiceNumbers}   ${1103}     ${row_Number}
                    Run Keyword    Write Long_Desc To Excel     ${invoiceNumbers}    ${1103_Long}    ${row_Number}
                    Close Browser
                    Run Keyword    Login 
                    OperatingSystem.Remove Directory    ${OUTPUT_DIR}     overwrite=True
                END
            END
        ELSE
            Run Keyword    Update Excel Status     ${invoiceNumbers}   2001    ${row_number}
            Run Keyword    Write Short_Desc To Excel    ${invoiceNumbers}    ${2001}    ${row_Number}
            Run Keyword    Write Long_Desc To Excel     ${invoiceNumbers}    ${2001_Long}    ${row_Number}
        # Update status when invoice is not found - Modify this according to your needs
            Wait Until Page Contains Element    Invoice_navItem_label    60
            Mouse Over     Invoice_navItem_label
            Wait Until Page Contains Element     Invoice_ttd_ManageInvoices_InvoiceSearch    50
            Click Element When Visible     Invoice_ttd_ManageInvoices_InvoiceSearch
            Sleep    10
        END

single Invoice

    ${SingleInvoiceText} =  Get Element Count   xpath=//table[@id='documentListTableId']//a[contains(@href,'invoiceInput?documentId=')]
    ${attachment_cell_xpath} =  Set Variable  xpath=//table[@id='documentListTableId']//tbody/tr[not(@id='orig_tableHeaderId')][1]/td[count(//table[@id='documentListTableId']//tr[@id='orig_tableHeaderId']/td[@id='HEADER31']/preceding-sibling::td)+1]
    ${attachment_text} =  Get Text  ${attachment_cell_xpath}
    ${attachment_count} =  Evaluate  1 if "${attachment_text}".strip() == "-" else int("${attachment_text}".strip() or 0)
    ${IsSingleInvoice} =  Evaluate  "${SingleInvoiceText}" == "1" and int(${attachment_count}) <= 1
    Log  One Invoice (<=1 attachment): ${IsSingleInvoice}
    [Return]  ${IsSingleInvoice}

Is Invoice Present

    [Arguments]    ${invoiceNumbers}
    ${inv_no_lower} =    Evaluate    "${invoiceNumbers}[Inv No]".lower()
    ${elementExists} =    Does Page Contain Element      //a[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), '${inv_no_lower}')]
    Log    Element exists: ${elementExists}
    IF    ${elementExists}
       RETURN     ${True}
    ELSE
        RETURN    ${False}
    END

Write Filename To Excel

    [Arguments]    ${invoiceNumbers}    ${status}    ${row_Number}
    Open Workbook    ${Excel_path}
    Set Cell Value    ${row_Number}    5    ${status}
    Save Workbook

Write Second Filename To Excel

    [Arguments]    ${invoiceNumbers}    ${status}    ${row_Number}
    Open Workbook    ${Excel_path}
    Set Cell Value    ${row_Number}    6    ${status}
    Save Workbook

Update Excel Status

    [Arguments]    ${invoiceNumbers}    ${status}    ${row_Number}
    Open Workbook    ${Excel_path}
    Set Cell Value    ${row_Number}    7    ${status}
    Save Workbook

Write Short_Desc To Excel

    [Arguments]    ${invoiceNumbers}    ${Short_Desc}    ${row_Number}
    Open Workbook    ${Excel_path}
    Set Cell Value    ${row_Number}    8     ${Short_Desc}
    Save Workbook

Write Long_Desc To Excel

    [Arguments]    ${invoiceNumbers}    ${Long_Desc}    ${row_Number}
    Open Workbook    ${Excel_path}
    Set Cell Value    ${row_Number}    9    ${Long_Desc}
    Save Workbook

Check If Partial String Contains Invoice
    [Arguments]    @{strings} 
    FOR    ${input_string}    IN    @{strings}
        ${contains_invoice}=    Run Keyword And Return Status    Should Contain    ${input_string.lower()}    invoice
        ${contains_inv}=    Run Keyword And Return Status    Should Contain    ${input_string.lower()}    inv
       
        IF    (${contains_invoice} or ${contains_inv})
            Return From Keyword    True 
        ELSE
            Return From Keyword    False 
        END      
    END
    [Return]  False

Wait For Download Complete
    [Arguments]    ${file_path}    ${timeout}=400
    ${dupe_path}=    Evaluate    r"${file_path}".replace(".pdf"," (1).pdf") if r"${file_path}".lower().endswith(".pdf") else r"${file_path}" + " (1)"
    Run Keyword And Ignore Error    OperatingSystem.Remove File    ${file_path}
    Run Keyword And Ignore Error    OperatingSystem.Remove File    ${dupe_path}
    Wait Until Keyword Succeeds    ${timeout}x    1s    Download Should Exist    ${file_path}    ${dupe_path}
    Wait Until Keyword Succeeds    ${timeout}x    1s    File Should Not Exist    ${file_path}.crdownload

Download Should Exist
    [Arguments]    ${file_path}    ${dupe_path}
    ${has_primary}=    Run Keyword And Return Status    OperatingSystem.File Should Exist    ${file_path}
    IF    ${has_primary}
        RETURN
    END
    ${has_dupe}=    Run Keyword And Return Status    OperatingSystem.File Should Exist    ${dupe_path}
    IF    ${has_dupe}
        Run Keyword And Return Status    File Should Not Exist    ${dupe_path}.crdownload
        OperatingSystem.Move File    ${dupe_path}    ${file_path}
        RETURN
    END
    Fail    Download not found yet
