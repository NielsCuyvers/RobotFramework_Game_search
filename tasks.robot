*** Settings ***
Documentation     Search review scores on the internet, then search the prices on Steam and send them by mail in an Excel file
Library           RPA.Excel.Files
Library           RPA.Outlook.Application
Library           Collections
Library           RPA.Desktop.Windows
Library           RPA.Desktop
Library           RPA.Windows
Library           OperatingSystem
Library           RPA.Browser    auto_close=${FALSE}
Library           String
Library           RPA.Tables
Task Setup        RPA.Outlook.Application.Open Application
Suite Teardown    RPA.Outlook.Application.Quit Application

*** Variables ***
${email_attachments_dir}=    ${OUTPUT_DIR}${/}email_attachments

*** Keywords ***
Check for a certain email
    # Aanmaken van een nieuwe folder om hierin de toegevoegde bestanden toe te voegen
    Create Directory    ${email_attachments_dir}
    ${emails}=    Get Emails
    ...    account_name=nielsc97@gmail.com
    ...    folder_name=Inbox
    ...    email_filter=[Subject]='Games'
    ...    save_attachments=${TRUE}
    ...    attachment_folder=${email_attachments_dir}
    ...    sort=${TRUE}
    ...    sort_key=Received
    ...    sort_descending=${FALSE}
    ${emails_length}    Get Length    ${emails}
    IF    ${emails_length} > 0
        ${Sender}=    Set Variable    ${emails}[0][Sender]
        Set Global Variable    ${Sender}
        ${Email_subject}=    Set Variable    ${emails}[0][Subject]
        Set Global Variable    ${Email_subject}
        Log    ${emails}[0]
        ${attachments}=    Set Variable    ${emails}[0][Attachments]
        Log    ${attachments}
        FOR    ${attachment}    IN    ${attachments}
            Log    ${attachment}
            ${File}=    String.Get Regexp Matches    ${attachment}[0][filename]    .+\.xlsx
            ${File}=    Set Variable    ${File}[0]
            Log    ${file}
            Set Global Variable    ${File}
        END
    ELSE
        Fail    No new mails
    END

Get games from Excel file
    Open Workbook    ${email_attachments_dir}${/}${File}
    ${table}=    Read Worksheet    name=Sheet1    header=${True}
    ${Games}    Create List
    FOR    ${game}    IN    @{table}
        Exit For Loop If    "${game}[Nr]" == "None"
        ${lower_case_game}=    Set Variable    ${game}[Game]
        # ${lower_case_game}=    Convert To Lower Case    ${lower_case_game}
        Append To List    ${Games}    ${lower_case_game}
    END
    # Log List    list_=${Games}
    Set Global Variable    ${Games}
    Close Workbook

Open Google in a browser
    Open Available Browser    url=https://www.google.com
    Sleep    2

Check game alt name
    ${Games_alt_name}    Create List
    ${cookie_popup_EN}=    RPA.Browser.Is Element Visible    locator=xpath://h1[contains(text(),"Before you continue to Google Search")]
    ${cookie_popup_NL}=    RPA.Browser.Is Element Visible    locator=xpath://h1[contains(text(),"Voordat je verdergaat naar Google Zoeken")]
    IF    ${cookie_popup_EN} == ${TRUE}
        Click Element    locator=xpath://div[contains(text(),"I agree")]
        Sleep    1
    END
    IF    ${cookie_popup_NL} == ${TRUE}
        Click Element    locator=xpath://div[contains(text(),"Ik ga akkoord")]
        Sleep    1
    END
    ${index}    Set Variable    ${0}
    FOR    ${game}    IN    @{Games}
        Wait Until Element Is Visible    xpath://input[@name="q"]
        RPA.Browser.Input Text    xpath://input[@name="q"]    text=${game}
        RPA.Browser.Press Keys    xpath://input[@name="q"]    ENTER
        ${name_of_game_in_h4}=    RPA.Browser.Is Element Visible    locator=xpath://h2[@data-attrid="title"]
        IF    ${name_of_game_in_h4} == ${TRUE}
            ${temp_name}=    RPA.Browser.Get Text    locator=xpath://h2[@data-attrid="title"]
            # ${temp_name}=    Convert To Lower Case    ${temp_name}
            Append To List    ${Games_alt_name}    ${temp_name}
        ELSE
            Append To List    ${Games_alt_name}    ${game}
        END
        ${index}    Evaluate    ${index} + 1
    END
    FOR    ${game}    IN    @{Games}
        Log    ${game}
    END
    Set Global Variable    ${Games_alt_name}
    Close Browser

Open Gamespot in a browser
    Open Available Browser    url=https://www.gamespot.com/search/?i=site&q=

Get Gamespots review score
    [Arguments]    ${game}
    RPA.Browser.Wait Until Element Is Visible    id:search-main
    Sleep    1
    RPA.Browser.Input Text    id:search-main    ${game}
    Sleep    0.5
    RPA.Browser.Press Keys    id:search-main    ENTER
    # RPA.Browser.Wait Until Element Is Visible    xpath://h4/span/a[contains(text(),"${game}")]    timeout=30
    # RPA.Browser.Click Element    xpath://h4/span/a[contains(text(),"${game}")]
    ${game_found}=    RPA.Browser.Is Element Visible    xpath://h4/span/a[contains(text(),"${game}")]
    ${first_game_found}    RPA.Browser.Is Element Visible    xpath://h4/span/a
    ${result}=    Set Variable    0
    Set Suite Variable    ${result}    ""
    IF    ${game_found} == ${TRUE}
        Click Element    xpath://h4/span/a[contains(text(),"${game}")]
        RPA.Browser.Wait Until Element Is Visible    class:gs-score__cell    timeout=30
        ${result}=    RPA.Browser.Get Text    class:gs-score__cell
        RPA.Browser.Go Back
    ELSE
        ${result}=    Set Variable    Game not found
    END
    [Return]    ${result}

Get Gamespots review scores
    ${Review_results}=    Create List
    ${counter}=    Set Variable    0
    ${games_size}    Get Length    ${Games}
    FOR    ${counter}    IN RANGE    ${games_size}
        ${cookie_popup}=    RPA.Browser.Is Element Visible    id:onetrust-accept-btn-handler
        IF    ${cookie_popup} == ${TRUE}
            RPA.Browser.Click button    id:onetrust-accept-btn-handler
            RPA.Browser.Wait Until Element Is Not Visible    class:onetrust-pc-dark-filter
        END
        ${game}    Get From List    ${Games}    ${counter}
        ${result}=    Get Gamespots review score    ${game}
        IF    "${result}" == "Game not found"
            ${game}    Get From List    ${Games_alt_name}    ${counter}
            ${result}=    Get Gamespots review score    ${game}
        END
        Append To List    ${Review_results}    ${result}
    END
    Set Global Variable    ${Review_results}
    Close Browser

Write review scores to Excel file
    Open Workbook    ${email_attachments_dir}${/}${File}
    Set Cell Value    1    3    Review
    ${counter}=    Set Variable    ${2}
    FOR    ${result}    IN    @{Review_results}
        Set Cell Value    ${counter}    3    ${result}
        ${counter}    Evaluate    ${counter} + 1
        Sleep    0.5
    END
    Save Workbook
    Close Workbook

Open Steam
    ${appName}=    Set Variable    Steam
    Open From Search    ${appName}    ${appName}    timeout=60

Search for price on Steam
    [Arguments]    ${game}
    ${game}    Convert To Lower Case    ${game}
    ${price_result}    Set Variable    0
    RPA.Desktop.Windows.Wait For Element    store_nav_search_term    timeout=60    interval=0.5
    Sleep    1
    RPA.Desktop.Windows.Mouse Click    store_nav_search_term
    Sleep    0.5
    RPA.Desktop.Press Keys    ctrl    a
    RPA.Desktop.Type Text    ${game}
    RPA.Desktop.Press Keys    ENTER
    Sleep    1
    ${Steam_Search_id}    Get Attribute    name:'Steam Search'    AutomationId
    RPA.Desktop.Windows.Wait For Element    ${Steam_Search_id}    timeout=60
    Sleep    1
    ${steam_page_in_text}=    RPA.Desktop.Windows.Get Text    ${Steam_Search_id}
    Log    ${steam_page_in_text}
    Sleep    1
    ${begin_search_results}=    String.Get Regexp Matches    ${steam_page_in_text}[children_texts]    your search.+
    ${begin_search_result}=    Get From List    ${begin_search_results}    0
    ${begin_search_result}=    Convert To Lower Case    ${begin_search_result}
    ${results}=    String.Get Regexp Matches    ${begin_search_result}    ${game}.+?€
    ${results_length}=    Get Length    ${results}
    IF    ${results_length} > 0
        ${result}=    Get From List    ${results}    0
        ${result_includes_date}=    Run Keyword And Return Status    Should Match Regexp    ${result}    (jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)
        Log    ${result_includes_date}
        IF    ${result_includes_date} == ${TRUE}
            ${price}=    String.Get Regexp Matches    ${result}    ,.+?€
            ${price}=    Get From List    ${price}    0
            ${price}=    String.Get Substring    ${price}    6
            ${result_has_discount}=    Run Keyword And Return Status    Should Match Regexp    ${price}    %
            IF    ${result_has_discount} == ${TRUE}
                ${results}=    String.Get Regexp Matches    ${begin_search_result}    ${game}.+${price}.+?€
                ${result}=    Get From List    ${results}    0
                ${get_prices}=    String.Get Regexp Matches    ${result}    pattern= .+?€
                ${get_prices_length}=    Get Length    ${get_prices}
                ${price}=    Get From List    ${get_prices}    ${get_prices_length-1}
            END
            ${result_is_free}=    Run Keyword And Return Status    Should Match Regexp    ${price}    ^free
            IF    ${result_is_free} == ${TRUE}
                ${price}=    Set Variable    Free
            END
            ${price_result}    Set Variable    ${price}
        END
    ELSE
        ${price_result}    Set Variable    Price not found
    END
    [Return]    ${price_result}

Search for prices on Steam
    ${Price_results}=    Create List
    ${counter}    Set Variable    ${0}
    ${game_list_size}    Get Length    ${Games}
    Log List    list_=${Games}
    FOR    ${counter}    IN RANGE    ${game_list_size}
        ${item}    Get From List    ${Games}    ${counter}
        ${price}=    Search for price on Steam    ${item}
        IF    "${price}" == "Price not found"
            ${item}    Get From List    ${Games_alt_name}    ${counter}
            ${price}=    Search for price on Steam    ${item}
        END
        Append To List    ${Price_results}    ${price}
    END
    Set Global Variable    ${Price_results}

Write prices to Excel file
    Open Workbook    ${email_attachments_dir}${/}${File}
    Set Cell Value    1    4    Price
    ${counter}=    Set Variable    ${2}
    FOR    ${result}    IN    @{Price_results}
        Set Cell Value    ${counter}    4    ${result}
        ${counter}    Evaluate    ${counter} + 1
        Sleep    0.5
    END
    Save Workbook
    Close Workbook

Remove email
    RPA.Outlook.Application.Open Application
    RPA.Outlook.Application.Move Emails
    ...    account_name=niels@brobots.be
    ...    source_folder=Inbox
    ...    email_filter=[Subject]='Games'
    ...    target_folder=Deleted Items
    RPA.Outlook.Application.Quit Application

Send mail back
    Send Email    recipients=${Sender}    subject=RE: ${Email_subject}    body=Check the attachment    attachments=${email_attachments_dir}${/}${File}

*** Tasks ***
Do the Office part
    Check for a certain email
    Get games from Excel file

Do the Browser part
    Open Google in a browser
    Check game alt name
    Open Gamespot in a browser
    Get Gamespots review scores
    Write review scores to Excel file

Do the Steam part
    Open Steam
    Search for prices on Steam
    Write prices to Excel file

Send mail back
    Send mail back

Remove email
    Remove email
    [Teardown]    OperatingSystem.Remove Directory    ${email_attachments_dir}    recursive=True
