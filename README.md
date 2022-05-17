# **Video games recensie en prijs**
## **Waarneming**
Dit project is tot stand gekomen met gebruikmakend van het Robot framework.

## **Uitvoering**
Hieronder wordt er stap voor stap uitgelegt hoe het project is gecodeerd.
De stappen zijn ingedeeld in vier hoofdstukken, namelijk:
* Settings
    * [Implementaties](#implementaties-settingssettings)
        * [Libraries](#libraries)
        * [Setup](#setup)
* [Variables](#variabelen)
* Keywords
    * [Office](#office-keywords)
    * [Browser](#browser-keywords)
    * [Steam](#steam-keywords)
* Tasks
    * [Samenvoeging](#samenvoeging-keywords)

Als u de redenering van Robot Framework wilt volgen staat er achter elke titel in welke sectie de code is neergetypt.
Als u het Robot Framework uitvoert zorg dat alle applicaties die tevoorschijn komen tijdens de uitvoreing op het primare scherm staan. Dit zou normaal automatisch gebeuren.

### **Implementaties (Settings)**
```Robot framework
Documentation     Search review scores on the internet, then search the prices on Steam and send them by mail in an Excel file
```
#### **Libraries**
Voor dit proejct maak ik gebruik van de **Outlook** en **Excel** libraries van RPA. Dit zorgt er voor dat ik email kan uitlezen met hun bijgevoegde bestanden en dat ik Excel bestanden kan uitlezen om zo data te lezen en weg te schrijven.
Er wordt ook gebruik gemaakt van de **Collections** library zodat er lijsten kunnen aangemaakt worden.
Ook is het nodig om interactie te hebben met de browsers maar ook met de GUI van Windows. Dit wordt mogelijk gemaakt door de **Browser**, **Desktop**, **Windows** en **OperatingSystem** libraries. Als laatste wordt er ook gebruik gemaakt van de **String** library om tekst variabelen op te slagen en de **Tables** library om data tabellen te maken.
```Robot framework
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
```
#### **Setup**
Deze keywords maken het mogelijke om iets te laten gebeuren in het begin van het uitvoeren van de code en op het einde van de code, ook al geeft de code een error. Dit is vergelijkbaar met de Try-catch-**finally** declaratie.
```Robot framework
Task Setup        RPA.Outlook.Application.Open Application
Suite Teardown    RPA.Outlook.Application.Quit Application
```
### **Variabelen**
Hier worden de globalen variabelen van het begin van de code aangemaakt. Hier wordt een pad opgeslagen waar later alle **email attachments** worden opgeslagen. De **OUTPUT_DIR** variabelen geeft het geconfigureerde pad mee van het project. Dit is standaard het pad van het project.
```Robot framework
${email_attachments_dir}=    ${OUTPUT_DIR}${/}email_attachments
```
### **Office (Keywords)**
In dit gedeelte gebeurd alles dat te maken heeft met Office 365 applicaties (Excel en Outlook). Alleen het wegschrijven van data naar Excel gebeurt hier niet.
#### **Lezen van mails**
Er wordt een mail ingelezen in een bepaalde inbox met een bepaald onderwerp. Deze mail wordt ingelezen en de toegevoegde bestanden worden lokaal opgeslagen op het systeem.
Hier wordt ook de naam van het bestand gelezen zodat dit later in het project gebruikt kan worden om data in op te slagen.

##### Extra info paramaters
|Parameter|Info|
|------|------|
|folder_name|Verander de naam van de folder als je ergens anders wilt zoeken. Dit is taal gevoelig. (Inbox != Postvak IN)|

```Robot framework
Check for a certain email
    Create Directory    ${email_attachments_dir}
    ${emails}=    Get Emails
    ...    account_name=niels@brobots.be
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
```
#### **Lezen van data uit Excel**
De data uit een gestructureerd Excel bestand wordt gelezen en de data uit de kolom met naam 'Game' wordt opgelaan in een lijst die later in het project gebuikt wordt als opzoek materiaal.
```Robot framework
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
```

#### **Terug sturen van mail met bestand**
Hier wordt een email verstuurd met de behandelde bijlage naar de zender van de binnenkomende email.
```Robot framework
Send mail back
    Send Email    recipients=${Sender}    subject=RE: ${Email_subject}    body=Check the attachment    attachments=${email_attachments_dir}${/}${File}
```

### **Browser (Keywords)**
Hier gebeurt alles dat te maken heeft met browsers. Het openen, opzoeken en lezen van data van websites.
#### **Checken voor alternatieve namen op Google**
Wegens de preciesheid van de robot wordt er online voor elke game een alternative titel schrijfwijzen opgezocht. Zo is er meer kans dat de gevraagde game gevonden kan worden op de recentie website en op de Steam applicatie. Voor de robot is 'Grand Theft Auto V' niet hetzelfde als 'Grand Theft Auto 5'.
##### Openen van een browser
Hey 'Open Avalible Browser' keyword opent de meegegeven website in eender welke geïnstalleerde browser.
```Robot framework
Open Google in a browser
    Open Available Browser    url=https://www.google.com
    Sleep    2
```
##### Checken naar een bestaand HTML element
Wegens dat er gebruik wordt gemaakt van de google zoekmachine kan het mogelijk zijn dat er een pop-up tevoorschijn komt. Deze zou geaccepteerd moeten worden. Het is mogelijk om deze pop-up te doen verwijderen door op een knop te drukken die 'Ik ga akkoord' heet. Deze knop is taal gebonden. Als de pop-up niet verschijnt gaat de code gewoon door.

Vervolgens wordt er gezocht naar de alternative titel schrijfwijzen van elke game. Dit gebeurt door de game op te zoeken en via Xpath te zoeken naar een html h2 element met als attribute 'data-attrid="title"'.
![Voorbeeld alternatieve titel schijfwijzen](images/Example_alt_name.jpg)
 Als deze gevonden is wordt deze opgeslagen in een nieuwe lijst. Anders wordt de gewone schijfwijzen opgeslagen is deze lijst.
```Robot framework
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
```
#### **Nakijken van recensie score op Gamespot**
##### Checken van één game
```Robot framework
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
```
##### Checken van een lijst van games
```Robot framework
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

```
#### **Wegschrijven van recensies naar een Excel bestand**
```Robot framework
Write review score to Excel file
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
```

### **Steam (Keywords)**
#### **Open van een applicatie**
```Robot framework
Open Steam
    ${appName}=    Set Variable    Steam
    Open From Search    ${appName}    ${appName}    timeout=60
```
#### **Zoeken van één prijs in Steam**
```Robot framework
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
```
#### **Zoeken van meerdere prijzen in Steam**
```Robot framework
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
```
#### **Wegschrijven van prijzen naar een Excel bestand**
```Robot framework
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
```







### **Samenvoeging (Tasks)**
Door nu alle Keywords samen te voegen kunnen we het project gebruiken om recensies en prijzen van video games te zoeken.

```Robot framework
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
```



