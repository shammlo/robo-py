*** Settings ***
Documentation       Template robot main suite.

Library             Collections
Library             MyLibrary
Library             MyLibrary
Library             RPA.Excel.Files
Library             RPA.Browser
Resource            keywords.robot
Variables           MyVariables.py


*** Tasks ***
Main
    Create Excel Report
    Open Available Browser
    Read Excel


*** Keywords ***
Create Excel Report
    Create Workbook    result.xlsx
    Save Workbook

Read Excel
    Open Workbook    robot_test.xlsx
    ${list}    Read Worksheet    header=True
    Log To Console    ${list}
    Close Workbook
    FOR    ${index}    IN    @{list}
        Search Cars    ${index}
    END

Search Cars
    [Arguments]    ${index}
    Go To    %{C_URL}
    Maximize Browser Window
    Wait Until Element Is Visible
    ...    xpath:/html/body/div[1]/div/main/div[2]/div[1]/div/div[1]/div[1]/form/div[1]/div[1]/div/div/div/div/div[1]/div[2]
    Click Element
    ...    xpath:/html/body/div[1]/div/main/div[2]/div[1]/div/div[1]/div[1]/form/div[1]/div[1]/div/div/div/div/div[1]/div[2]
    Press Keys    NONE    ${index}[make]
    Sleep    333ms
    Press Keys    NONE    TAB
    Press Keys    NONE    TAB
    Sleep    500ms
    Press Keys    NONE    ${index}[model]
    Sleep    333ms
    Press Keys    NONE    TAB
    Press Keys    NONE    TAB
    Sleep    500ms
    Press Keys    NONE    ${index}[max_km]
    Sleep    500ms
    Click Element
    ...    xpath:/html/body/div[1]/div/main/div[2]/div[1]/div/div[1]/div[1]/form/div[1]/div[3]/div/div/div[1]
    Sleep    500ms
    Click Element    xpath:/html/body/div[1]/div/main/div[2]/div[1]/div/div[1]/div[1]/form/div[2]/div[1]/button/span
    Sleep    3s
    Click Element
    ...    xpath:/html/body/div[1]/div/main/div[2]/div[2]/section/div/div[1]/div[2]/div[1]/div[2]/div/div/div/span
    Sleep    333ms

    # Click Element    xpath=/html/body/div[9]/div/div/div/div[2]/div/div[6]/p[text()='Lowest price']
    # Wait Until Page Contains Element    //span[text()="Add a car"]
    Click Element    //p[text()='Lowest price']
    Sleep    1s
    ${name}    Get Text
    ...    xpath:/html/body/div[1]/div/main/div[2]/div[2]/section/div/div[2]/div[13]/div/a/div/div[2]/h6
    Sleep    1s
    ${total_km}    Get Text    //span[contains(., "km")]
    Sleep    1s
    ${seller}    Get Text    //span[@class="css-r2knv9 eai3djn1"]
    Sleep    1s
    ${country}    Get Text
    ...    xpath:/html/body/div[1]/div/main/div[2]/div[2]/section/div/div[2]/div[1]/div/a/div/div[2]/div[3]/div[1]/div/div[2]/span
    Sleep    1s
    ${fuel}    Get Text
    ...    xpath:/html/body/div[1]/div/main/div[2]/div[2]/section/div/div[2]/div[1]/div/a/div/div[2]/div[1]/div[5]/span[2]
    Sleep    1s

    ${car_dict}    Create Dictionary
    ...    name=${name}
    ...    total_km=${total_km}
    ...    seller=${seller}
    ...    country=${country}
    ...    fuel=${fuel}
    Log To Console    ${car_dict}

    Append Excel    ${car_dict}

Append Excel
    [Arguments]    ${car_dict}
    Open Workbook    result.xlsx
    Append Rows To Worksheet    ${car_dict}    header=True
    Save Workbook
