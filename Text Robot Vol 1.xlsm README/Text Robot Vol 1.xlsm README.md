# <h1 id="oa-robot-definitions">OA Robot Definitions</h1>

\*\*Text Robot Vol 1.xlsm\*\* contains definitions for:

[53 Robot Commands](#command-definitions)<BR>[8 Robot Texts](#text-definitions)<BR>

<BR>

## Available Robot Commands

| Name | Description |
| --- | --- |
| [Append Specified Text](#append-specified-text) | Append the user provided text to the right of the selected text. |
| [Convert To Camel Case](#convert-to-camel-case) | Convert text in selection to camel case. |
| [Convert To Kabob Case](#convert-to-kabob-case) | Convert text in selection to kabob case. |
| [Convert To Lower Case](#convert-to-lower-case) | Convert text in selection to lower case. |
| [Convert To Pascal Case](#convert-to-pascal-case) | Convert the text in selection to pascal case. |
| [Convert To Proper Case](#convert-to-proper-case) | Convert text in selection to proper case. |
| [Convert To Upper Case](#convert-to-upper-case) | Convert text in selection to upper case. |
| [Keep First Character](#keep-first-character) | Returns first character in the active cell. |
| [Keep First N Characters](#keep-first-n-characters) | Returns first N characters in the active cell. |
| [Keep First Word](#keep-first-word) | Return the first word in the active cell. |
| [Keep Last Character](#keep-last-character) | Returns last character in the active cell. |
| [Keep Last N Characters](#keep-last-n-characters) | Returns last N characters in the active cell. |
| [Keep Last Word](#keep-last-word) | Return the last word in the active cell. |
| [Keep Nth Word](#keep-nth-word) | Return the Nth word in the active cell (use negatives to count from the end). |
| [Keep Text After Specified Text](#keep-text-after-specified-text) | Keeps the text right of the first instance of the specified text. |
| [Keep Text Before Specified Text](#keep-text-before-specified-text) | Keeps the text left of the last instance of the specified text. |
| [Left Trim Spaces](#left-trim-spaces) | Removes spaces from start of a text string. |
| [List Characters](#list-characters) | Spill characters of text in active cell to below. |
| [List Words](#list-words) | Spill words of text in active cell to below. |
| [Prepend Specified Text](#prepend-specified-text) | Prepend the user provided text to the left of the selected text. |
| [Remove First N Characters](#remove-first-n-characters) | Remove the first N characters of the text in selection. |
| [Remove First N Words](#remove-first-n-words) | Removes the first N words from the text in the selection. |
| [Remove First Word](#remove-first-word) | Removes the first word from the text in the selection. |
| [Remove Last N Characters](#remove-last-n-characters) | Remove the last N characters of the text in selection. |
| [Remove Last N Words](#remove-last-n-words) | Removes the last N words from the text in the selection. |
| [Remove Last Word](#remove-last-word) | Removes the last word from the text in the selection. |
| [Remove Specified Text](#remove-specified-text) | Remove specified text from the selection. |
| [Remove Text After Specified Text](#remove-text-after-specified-text) | Removes the text right of the first instance of the specified text. |
| [Remove Text Before Specified Text](#remove-text-before-specified-text) | Removes the text left of the last instance of the specified text. |
| [Replace Commas With Pipes](#replace-commas-with-pipes) | Replace commas with pipes in selected text. |
| [Replace Commas With Semicolons](#replace-commas-with-semicolons) | Replace commas with semicolons in selected text. |
| [Replace Curly Braces With Parenthesis](#replace-curly-braces-with-parenthesis) | Replace curly braces with parenthesis in selected text. |
| [Replace Curly Braces With Square Brackets](#replace-curly-braces-with-square-brackets) | Replace curly braces with square brackets in selected text. |
| [Replace Parenthesis With Curly Braces](#replace-parenthesis-with-curly-braces) | Replace parenthesis with curly braces in selected text. |
| [Replace Parenthesis With Square Brackets](#replace-parenthesis-with-square-brackets) | Replace parenthesis with square brackets in selected text. |
| [Replace Pipes With Commas](#replace-pipes-with-commas) | Replace pipes with commas in selected text. |
| [Replace Pipes With Semicolons](#replace-pipes-with-semicolons) | Replace pipes with semicolons in selected text. |
| [Replace Semicolons With Commas](#replace-semicolons-with-commas) | Replace semicolons with commas in selected text. |
| [Replace Semicolons With Pipes](#replace-semicolons-with-pipes) | Replace semicolons with pipes in selected text. |
| [Replace Spaces With Hyphens](#replace-spaces-with-hyphens) | Replace spaces with hyphens in selected text. |
| [Replace Spaces With Periods](#replace-spaces-with-periods) | Replace spaces with periods in selected text. |
| [Replace Spaces With Underscores](#replace-spaces-with-underscores) | Replace spaces with underscores in selected text. |
| [Replace Square Brackets With Curly Braces](#replace-square-brackets-with-curly-braces) | Replace square brackets with curly braces in selected text. |
| [Replace Square Brackets With Parenthesis](#replace-square-brackets-with-parenthesis) | Replace square brackets with parenthesis in selected text. |
| [Replace Underscores With Hyphens](#replace-underscores-with-hyphens) | Replace underscores with hyphens in selected text. |
| [Replace Underscores With Spaces](#replace-underscores-with-spaces) | Replace underscores with spaces in selected text. |
| [Right Trim Spaces](#right-trim-spaces) | Removes spaces from end of a text string. |
| [Split By Characters](#split-by-characters) | Spill characters of text in active cell to the right. |
| [Split By N Characters](#split-by-n-characters) | Spill characters of text in active cell to the right in chunks of length N. |
| [Split By Words](#split-by-words) | Spill words of text in active cell to the right. |
| [Trim Spaces](#trim-spaces) | Removes all spaces from a text string except single spaces between words. |
| [Trim Spaces From End](#trim-spaces-from-end) | Removes spaces from end of a text string. |
| [Trim Spaces From Start](#trim-spaces-from-start) | Removes spaces from start of a text string. |

<BR>

## Available Robot Texts

| Name | Description |
| --- | --- |
| [CamelCase.lambda](#camelcaselambda) | Definition of CamelCase lambda function. |
| [Characters.lambda](#characterslambda) | Definition of Characters lambda function. |
| [KabobCase.lambda](#kabobcaselambda) | Definition of KabobCase lambda function. |
| [PascalCase.lambda](#pascalcaselambda) | Definition of PascalCase lambda function. |
| [Substitutions.lambda](#substitutionslambda) | Definition of Substitutions lambda function. |
| [TrimEnd.lambda](#trimendlambda) | Definition of TrimEnd lambda function. |
| [TrimStart.lambda](#trimstartlambda) | Definition of TrimStart lambda function. |
| [Words.lambda](#wordslambda) | Definition of Words lambda function. |

<BR>

## Command Definitions

<BR>

### Append Specified Text

*Append the user provided text to the right of the selected text.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=(\[\[ActiveCell::Formula\]\])&"{{Text\_To\_Append}}"</code> |

[^Top](#oa-robot-definitions)

<BR>

### Convert To Camel Case

*Convert text in selection to camel case.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=CamelCase(\[\[ActiveCell::Formula\]\])</code> |
| Formula Dependencies | [CamelCase.lambda](#camelcaselambda) |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

### Convert To Kabob Case

*Convert text in selection to kabob case.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=KabobCase(\[\[ActiveCell::Formula\]\])</code> |
| Formula Dependencies | [KabobCase.lambda](#kabobcaselambda) |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

### Convert To Lower Case

*Convert text in selection to lower case.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=LOWER(\[\[ActiveCell::Formula\]\])</code> |
| User Context Filter | ExcelActiveCellValueIsText |
| Launch Codes | <code>lc</code> |

[^Top](#oa-robot-definitions)

<BR>

### Convert To Pascal Case

*Convert the text in selection to pascal case.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=PascalCase(\[\[ActiveCell::Formula\]\])</code> |
| Formula Dependencies | [PascalCase.lambda](#pascalcaselambda) |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

### Convert To Proper Case

*Convert text in selection to proper case.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=PROPER(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellValueIsText |
| Launch Codes | <code>pc</code> |

[^Top](#oa-robot-definitions)

<BR>

### Convert To Upper Case

*Convert text in selection to upper case.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=UPPER(\[\[ActiveCell::Formula\]\])</code> |
| User Context Filter | ExcelActiveCellValueIsText |
| Launch Codes | <code>uc</code> |

[^Top](#oa-robot-definitions)

<BR>

### Keep First Character

*Returns first character in the active cell.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=LEFT(\[\[ActiveCell::Formula\]\])</code> |
| Launch Codes | <ol><li><code>l</code></li><li><code>fc</code></li><li><code>kfc</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Keep First N Characters

*Returns first N characters in the active cell.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=LEFT(\[\[ActiveCell::Formula\]\],{{N}})</code> |
| Launch Codes | <ol><li><code>l</code></li><li><code>fc</code></li><li><code>kfc</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Keep First Word

*Return the first word in the active cell.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=TAKE(TEXTSPLIT(\[\[ActiveCell::Formula\]\]," "),,1)</code> |
| Launch Codes | <ol><li><code>fw</code></li><li><code>kfw</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Keep Last Character

*Returns last character in the active cell.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=RIGHT(\[\[ActiveCell::Formula\]\])</code> |
| Launch Codes | <ol><li><code>r</code></li><li><code>lc</code></li><li><code>klc</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Keep Last N Characters

*Returns last N characters in the active cell.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=RIGHT(\[\[ActiveCell::Formula\]\],{{N}})</code> |
| Launch Codes | <ol><li><code>r</code></li><li><code>lc</code></li><li><code>klc</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Keep Last Word

*Return the last word in the active cell.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=TAKE(TEXTSPLIT(\[\[ActiveCell::Formula\]\]," "),,\-1)</code> |
| Launch Codes | <ol><li><code>lw</code></li><li><code>klw</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Keep Nth Word

*Return the Nth word in the active cell (use negatives to count from the end).*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=TAKE(TAKE(TEXTSPLIT(\[\[ActiveCell::Formula\]\]," "),,{{N}}),,\-SIGN({{N}}))</code> |
| Parameters | <ol><li>[N](#keep-nth-word--n)</li></ol> |
| Outputs | <ol></ol> |
| Launch Codes | <ol><li><code>nw</code></li><li><code>knw</code></li></ol> |

<BR>

#### Keep Nth Word \>\> N

<sup>`!Input Parameter` </sup>

| Property | Value |
| --- | --- |
| Prompt | <code>Enter N to return Nth word. Enter an integer or a cell reference.</code><br><code></code><br><code>(Use negatives to count from the end.)</code> |
| Caching Policy | CacheForDurationOfCommandExecution |

[^Top](#oa-robot-definitions)

<BR>

### Keep Text After Specified Text

*Keeps the text right of the first instance of the specified text.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=TEXTAFTER(\[\[ActiveCell::Formula\]\],"{{Specified\_Text}}")</code> |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

### Keep Text Before Specified Text

*Keeps the text left of the last instance of the specified text.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=TEXTBEFORE(\[\[ActiveCell::Formula\]\],"{{Specified\_Text}}",\-1)</code> |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

### Left Trim Spaces

*Removes spaces from start of a text string.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=TrimStart(\[\[ActiveCell::Formula\]\], " ")</code> |
| Formula Dependencies | [TrimStart.lambda](#trimstartlambda) |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

### List Characters

*Spill characters of text in active cell to below.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=Characters(\[\[ActiveCell::Formula\]\], TRUE)</code> |
| Formula Dependencies | [Characters.lambda](#characterslambda) |
| User Context Filter | ExcelActiveCellValueIsText |
| Launch Codes | <code>c</code> |

[^Top](#oa-robot-definitions)

<BR>

### List Words

*Spill words of text in active cell to below.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=Words(\[\[ActiveCell::Formula\]\], TRUE)</code> |
| Formula Dependencies | [Words.lambda](#wordslambda) |
| User Context Filter | ExcelActiveCellValueIsText |
| Launch Codes | <code>w</code> |

[^Top](#oa-robot-definitions)

<BR>

### Prepend Specified Text

*Prepend the user provided text to the left of the selected text.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\="{{Text\_To\_Prepend}}"&(\[\[ActiveCell::Formula\]\])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Remove First N Characters

*Remove the first N characters of the text in selection.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=MID(\[\[ActiveCell::Formula\]\],{{N}}+1,999999)</code> |
| Parameters | <ol><li>[N](#remove-first-n-characters--n)</li></ol> |
| User Context Filter | ExcelActiveCellValueIsText |
| Outputs | <ol></ol> |

<BR>

#### Remove First N Characters \>\> N

*Number of characters*

<sup>`!Input Parameter` </sup>

| Property | Value |
| --- | --- |
| Prompt | <code>Enter number of characters to remove:</code> |
| Data Type | Integer |

[^Top](#oa-robot-definitions)

<BR>

### Remove First N Words

*Removes the first N words from the text in the selection.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=TEXTJOIN(" ",TRUE,DROP(TEXTSPLIT(\[\[ActiveCell::Formula\]\]," "),,{{Number\_Of\_Words\_To\_Remove}}))</code> |

[^Top](#oa-robot-definitions)

<BR>

### Remove First Word

*Removes the first word from the text in the selection.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=TEXTJOIN(" ",TRUE,DROP(TEXTSPLIT(\[\[ActiveCell::Formula\]\]," "),,1))</code> |

[^Top](#oa-robot-definitions)

<BR>

### Remove Last N Characters

*Remove the last N characters of the text in selection.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=LEFT(\[\[ActiveCell::Formula\]\],MAX(LEN(\[\[ActiveCell::Formula\]\])\-{{N}},0))</code> |
| Parameters | <ol><li>[N](#remove-last-n-characters--n)</li></ol> |
| User Context Filter | ExcelActiveCellValueIsText |
| Outputs | <ol></ol> |

<BR>

#### Remove Last N Characters \>\> N

*Number of characters*

<sup>`!Input Parameter` </sup>

| Property | Value |
| --- | --- |
| Prompt | <code>Enter number of characters to remove:</code> |
| Data Type | Integer |

[^Top](#oa-robot-definitions)

<BR>

### Remove Last N Words

*Removes the last N words from the text in the selection.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=TEXTJOIN(" ",TRUE,DROP(TEXTSPLIT(\[\[ActiveCell::Formula\]\]," "),,\-{{Number\_Of\_Words\_To\_Remove}}))</code> |

[^Top](#oa-robot-definitions)

<BR>

### Remove Last Word

*Removes the last word from the text in the selection.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=TEXTJOIN(" ",TRUE,DROP(TEXTSPLIT(\[\[ActiveCell::Formula\]\]," "),,\-1))</code> |

[^Top](#oa-robot-definitions)

<BR>

### Remove Specified Text

*Remove specified text from the selection.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=SUBSTITUTE(\[\[ActiveCell::Formula\]\],"{{Text\_To\_Remove}}","")</code> |

[^Top](#oa-robot-definitions)

<BR>

### Remove Text After Specified Text

*Removes the text right of the first instance of the specified text.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=TEXTBEFORE(\[\[ActiveCell::Formula\]\],"{{Specified\_Text}}")&"{{Specified\_Text}}"</code> |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

### Remove Text Before Specified Text

*Removes the text left of the last instance of the specified text.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\="{{Specified\_Text}}"&TEXTAFTER(\[\[ActiveCell::Formula\]\],"{{Specified\_Text}}",\-1)</code> |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

### Replace Commas With Pipes

*Replace commas with pipes in selected text.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=SUBSTITUTE(\[\[ActiveCell::Formula\]\],",","\\|")</code> |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

### Replace Commas With Semicolons

*Replace commas with semicolons in selected text.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=SUBSTITUTE(\[\[ActiveCell::Formula\]\],",",";")</code> |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

### Replace Curly Braces With Parenthesis

*Replace curly braces with parenthesis in selected text.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=Substitutions(\[\[ActiveCell::Formula\]\], {"{","}"}, {"(",")"})</code> |
| Formula Dependencies | [Substitutions.lambda](#substitutionslambda) |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

### Replace Curly Braces With Square Brackets

*Replace curly braces with square brackets in selected text.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=Substitutions(\[\[ActiveCell::Formula\]\], {"{","}"}, {"\[","\]"})</code> |
| Formula Dependencies | [Substitutions.lambda](#substitutionslambda) |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

### Replace Parenthesis With Curly Braces

*Replace parenthesis with curly braces in selected text.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=Substitutions(\[\[ActiveCell::Formula\]\], {"(",")"}, {"{","}"})</code> |
| Formula Dependencies | [Substitutions.lambda](#substitutionslambda) |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

### Replace Parenthesis With Square Brackets

*Replace parenthesis with square brackets in selected text.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=Substitutions(\[\[ActiveCell::Formula\]\], {"(",")"}, {"\[","\]"})</code> |
| Formula Dependencies | [Substitutions.lambda](#substitutionslambda) |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

### Replace Pipes With Commas

*Replace pipes with commas in selected text.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=SUBSTITUTE(\[\[ActiveCell::Formula\]\],"\\|",",")</code> |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

### Replace Pipes With Semicolons

*Replace pipes with semicolons in selected text.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=SUBSTITUTE(\[\[ActiveCell::Formula\]\],"\\|",";")</code> |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

### Replace Semicolons With Commas

*Replace semicolons with commas in selected text.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=SUBSTITUTE(\[\[ActiveCell::Formula\]\],";",",")</code> |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

### Replace Semicolons With Pipes

*Replace semicolons with pipes in selected text.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=SUBSTITUTE(\[\[ActiveCell::Formula\]\],";","\\|")</code> |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

### Replace Spaces With Hyphens

*Replace spaces with hyphens in selected text.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=SUBSTITUTE(\[\[ActiveCell::Formula\]\]," ","\-")</code> |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

### Replace Spaces With Periods

*Replace spaces with periods in selected text.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=SUBSTITUTE(\[\[ActiveCell::Formula\]\]," ",".")</code> |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

### Replace Spaces With Underscores

*Replace spaces with underscores in selected text.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=SUBSTITUTE(\[\[ActiveCell::Formula\]\]," ","\_")</code> |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

### Replace Square Brackets With Curly Braces

*Replace square brackets with curly braces in selected text.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=Substitutions(\[\[ActiveCell::Formula\]\], {"\[","\]"}, {"{","}"})</code> |
| Formula Dependencies | [Substitutions.lambda](#substitutionslambda) |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

### Replace Square Brackets With Parenthesis

*Replace square brackets with parenthesis in selected text.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=Substitutions(\[\[ActiveCell::Formula\]\], {"\[","\]"}, {"(",")"})</code> |
| Formula Dependencies | [Substitutions.lambda](#substitutionslambda) |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

### Replace Underscores With Hyphens

*Replace underscores with hyphens in selected text.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=SUBSTITUTE(\[\[ActiveCell::Formula\]\],"\_","\-")</code> |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

### Replace Underscores With Spaces

*Replace underscores with spaces in selected text.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=SUBSTITUTE(\[\[ActiveCell::Formula\]\],"\_"," ")</code> |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

### Right Trim Spaces

*Removes spaces from end of a text string.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=TrimEnd(\[\[ActiveCell::Formula\]\], " ")</code> |
| Formula Dependencies | [TrimEnd.lambda](#trimendlambda) |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

### Split By Characters

*Spill characters of text in active cell to the right.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=Characters(\[\[ActiveCell::Formula\]\])</code> |
| Formula Dependencies | [Characters.lambda](#characterslambda) |
| User Context Filter | ExcelActiveCellValueIsText |
| Launch Codes | <code>c</code> |

[^Top](#oa-robot-definitions)

<BR>

### Split By N Characters

*Spill characters of text in active cell to the right in chunks of length N.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=Characters(\[\[ActiveCell::Formula\]\],,{{N}})</code> |
| Formula Dependencies | [Characters.lambda](#characterslambda) |
| User Context Filter | ExcelActiveCellValueIsText |
| Launch Codes | <code>c</code> |

[^Top](#oa-robot-definitions)

<BR>

### Split By Words

*Spill words of text in active cell to the right.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=Words(\[\[ActiveCell::Formula\]\])</code> |
| Formula Dependencies | [Words.lambda](#wordslambda) |
| User Context Filter | ExcelActiveCellValueIsText |
| Launch Codes | <code>w</code> |

[^Top](#oa-robot-definitions)

<BR>

### Trim Spaces

*Removes all spaces from a text string except single spaces between words.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=TRIM(\[\[ActiveCell::Formula\]\])</code> |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

### Trim Spaces From End

*Removes spaces from end of a text string.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=TrimEnd(\[\[ActiveCell::Formula\]\], " ")</code> |
| Formula Dependencies | [TrimEnd.lambda](#trimendlambda) |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

### Trim Spaces From Start

*Removes spaces from start of a text string.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=TrimStart(\[\[ActiveCell::Formula\]\], " ")</code> |
| Formula Dependencies | [TrimStart.lambda](#trimstartlambda) |
| User Context Filter | ExcelActiveCellValueIsText |

[^Top](#oa-robot-definitions)

<BR>

## Text Definitions

<BR>

### CamelCase.lambda

*Definition of CamelCase lambda function.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [CamelCase.lambda](<./Text/CamelCase.lambda.txt>) |
| Value | <code>CamelCase \= LAMBDA(input, LET(</code><br><code> \\\\LambdaName, "CamelCase",</code><br><code> Words, LAMBDA(text,\[spill\_down\], LET(</code><br><code> \_Words, TEXTSPLIT(text, " "),</code><br><code> \_Transpose, IF(ISOMITTED(spill\_down), FALSE, spill\_down),</code><br><code> \_Result, IF(\_Transpose, TRANSPOSE(\_Words), \_Words),</code><br><code> \_Result</code><br><code> )),</code><br><code> \_Words, Words(input, 1),</code><br><code> \_FirstWord, LOWER(TAK... |
| Content Type | ExcelFormula |
| Location | <code>CamelCase</code> |

[^Top](#oa-robot-definitions)

<BR>

### Characters.lambda

*Definition of Characters lambda function.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [Characters.lambda](<./Text/Characters.lambda.txt>) |
| Value | <code>Characters \= LAMBDA(text,\[spill\_down\],\[chunk\_size\], LET(</code><br><code> \\\\LambdaName, "Characters",</code><br><code> \_ChunkSize, IF(chunk\_size \= 0, 1, chunk\_size),</code><br><code> \_Characters, MID(text, SEQUENCE(, ROUNDUP(LEN(text) \/ \_ChunkSize, 0), 1, \_ChunkSize), \_ChunkSize),</code><br><code> \_Transpose, IF(ISOMITTED(spill\_down), FALSE, spill\_down),</code><br><code> \_Result, IF(\_Transpose, TRANSPOSE(\_Characters), \_Characters),</code><br><code> \_Re... |
| Content Type | ExcelFormula |
| Location | <code>Characters</code> |

[^Top](#oa-robot-definitions)

<BR>

### KabobCase.lambda

*Definition of KabobCase lambda function.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [KabobCase.lambda](<./Text/KabobCase.lambda.txt>) |
| Value | <code>\/\*Convert text in selection to kabob case. \*\/</code><br><code>KabobCase \= LAMBDA(input, LET(</code><br><code> \\\\LambdaName, "KabobCase",</code><br><code> \\\\CommandName, "Convert To Kabob Case",</code><br><code> \\\\Description, "Convert text in selection to kabob case.",</code><br><code> \_Words, Words(input, 1),</code><br><code> \_Result, TEXTJOIN("\-", , LOWER(\_Words)),</code><br><code> \_Result</code><br><code>));</code> |
| Content Type | ExcelFormula |
| Location | <code>KabobCase</code> |

[^Top](#oa-robot-definitions)

<BR>

### PascalCase.lambda

*Definition of PascalCase lambda function.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [PascalCase.lambda](<./Text/PascalCase.lambda.txt>) |
| Value | <code>\/\*Convert the text in selection to pascal case. \*\/</code><br><code>PascalCase \= LAMBDA(input, LET(</code><br><code> \\\\LambdaName, "PascalCase",</code><br><code> \\\\CommandName, "Conver To Pascal Case",</code><br><code> \\\\Description, "Convert the text in selection to pascal case.",</code><br><code> \_Words, Words(input, 1),</code><br><code> \_Result, CONCAT(PROPER(\_Words)),</code><br><code> \_Result</code><br><code>));</code> |
| Content Type | ExcelFormula |
| Location | <code>PascalCase</code> |

[^Top](#oa-robot-definitions)

<BR>

### Substitutions.lambda

*Definition of Substitutions lambda function.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [Substitutions.lambda](<./Text/Substitutions.lambda.txt>) |
| Value | <code>Substitutions \= LAMBDA(text,old\_texts,new\_texts, LET(</code><br><code> \\\\LambdaName, "Substitutions",</code><br><code> \_OldTexts, TOCOL(old\_texts),</code><br><code> \_NewTexts, TOCOL(new\_texts),</code><br><code> \_Seq, SEQUENCE(ROWS(\_OldTexts)),</code><br><code> \_Result, REDUCE(</code><br><code> text,</code><br><code> \_Seq,</code><br><code> LAMBDA(txt,idx,</code><br><code> SUBSTITUTE(txt, INDEX(\_OldTexts, idx, 1), INDEX(\_NewTexts, idx, 1))</cod... |
| Content Type | ExcelFormula |
| Location | <code>Substitutions</code> |

[^Top](#oa-robot-definitions)

<BR>

### TrimEnd.lambda

*Definition of TrimEnd lambda function.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [TrimEnd.lambda](<./Text/TrimEnd.lambda.txt>) |
| Value | <code>TrimEnd \= LAMBDA(text,\[character\], LET(</code><br><code> \\\\LambdaName, "TrimEnd",</code><br><code> \_TrimChar, IF(ISOMITTED(character), " ", character),</code><br><code> \_Match, MID(text, SEQUENCE(LEN(text)), 1) \= \_TrimChar,</code><br><code> \_Last, XMATCH(FALSE, \_Match, , \-1),</code><br><code> \_Result, IF(ISNA(\_Last), "", LEFT(text, \_Last)),</code><br><code> \_Result</code><br><code>));</code> |
| Content Type | ExcelFormula |
| Location | <code>TrimEnd</code> |

[^Top](#oa-robot-definitions)

<BR>

### TrimStart.lambda

*Definition of TrimStart lambda function.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [TrimStart.lambda](<./Text/TrimStart.lambda.txt>) |
| Value | <code>TrimStart \= LAMBDA(text,\[character\], LET(</code><br><code> \\\\LambdaName, "TrimStart",</code><br><code> \_TrimChar, IF(ISOMITTED(character), " ", character),</code><br><code> \_Chars, MID(text, SEQUENCE(LEN(text)), 1),</code><br><code> \_Match, \_Chars \= \_TrimChar,</code><br><code> \_First, XMATCH(FALSE, \_Match),</code><br><code> \_Result, IF(ISNA(\_First), "", MID(text, \_First, LEN(text))),</code><br><code> \_Result</code><br><code>));</code> |
| Content Type | ExcelFormula |
| Location | <code>TrimStart</code> |

[^Top](#oa-robot-definitions)

<BR>

### Words.lambda

*Definition of Words lambda function.*

<sup>`@Text Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [Words.lambda](<./Text/Words.lambda.txt>) |
| Value | <code>Words \= LAMBDA(text,\[spill\_down\], LET(</code><br><code> \\\\LambdaName, "Words",</code><br><code> \_Words, TEXTSPLIT(text, " "),</code><br><code> \_Transpose, IF(ISOMITTED(spill\_down), FALSE, spill\_down),</code><br><code> \_Result, IF(\_Transpose, TRANSPOSE(\_Words), \_Words),</code><br><code> \_Result</code><br><code>));</code> |
| Content Type | ExcelFormula |
| Location | <code>Words</code> |

[^Top](#oa-robot-definitions)
