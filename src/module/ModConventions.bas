Attribute VB_Name = "ModConventions"
' Name conventions to be used in this project
'
' Naming of variables
'
' Make public names long and descriptive, keep local names short
'
' Types
' str: String
' bln: Boolean
' int: Integer
' dbl: Double
' obj: Object
'
' Classes:
' sht: Sheet
' wkb: Workbook
' mod: Module
'
' Constants ALLCAPITALS_WITH_UNDERSCORE
'
' Public variable start with capital
' followed by type
' followed by description
'
' Examples:
' StrName
' modGlobal
'
' Private variables start with camelcase
' followed by type
' followed by description
'
' Examples:
' strName
'
' When using a function/variable from another module, always prefix with module name
' to avoid namespace conflicts
'
' Sheet is prefixed with sht
' followed by type
' Gui: Graphical User Interface
' Ber: Calculation sheet
' Tbl: Table sheet
' Pat: Contains patient data
' Div: Is a divider sheet
