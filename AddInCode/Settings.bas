Attribute VB_Name = "Settings"
Option Explicit

'*************Constants, usually not changed
Public Const c_noPWrequired As String = "%nopw%"                'String returned by some stubs, indicating that no password is required to protext/unprotect sheets
Public Const c_noBookPWrequired As String = "%nopw%"            'String returned by some stubs, indicating that no password is required to protext/unprotect the workBook

Public Const c_firstLanguageName As String = "English"          'Name of the first language in the NLS table editor,; used to find the the start of the language columns
Public Const c_useFirstLanguageIfBlank As Boolean = True        'Flag to control if the first language result should replace a blank entry in another language
Public Const c_mainLanguage As Long = 0                         'Defines which language will be displayed as default language in the NLS table editor (0 based)
Public Const c_moduleForShapes As String = "Button"             'Module name for shape text
Public Const c_moduleForMenus As String = "ContextMenu"         'Module name for Menu elements
Public Const c_parmSep As String = "°°"                         'String used to separate parameters on a call where a parameter string is passed as module info
Public Const c_defaultColWidths As String = "40;85;100;60;40"   'Sets the column widths for the columns in the NLS table editor - language column width is the remainder


'*************Default Constants
Public Const c_NlsSheetName As String = "NLS"                   'Names for the system sheets
Public Const c_infoSheetName As String = "%info%"
Public Const c_sortSheetname As String = "%sort%"

Public Const c_NLSCellFunctions As String = "GetNlsText,IsYes,IsNo"
