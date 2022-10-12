Imports System.Environment
Imports System.Environment.SpecialFolder

Friend Module Constant
    Friend Const XL_NAME = "excel"
    Friend ReadOnly FILE_SETUP_NAME = $"{My.Resources.app_name} Setup.msi"
    Friend ReadOnly BACK_PATH = GetFolderPath(ApplicationData)
    Friend ReadOnly FRNT_PATH = $"{BACK_PATH}\{My.Resources.co_name}"
    Friend ReadOnly FILE_SETUP_ADR = $"{FRNT_PATH}\{FILE_SETUP_NAME}"
End Module
