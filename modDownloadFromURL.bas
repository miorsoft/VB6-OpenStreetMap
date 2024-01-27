Attribute VB_Name = "modDownloadfromURL"
Option Explicit
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
                                           (ByVal pCaller As Long, _
                                            ByVal szURL As String, _
                                            ByVal szFileName As String, _
                                            ByVal dwReserved As Long, _
                                            ByVal lpfnCB As Long) As Long

Private Const ERROR_SUCCESS As Long = 0
Private Const BINDF_GETNEWESTVERSION As Long = &H10
'Private Const INTERNET_FLAG_RELOAD As Long = &H80000000

Public Function DownloadFile(sSourceUrl As String, sLocalFile As String) As Boolean
    '
    ' Download the file. BINDF_GETNEWESTVERSION forces
    ' the API to download from the specified source.
    ' Passing 0& as dwReserved causes the locally-cached
    ' copy to be downloaded, if available. If the API
    ' returns ERROR_SUCCESS (0), DownloadFile returns True.
    '
    DownloadFile = URLDownloadToFile(0&, sSourceUrl, sLocalFile, BINDF_GETNEWESTVERSION, 0&) = ERROR_SUCCESS
End Function
