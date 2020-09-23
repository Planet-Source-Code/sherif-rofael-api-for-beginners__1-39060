Attribute VB_Name = "disk_infos"
Option Explicit
Public thedrivechoosen As String
Public diskcapacity As String
Public freediskspace As String
Dim cstart
Public Declare Function GetDiskFreeSpaceEx Lib "kernel32" _
    Alias "GetDiskFreeSpaceExA" (ByVal lpDirectoryName As String, _
    lpFreeBytesAvailableToCaller As Currency, _
    lpTotalNumberOfBytes As Currency, _
    lpTotalNumberOfFreeBytes As Currency) As Long
    Dim Status As Long
    Dim TotalBytes, TempTotalBytes As Currency
    Dim FreeBytes, TempFreeBytes As Currency
    Dim BytesAvailableToCaller As Currency
    Dim TotalSpaceValue, FreeSpaceValue As String
 
