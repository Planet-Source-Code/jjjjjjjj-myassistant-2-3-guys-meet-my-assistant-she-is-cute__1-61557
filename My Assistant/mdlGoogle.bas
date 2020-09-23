Attribute VB_Name = "mdlGoogle"
Option Explicit

Public Enum SearchTypeEnum
    [Srch_Web] = 1
    [Srch_Image] = 2
    [Srch_Group] = 3
    [Srch_News] = 4
End Enum


Public Function GetGoogleSearchString(ByVal SearchText As String, _
                            ByVal NumResults As Long, _
                            ByVal SearchType As SearchTypeEnum, _
                            ByVal ExactPhrase As Boolean)
Dim SearchID As String
Dim ExactString As String

    Select Case SearchType
        Case Srch_Web
            SearchID = "search"
        Case Srch_Image
            SearchID = "images"
        Case Srch_News
            SearchID = "news"
        Case Srch_Group
            SearchID = "groups"
    End Select

    If ExactPhrase Then ExactString = "%22" Else ExactString = vbNullString
    GetGoogleSearchString = "http://www.google.com/" & SearchID & "?num=" & NumResults & "&hl=en&lr=&ie=ISO-8859-1&as_qdr=all&q=" & ExactString & Replace(SearchText, " ", "+") & ExactString
  
End Function

Public Function GetGoolgeAdvancedSearchPage(ByVal SearchText As String)
    GetGoolgeAdvancedSearchPage = "http://www.google.com/advanced_search?q=" & Replace(SearchText, " ", "+") & "&hl=en&lr=&ie=UTF-8"
End Function
