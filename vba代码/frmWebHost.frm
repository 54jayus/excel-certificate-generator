VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWebHost 
   Caption         =   "Copyright (c) 2026 Rongny"
   ClientHeight    =   10320
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16845
   OleObjectBlob   =   "frmWebHost.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "frmWebHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Click()

End Sub

Private Sub wbMain_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
      Dim cmd As String
      Dim requestId As String
      Dim resultText As String
      Dim statusText As String

      cmd = WebHost.ParseCommand(CStr(URL))
      If cmd = "" Then Exit Sub

      requestId = WebHost.ParseRequestId(CStr(URL))
      Cancel = True

      resultText = WebHost.HandleBrowserCommand(CStr(URL))

      Select Case LCase$(cmd)
          Case "ping"
          statusText = "VBA responded"
      Case "pickrange"
          statusText = "Range selected"
      Case "getinitdata"
          statusText = "Init data loaded"
      Case "getsheetfields"
          statusText = "Fields loaded"
      Case "validaterange"
          statusText = "Range checked"
      Case "generatetemplate"
          statusText = "Template generated"
      Case "writedata"
          statusText = "Data written"
      Case Else
          statusText = "Command handled"
      End Select

      WebHost.PushResultToBrowser requestId, resultText, statusText
End Sub

