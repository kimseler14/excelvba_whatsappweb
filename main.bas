Option Explicit
'Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Sub whatsapp()

Dim saveLocation As String
Dim rng As Range
Dim filename As String
Dim filename2 As String
Dim strflm As String
Dim sayı As Integer
Application.DisplayAlerts = False



Dim bot As New Selenium.ChromeDriver

saveLocation = ActiveWorkbook.Path & "\"

For sayı = 1 To 850
    Range("M3").Select
    ActiveCell.FormulaR1C1 = sayı
    filename = Range("B3").Text
    
    
    If filename = "#YOK" Then
        Do
            sayı = sayı + 1
            Range("M3").Select
            ActiveCell.FormulaR1C1 = sayı
            filename = Range("B3").Text
        Loop Until filename <> "#YOK"
            
    End If
    filename2 = Range("A2").Text
    strflm = filename2 & " " & filename & ".pdf"
    
    Set rng = Range("A1:D22")

    rng.ExportAsFixedFormat Type:=xlTypePDF, _
    filename:=saveLocation & strflm, IgnorePrintAreas:=True, Quality:=xlQualityStandard

    With bot
    .AddArgument "--disable-plugins-discovery"
    .AddArgument "--disable-extensions"
    .AddArgument "--disable-infobars"
    .AddArgument "--disable-popup-blocking"
    
    
    .SetProfile ("C:\Users\LENOVO\AppData\Local\SeleniumBasic")
    
    .Get "https://web.whatsapp.com/send?phone=905337480363&text=Sayın+" + filename + " haber kağıdınız."
    .Timeouts.ImplicitWait = 600000

    
    On Error Resume Next
    Do While .FindElementByClass("_35EW6") Is Nothing
        DoEvents
    Loop
    
    On Error GoTo 0
    .FindElementByClass("_35EW6").Click
    Sleep (3000)
    .FindElementByXPath("//div[@title = 'Ekle']").Click
    Sleep (3000)
    .FindElementByXPath("//input[@accept='image/*,video/mp4,video/3gpp,video/quicktime']").SendKeys (saveLocation + strflm)
    Sleep (3000)
    .FindElementByXPath("//span[@data-icon='send-light']").Click
    Sleep (5000)

    
    End With
Next sayı
End Sub




