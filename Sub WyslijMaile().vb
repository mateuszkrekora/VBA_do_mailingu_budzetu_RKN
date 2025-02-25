Sub WyslijMaile()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim wsData As Worksheet, wsText As Worksheet
    Dim i As Integer, lastRow As Integer
    Dim emailAddress As String, kwota As String, bodyText As String, emailSubject As String
    Dim attachmentPath As String
    Dim wbPath As String
    Dim FileDialog As FileDialog
    
    ' Ustawienie referencji do arkuszy
    Set wsData = ThisWorkbook.Sheets("maile i kwota")
    Set wsText = ThisWorkbook.Sheets("text")
    
    ' Pobranie tytulu maila z komórki A3
    emailSubject = wsText.Range("A3").Value
    If Trim(emailSubject) = "" Then
        MsgBox "Tytul wiadomosci w A3 jest pusty!", vbExclamation
        Exit Sub
    End If
    
    ' Pobranie tresci wiadomosci (zakres H5:H16)
    Dim cell As Range
    bodyText = ""
    For Each cell In wsText.Range("H5:H16")
        bodyText = bodyText & cell.Value & vbNewLine
    Next cell
    
    If Trim(bodyText) = "" Then
        MsgBox "Tresc wiadomosci w H5:H16 jest pusta!", vbExclamation
        Exit Sub
    End If
    
    ' Pobranie sciezki do zalacznika
    attachmentPath = "LINK DO SCIEZKI"
    
    ' Sprawdzenie czy plik istnieje, jesli nie, otworzenie okna wyboru
    If Dir(attachmentPath) = "" Then
        MsgBox "Nie znaleziono pliku: " & attachmentPath & vbNewLine & "Wybierz plik recznie.", vbExclamation
        Set FileDialog = Application.FileDialog(msoFileDialogFilePicker)
        
        With FileDialog
            .Title = "Wybierz plik do zalaczenia"
            .Filters.Clear
            .Filters.Add "Pliki Excel", "*.xlsx;*.xlsm"
            If .Show = -1 Then
                attachmentPath = .SelectedItems(1)
            Else
                MsgBox "Nie wybrano pliku. Przerwano wysylanie.", vbExclamation
                Exit Sub
            End If
        End With
    End If
    
    ' Tworzenie obiektu Outlook
    On Error Resume Next
    Set OutApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    
    If OutApp Is Nothing Then
        MsgBox "Nie mozna uruchomic Outlooka. Sprawdz, czy jest poprawnie zainstalowany.", vbExclamation
        Exit Sub
    End If
    
    ' Znalezienie ostatniego wiersza w arkuszu
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    
    ' Petla przez wszystkie wiersze z danymi
    For i = 2 To lastRow
        emailAddress = wsData.Cells(i, 3).Value ' Kolumna C: e-mail
        kwota = wsData.Cells(i, 4).Value ' Kolumna D: kwota
        
        ' Sprawdzenie czy e-mail nie jest pusty
        If emailAddress <> "" Then
            ' Tworzenie nowej wiadomosci
            Set OutMail = OutApp.CreateItem(0)
            
            With OutMail
                .To = emailAddress
                .Subject = emailSubject
                .Body = Replace(bodyText, "XXX", kwota)
                
                ' Dolaczanie zalacznika
                .Attachments.Add attachmentPath
                
                .Display ' Otwieranie okna wiadomosci zamiast wysylania
            End With
            
            ' Zwolnienie obiektu wiadomosci
            Set OutMail = Nothing
        End If
    Next i
    
    ' Zwolnienie obiektu Outlooka
    Set OutApp = Nothing
    
    MsgBox "Maile zostaly przygotowane do wyslania!", vbInformation
End Sub

