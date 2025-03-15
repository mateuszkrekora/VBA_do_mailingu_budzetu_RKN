Sub GenerujMaile()
    Dim OutlookApp As Object
    Dim MailItem As Object
    Dim Wb As Workbook
    Dim WsDane As Worksheet, WsTresc As Worksheet
    Dim LastRow As Long
    Dim i As Integer
    Dim EmailAddr As String, Zwrot As String, Plec As String, Tresc As String, Tytul As String
    Dim AttachmentPath As String
    
    ' Sciezka do zalacznika
    AttachmentPath = "TWOJA/SCIEZKA/DO/PLIKU"
    
    ' Ustawienie referencji do Excela
    Set Wb = ThisWorkbook
    Set WsDane = Wb.Sheets("dane")
    Set WsTresc = Wb.Sheets("tresc")
    
    ' Pobranie tytulu maila
    Tytul = WsTresc.Cells(1, 2).Value
    
    ' Pobranie pelnej tresci maila jako jeden ciag znaków
    Tresc = Join(Application.Transpose(WsTresc.Range(WsTresc.Cells(2, 2), WsTresc.Cells(WsTresc.Cells(Rows.Count, 2).End(xlUp).Row, 2))), vbNewLine & vbNewLine)
    
    ' Znalezienie ostatniego wiersza w arkuszu "dane"
    LastRow = WsDane.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Uruchomienie Outlooka
    On Error Resume Next
    Set OutlookApp = GetObject(, "Outlook.Application")
    If OutlookApp Is Nothing Then Set OutlookApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    
    ' Iteracja przez dane
    For i = 2 To LastRow
        EmailAddr = WsDane.Cells(i, 1).Value
        Zwrot = WsDane.Cells(i, 3).Value
        Plec = WsDane.Cells(i, 4).Value
        
        ' Usuniecie podwójnego przecinka, jesli wystepuje
        If Right(Zwrot, 1) = "," Then Zwrot = Left(Zwrot, Len(Zwrot) - 1)
        
        ' Tworzenie nowej wiadomosci
        Set MailItem = OutlookApp.CreateItem(0)
        With MailItem
            .To = EmailAddr
            .Subject = Tytul
            .Body = Replace(Replace(Tresc, "<Szanowny/a>", Zwrot & ","), "<plec>", Plec)
            .Attachments.Add AttachmentPath ' Dodanie zalacznika
            .SentOnBehalfOfName = "twojmail@mail.com" ' Wymuszenie nadawcy
            .Display ' Wyswietla wiadomosc zamiast wysylac
        End With
        
        ' Zwolnienie obiektu maila
        Set MailItem = Nothing
    Next i
    
    ' Zwolnienie obiektu Outlooka
    Set OutlookApp = Nothing
    
    MsgBox "Maile zostaly wygenerowane i sa gotowe do wysylki!", vbInformation
End Sub

