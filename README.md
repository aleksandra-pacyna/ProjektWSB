# Projekt WSB
Program służy do tworzenia plików wsadowych do systemu SAP na podstawie plików z poprzedniego okresu i pliku Excela z danymi z obecnego miesiąca. Użytkownik programu podaje ścieżkę, z której kopiuje pliki tekstowe z poprzedniego miesiąca. Następnie tworzony jest nowy folder, do którego te pliki są kopiowane. Nazwy plików zostają zmienione zgodnie z obowiązujacą datą, ponieważ poprzednie mają w nazwie poprzedni miesiąc. Przy użyciu biblioteki regex zostaje sprawdzona zawartość plików tekstowych- sprawdzana jest odpowiednia struktura. Do programu zostaje wykorzystany plik Excela skąd pobierane są dane dla nowego okresu, konieczne jest znalezienie odpowiedniego wiersza z danymi na podstawie nazwy wczytanego wcześniej pliku tekstowego. Ponownie przy użyciu biblioteki regex zostaje zmieniona kwota w nowym pliku zgodnie ze strukturą na odpowiedniej pozycji w pliku.
Fragemnt przykładowego pliku z danymi:
> 110,20230417,1234,1234,0,"1234","Nr konta","Adres","ZUS|||",0,1234,"1234|R1234|S1234","","","1"
# Wykorzystane biblioteki
- *datetime* - biblioteka wykorzystana do operacji na datach,
- *os* - działanie na plikach i folderach,
- *regex* - sprawdzenie czy plik tekstowy posiada odpowiednią strukturę, wpisanie danych w określonym miejscu pliku,
- *openpyxl* - działania w pliku excela (czytanie danych z kolumny w określonym wierszu)
