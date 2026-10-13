# Szablon spotkania – add-in do Outlooka

Przy każdym nowym spotkaniu add-in automatycznie wstawia na początek opisu:

```
Agenda:
1.
2.
3.

Link do logowania czasu:
```

Link po dwukropku wkleja sam użytkownik.

Na wstążce nowego spotkania jest też przycisk **Wstaw szablon**, gdyby ktoś szablon usunął. Jeśli szablon jest już w opisie, add-in nie wstawi go drugi raz.

## Zawartość

| Plik | Do czego służy |
|---|---|
| `manifest.xml` | Manifest wgrywany w centrum administracyjnym M365 |
| `launchevent.js` | Logika i konfiguracja szablonu |
| `commands.html` | Strona ładująca skrypt (nowy Outlook / web) |
| `assets/` | Ikony |

## 1. Konfiguracja

Nic nie trzeba zmieniać. Opcjonalnie w `launchevent.js`, w bloku `CONFIG` na górze pliku:

- `agendaItems` – liczba punktów agendy (domyślnie 3)
- `timeLogLabel` – treść napisu pod agendą

## 2. Hosting (HTTPS)

Pliki `commands.html`, `launchevent.js` i folder `assets/` muszą być dostępne publicznie przez HTTPS. Najprościej:

- **GitHub Pages**: wrzuć folder do repozytorium i włącz Pages. Adres będzie wyglądał tak: `https://<konto>.github.io/<repo>`.
- **Azure Static Web Apps** albo dowolny serwer HTTPS też się nadadzą.

Następnie w `manifest.xml` zamień każde `https://YOUR-HOST` na swój adres, bez ukośnika na końcu.

## 3. Test na swoim koncie (opcjonalnie)

1. Otwórz https://aka.ms/olksideload. Zobaczysz okno „Dodatki dla programu Outlook”.
2. Wybierz **Moje dodatki** → **Dodaj dodatek niestandardowy** → **Dodaj z pliku** → `manifest.xml`.
3. Utwórz nowe spotkanie w nowym Outlooku. Szablon powinien pojawić się w opisie.

## 4. Wdrożenie dla całej firmy

1. Wejdź na https://admin.microsoft.com → **Ustawienia** → **Zintegrowane aplikacje** → **Przekaż aplikacje niestandardowe**.
2. Wybierz typ „Dodatek pakietu Office”, wgraj `manifest.xml`.
3. Przypisz użytkowników: wszystkich, wybrane grupy lub tylko siebie.
4. Propagacja trwa zwykle kilka godzin, maksymalnie do 24 h.

Przy wdrożeniu przez admina użytkownicy nie muszą niczego instalować ani akceptować.

## Zmiany w przyszłości

- **Zmiana treści szablonu**: wystarczy podmienić `launchevent.js` na serwerze. Nie trzeba ponownie wgrywać manifestu, ale nowy Outlook może trzymać starą wersję w cache przez jakiś czas.
- **Zmiana manifestu** (adres, nazwa, ikony): podbij `<Version>` w `manifest.xml` i zaktualizuj aplikację w Zintegrowanych aplikacjach.

## Uwagi

- Działa w nowym Outlooku, Outlooku w przeglądarce, klasycznym Outlooku dla Windows (wersja 2108+) i na Macu.
- Jeśli macie włączone automatyczne dodawanie Teams do spotkań, link Teams zostaje, a szablon trafia nad niego.
