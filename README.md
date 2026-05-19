# Telegram Godziny Bot

Bot zapisuje wejscia do pracy, przerwy, koniec dnia, zarobki oraz eksportuje historie do Excela.

## Uruchomienie

1. Otworz plik `.env`.
2. Wklej token z BotFather:

```env
TELEGRAM_BOT_TOKEN=123456:ABC...
HOURLY_RATE=31.5
TIMEZONE=Europe/Warsaw
```

3. Uruchom:

```powershell
.\run.ps1
```

## Komendy

- `/start` - rozpoczyna prace albo wraca z przerwy.
- `/przerwa` - rozpoczyna przerwe.
- `/koniec` - konczy dzien i pokazuje zarobki.
- `/status` - pokazuje aktualny stan i podsumowanie.
- `/uzupelnij` - dodaje zapomniana zmiane jednym zdaniem; po komendzie wpisz np. `wczoraj 8-16 przerwa 12-12:30`.
- `/uzupelnij wczoraj 8-16` - dodaje zapomniana zmiane od razu.
- `/anuluj` - anuluje dodawanie krok po kroku.
- `/dodaj 8-16` - szybko dodaje dzisiejsza zmiane.
- `/dodaj wczoraj 8:00 16:00` - szybko dodaje wczorajsza zmiane.
- `/dodaj 11.05 8:00-16:00 przerwa 12-12:30` - dodaje zmiane z przerwa.
- `/lista` - pokazuje 5 ostatnich wpisow albo mniej, jesli tyle jeszcze nie ma.
- `/usun` - pokazuje ostatnie wpisy z numerami.
- `/usun NUMER` - usuwa konkretny wpis wedlug numeru z listy, np. `/usun 3`.
- `/usun ostatni` - usuwa najnowszy wpis.
- `/raport` - pokazuje podsumowanie biezacego miesiaca.
- `/raport YYYY-MM` - pokazuje podsumowanie wybranego miesiaca, np. `/raport 2026-05`.
- `/stawka` - pokazuje aktualna stawke.
- `/stawka 31,50` - zmienia stawke dla kolejnych zmian oraz aktywnej zmiany.
- `/wyplata` - zeruje kwote "od ostatniej wyplaty" przez oznaczenie zakonczonych zmian jako rozliczone.
- `/excel` - wysyla plik `.xlsx` z historia i arkuszem podsumowania.
- `/ksiegowa` - wysyla prosty Excel dla ksiegowej za biezacy miesiac: data, start, koniec, liczba godzin.
- `/ksiegowa YYYY-MM` - wysyla prosty Excel dla ksiegowej za wybrany miesiac, np. `/ksiegowa 2026-05`.
- `/popraw` - reczna edycja godzin.
- `/pomoc` - pokazuje pomoc.

## Reczna edycja

Przyklady:

```text
/popraw lista
/popraw lista 2026-05-11
/popraw dodaj 2026-05-11 08:00 16:00
/popraw start 3 07:45
/popraw koniec 3 15:30
/popraw koniec 3 brak
/popraw usun 3
/popraw przerwa 3 dodaj 12:00 12:30
/popraw przerwa 3 usun 2
```

Numery wpisow sa liczone wedlug kolejnosci historii. Po usunieciu wpisu numeracja sama sie przesuwa, wiec nie ma dziur.
Godziny bez daty sa interpretowane w dacie danej zmiany.

## Zasada przerwy

Jesli wpiszesz `/przerwa`, a tego samego dnia nie wpiszesz ponownie `/start`, bot przy nastepnej komendzie automatycznie uzna poczatek przerwy za koniec pracy.
