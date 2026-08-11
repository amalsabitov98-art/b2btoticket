# Pax Data Converter

Converts the simple passenger list partners send over (`№, Full Name, Date of
Birth, Passport №, Expiry, Room`) into the "Pax Data" upload format
(`Sequence, Traveling With, Pax Type, Title, First Name, Last Name, Gender,
DOB, Nationality, Mobile Number, Email, ID Number, Passport Number, Passport
Expiry, Passport Issued, Passport Issued Country, Passport Nationality`).

## Usage

```bash
python3 convert.py INPUT.xlsx [-o OUTPUT.xlsx] [--year YYYY] [--date DD.MM.YYYY]
```

- `INPUT.xlsx` — the partner's file. The script auto-detects the header row
  and the columns `Full Name`, `Date of Birth`, `Passport №`, `Expiry`
  (`Room` is read but not carried over, since it has no equivalent column in
  Pax Data).
- `-o/--output` — output path. Defaults to `<input>_pax_data.xlsx` next to
  the input file.
- `--date DD.MM.YYYY` — explicit travel date used to compute each
  passenger's age (and therefore Adult/Child/Infant). If omitted, the script
  looks for a `DD.MM` pattern in the input filename (e.g. `TASBUS_14.08.xlsx`
  → `14.08`) and combines it with `--year` (defaults to the current year).

Example:

```bash
python3 convert.py TASBUS_14.08.xlsx --year 2026
# -> Converted 61 passengers -> TASBUS_14.08_pax_data.xlsx
```

## Derivation rules

| Pax Data field | How it's derived |
|---|---|
| Sequence | Row order (1..N) |
| Pax Type | Age at travel date: < 2 → `INF`, 2–11 → `CHD`, 12+ → `ADT` |
| Title | `mr`/`mrs` for ADT, `mstr`/`miss` for CHD/INF, based on Gender |
| First Name / Last Name | Input `Full Name` is `SURNAME GIVEN NAME` → swapped to `First Last` |
| Gender | Heuristic: Female if the first *or* last name ends in `-a`, else Male |
| Nationality / Passport Issued Country / Passport Nationality | Hardcoded `Uzbekistan` |
| Passport Number / Expiry / DOB | Copied as-is |
| Traveling With, Mobile Number, Email, ID Number, Passport Issued | Left blank (not present in the source list) |

**Gender is a heuristic and not 100% reliable** — surnames that don't
decline by gender (e.g. a whole family sharing one surname) are
cross-checked against the first name, but rare names that break the
"ends in -a = female" pattern (e.g. a male named "Muradulla") will be
misclassified. Review the `Gender`/`Title` columns before sending the file
onward.

## Running the tests

```bash
python3 -m unittest tools/pax_data_converter/test_convert.py
```
