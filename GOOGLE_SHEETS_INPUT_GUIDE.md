# Google Sheets Input Guide

This guide explains the text formats the parser accepts from the Google Sheets case/archive rows.

## General Rules

- Use `;` to separate multiple `sets`, `sets_returned`, and `bonegraft` values.
- Use `;` or `/` to separate multiple `plates`.
- Use spaces, `,`, `/`, or `;` to separate multiple `powertools`.
- Use `;` or `,` to separate multiple `extra_items`.
- Upper/lower case is normalized, but uppercase is easier to read and review.
- For sets, exact UID is always safer than shorthand.

## Sets

Preferred input:

- Exact UID when you know the physical set: `PFN09`, `REAM02`, `2.712`
- Multiple exact UIDs: `PFN09;REAM02`

Accepted shorthand/category input:

- `1.5`, `2.0`, `2.0ONLY`, `2.4`, `2.7`, `3.5`, `ANKLE`, `C2.5`, `C3.5`, `C4.0`, `C5.2`, `COATL`, `FIBN`, `FNS`, `FOOT`, `FOOTI`, `ILNF`, `ILNH`, `ILNRU`, `ILNT`, `LPFN`, `P5400`, `P5503`, `P55031006`, `P8400`, `PFN`, `PFNIIIMP`, `PFNIISYS`, `PFNR`, `REAM`, `RFN`, `ROI`, `SPFNII`, `STDC2.4`, `STDC3.0`, `STDC4.0`, `STDC6.5`, `SUPSUB`, `TENS`

Notes:

- Shorthand like `PFN` or `2.7` helps the system classify the request, but it does not identify one unique physical set.
- If a case should count as a delivered UID set, enter the exact UID, not only the shorthand.

## Plates

Enter plates using the plate UID, plus an optional suffix.

Base examples:

- `PHILOS` = standard range
- `PLTI` = standard range
- `TSP` = standard range

Range suffix rules:

- `-L` = standard + long
- `-EL` = standard + long + extra long
- `-S` = standard + short
- `-LONLY` = long only
- `-ELONLY` or `-XLONLY` = extra long only
- `-SONLY` = short only
- `*` = from stock

Examples:

- `PHILOS` = standard PHILOS
- `PHILOS-L` = standard and long PHILOS
- `PHILOS-EL` = standard, long, and extra long PHILOS
- `PHILOS-LONLY` = only long PHILOS
- `PHILOS-ELONLY` = only extra long PHILOS
- `PHILOS-XLONLY` = same as `PHILOS-ELONLY`
- `PHILOS*` = standard PHILOS from stock
- `PHILOS-ELONLY*` = extra long PHILOS from stock

Important:

- If only extra long PHILOS is sent, enter `PHILOS-ELONLY` or `PHILOS-XLONLY`.
- Do not enter `PHILOS-EL` for that case. `PHILOS-EL` means standard + long + extra long together.

Supported plate UIDs:

- `ADT`, `APP`, `CCOMBO`, `CHOOK`, `CMESH`, `DFIBII`, `DFIBIII`, `DIA`, `DLF`, `DLHI`, `DLHII`, `DLT`, `DLTII`, `DMH`, `DMT`, `DPLH`, `DPT`, `DRVL`, `DSC`, `FIBHOOK`, `FSP`, `METAI`, `METAII`, `MSC`, `OLEI`, `OLEII`, `PFP`, `PHILOS`, `PLTI`, `PLTII`, `PMT`, `PPMT`, `RECON`, `TSP`, `TUBULAR`, `URS`

## Powertools

Accepted formats:

- Full UID: `P55030015`
- Category shorthand: `P5503`, `P5400`, `P8400`
- Full UID without leading `P` also normalizes: `55030015`

Examples:

- `P55030015`
- `P55030015 P5400`
- `P55030015/P8400`

## Bonegraft

You can enter either the ref code or the shorthand.

Current shorthand mappings:

- `n1` = `PAS1` = Reprobone novo 1 cc
- `n2.5` = `PAS2.5` = Reprobone novo 2.5 cc
- `f1` = `GNP1` = Reprobone fusion 1 cc
- `f2.5` = `GNP2.5` = Reprobone fusion 2.5 cc
- `G5` = `RBG5` = Reprobone 5 cc
- `R10` = `RBG10` = Reprobone 10 cc
- `r20` = `RBG20` = Reprobone 20 cc
- `r30` = `RBG30` = Reprobone 30 cc
- `b10` = `RB111` = Reprobone 10x10x10 mm
- `b20` = `RB222` = Reprobone 20x20x20 mm
- `c5` = Collatamp G 5 cm x 5 cm
- `c10` = Collatamp G 10 cm x 10 cm

Example:

- `f1;b10`

## Extra Items

`extra_items` is free text. The parser keeps what you type.

Examples:

- `DRV-L`
- `DRN-L`
- `DRV-L;DRN-L`

## Recommended Sheet Examples

- `sets`: `PFN09;REAM02`
- `sets_returned`: `REAM02`
- `plates`: `PHILOS-ELONLY;TSP-L`
- `powertools`: `P55030015 P5400`
- `bonegraft`: `f1;b10`
- `extra_items`: `DRV-L;DRN-L`

## Quick Reference

- Need a unique physical set: use exact set UID like `PFN09`
- Need extra long PHILOS only: use `PHILOS-ELONLY`
- Need PHILOS standard + long + extra long: use `PHILOS-EL`
- Need stock plate instead of drawer plate: add `*`
