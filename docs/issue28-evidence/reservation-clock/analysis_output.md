## (a) Uploaded minus file time
SA4#138 HEAD sample n=99: min=-79988s median=3s max=213s; within 0..60 s: 69 ; within +-2 h of 1h/2h/-6h: {1: 0, 2: 0, -6: 0}
sa4_138: listing n=596 uploaded-listing in [0,60)s: 355; by nearest hour: {-223: 1, -53: 2, -32: 1, -29: 1, -24: 2, -23: 1, -22: 1, -8: 1, -5: 1, -2: 2, -1: 62, 0: 521}
TSGS4_136_Montreal: listing n=497 uploaded-listing in [0,60)s: 415; by nearest hour: {-125: 1, -18: 1, 0: 495}
TSGS4_135_India: listing n=469 uploaded-listing in [0,60)s: 247; by nearest hour: {-94: 1, -78: 1, -1: 1, 0: 466}
TSGS4_137-e: listing n=470 uploaded-listing in [0,60)s: 405; by nearest hour: {0: 470}

## (b) upload - reservation (minutes) under each hypothesis
| meeting | hypothesis | n | inversions (res after upload) | min gap | 1st pct | gaps 0-10 min | gaps 10-60 min | gaps 60-130 min |
|---|---|---|---|---|---|---|---|---|
| sa4_138 | H1 UTC | 596 | 0 | 0.8 | 1.5 | 91 | 62 | 39 |
| sa4_138 | H2 Europe/Paris | 596 | 0 | 60.8 | 61.5 | 0 | 0 | 159 |
| sa4_138 | H3 Calgary (UTC-6 fixed; Alberta ABT from 2026-11-01, MDT before) | 596 | 282 | -359.2 | -358.5 | 0 | 8 | 7 |
| TSGS4_136_Montreal | H1 UTC | 497 | 0 | 1.1 | 2.0 | 69 | 115 | 57 |
| TSGS4_136_Montreal | H2 Europe/Paris | 497 | 0 | 121.1 | 122.0 | 0 | 0 | 69 |
| TSGS4_136_Montreal | H3 Montreal (EDT UTC-4) | 497 | 272 | -238.9 | -237.9 | 2 | 12 | 7 |
| TSGS4_135_India | H1 UTC | 469 | 0 | 0.8 | 2.0 | 48 | 90 | 39 |
| TSGS4_135_India | H2 Europe/Paris | 469 | 0 | 60.8 | 62.0 | 0 | 0 | 145 |
| TSGS4_135_India | H3 India (IST UTC+5:30) | 469 | 0 | 330.8 | 332.0 | 0 | 0 | 0 |
| TSGS4_137-e | H1 UTC | 471 | 0 | 1.5 | 2.1 | 66 | 100 | 43 |
| TSGS4_137-e | H2 Europe/Paris | 471 | 0 | 121.5 | 122.1 | 0 | 0 | 66 |

Raw (H1) gaps < 180 min histogram, 10-min bins, all meetings:
      0..10   min: 273
     10..20   min: 153
     20..30   min: 56
     30..40   min: 65
     40..50   min: 50
     50..60   min: 44
     60..70   min: 25
     70..80   min: 41
     80..90   min: 36
     90..100  min: 27
    100..110  min: 10
    110..120  min: 19
    120..130  min: 20
    130..140  min: 11
    140..150  min: 19
    150..160  min: 18
    160..170  min: 9
    170..180  min: 14

## (d) raw gaps (H1 reading) split by Paris DST state at reservation
sa4_138 CEST (UTC+2): n=28 min=1.6 n<60min=19 n<120min=19 n<10min=18
sa4_138 CET (UTC+1): n=568 min=0.8 n<60min=134 n<120min=169 n<10min=73
TSGS4_136_Montreal CEST (UTC+2): n=497 min=1.1 n<60min=184 n<120min=235 n<10min=69
TSGS4_135_India CET (UTC+1): n=469 min=0.8 n<60min=138 n<120min=172 n<10min=48
TSGS4_137-e CEST (UTC+2): n=471 min=1.5 n<60min=166 n<120min=204 n<10min=66

## (c) SA4#138 reservations 2026-11-16..20 by raw hour-of-day (as written)
| raw hour | reservations | uploads | = Calgary hour if UTC | = Calgary hour if Paris |
|---|---|---|---|---|
| 00 | 11 | 1 | 18 | 17 |
| 01 | 7 | 2 | 19 | 18 |
| 02 | 2 | 1 | 20 | 19 |
| 03 | 23 | 2 | 21 | 20 |
| 04 | 3 | 7 | 22 | 21 |
| 05 | 1 | 1 | 23 | 22 |
| 06 | 1 | 0 | 00 | 23 |
| 07 | 2 | 0 | 01 | 00 |
| 08 | 0 | 0 | 02 | 01 |
| 09 | 4 | 0 | 03 | 02 |
| 10 | 1 | 0 | 04 | 03 |
| 11 | 6 | 3 | 05 | 04 |
| 12 | 2 | 1 | 06 | 05 |
| 13 | 3 | 0 | 07 | 06 |
| 14 | 21 | 49 | 08 | 07 |
| 15 | 18 | 5 | 09 | 08 |
| 16 | 15 | 8 | 10 | 09 |
| 17 | 27 | 8 | 11 | 10 |
| 18 | 25 | 67 | 12 | 11 |
| 19 | 10 | 5 | 13 | 12 |
| 20 | 4 | 8 | 14 | 13 |
| 21 | 11 | 6 | 15 | 14 |
| 22 | 12 | 40 | 16 | 15 |
| 23 | 10 | 3 | 17 | 16 |
