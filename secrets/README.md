# secrets/

`master-link.json` = the locked copy of this app's private links (/exec link,
deployment ID, Sheet ID, Apps Script project ID), made on `…/exec?page=master`
and locked with the **Master Key** (PBKDF2-SHA256 150k → AES-GCM 256).
It is safe to keep here: without the Master Key it is unreadable.

Copy it from Drive ("Master Link.json", next to the LOTTERY JACKPOT Sheet)
each time the Master Key is set or changed.
