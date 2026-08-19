# Installing tANk

tANk adds an AI chat pane to Excel, Word and PowerPoint. It runs on your own free API key, so
there is nothing to pay for and no account to create.

Works on Windows with desktop Office (Microsoft 365, or Office 2019 and later).

## Install

1. Close Excel, Word and PowerPoint.
2. Download `tANk-Setup.exe` from the latest release.
3. Run it. Windows shows a blue warning first, see below. No administrator password is needed;
   it installs for you only.
4. Open Excel. On the **Home** tab choose **Add-ins**, then pick **tANk** under Developer Add-ins.

![The Add-ins menu on the Home tab, with tANk listed under Developer Add-ins](docs/where-to-find-tank.png)

The pane opens on the right and stays open while you work, so this is a once-per-session step.

## Windows will warn you

The first time you run the installer, Windows puts up a blue "Windows protected your PC" screen.
Click **More info**, then **Run anyway**.

The warning is there because the installer is not code signed. Signing means buying a certificate
and renewing it every year, and this is a free project, so it does not have one. Windows shows
that screen for any unsigned program, so it tells you nothing about what is inside this one.

If you would rather not run an unsigned exe, use the script instead. It downloads the manifest,
writes the registry value and clears Office's add-in cache, which is everything the installer does.

```powershell
irm https://anjan-NZ.github.io/tANk/install.ps1 -OutFile install.ps1
powershell -ExecutionPolicy Bypass -File install.ps1
```

To check that the exe is the one built from this repository, and not something swapped in later,
verify the provenance GitHub records at build time:

```
gh attestation verify tANk-Setup.exe --repo anjan-NZ/tANk
```

## First run

Open **Settings** in the pane and paste an API key. The Keys section lists every provider with a
button that opens its signup page. All of them are free and only Cerebras asks for a card.

Start with Groq, then add Gemini as a second key. Each provider counts its own limit, so with two
keys tANk switches over when one is busy instead of making you wait.

## What it does

**Excel.** Reads and edits the cells you select, writes formulas, applies real conditional
formatting, builds charts, pivots, tables and filters. Reconciles two ranges, ties out debits
against credits, picks audit samples, finds duplicates and missing numbers, ages balances.

**Word.** Rewrites the selection, inserts headings, lists and tables, runs find and replace, turns
track changes on, reads and answers comments, adds footnotes, headers and a table of contents.

**PowerPoint.** Builds slides from a brief, edits the text on a slide, adds shapes and images.

It asks before changing anything that already has content, and the Undo button in the pane puts
back whatever it last did.

## Privacy

There is no tANk server. Your key is stored on your own PC and goes only to the provider it
belongs to. Your file is never uploaded. When you ask a question, the cells or text you are
working on are sent to the AI provider you picked, because the model has to see them to answer.
Nothing is logged anywhere.

## Removing it

1. Open tANk, press **Settings**, then **Erase everything**. That clears your API keys and all
   settings immediately.
2. Close Excel, Word and PowerPoint.
3. Windows Settings, Apps, find tANk, Uninstall. Or run `uninstall.ps1` if you installed with the
   script.

Step 1 matters. The keys are held by the pane inside Office's own web storage, which an
uninstaller cannot pick apart without wiping what every other add-in on the PC has saved, so it
does not try.

Everything else goes on uninstall: tANk writes one registry value and one folder in your AppData,
and both are removed. It touches nothing else on the machine.

## If something goes wrong

**tANk is not in the Add-ins list.** Office was open during the install. Close all Office apps
and run the installer again.

**There is no tANk button on the ribbon.** There isn't one. Office only puts add-in buttons on the
ribbon for add-ins installed from the store or pushed out by a Microsoft 365 administrator, so
tANk lives under Home > Add-ins instead.

**The pane is blank.** Check your internet connection. The pane loads from the web.

**"has run out for now".** That provider hit its per minute limit. Add a second key in Settings
and tANk will switch by itself.

**A model stops working.** Open Settings and press Refresh models. Providers retire models often.

## Note

tANk is a personal project shared for learning purposes. It comes with no warranty and is not
connected to Microsoft or to any of the AI providers. Check whatever it writes before you rely on
it for real work.
