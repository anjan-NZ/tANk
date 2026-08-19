# Installing tANk

tANk adds an AI chat pane to Excel, Word and PowerPoint. It runs on your own free API key, so
there is nothing to pay for and no account to create.

Works on Windows with desktop Office (Microsoft 365, or Office 2019 and later).

## Install

1. Close Excel, Word and PowerPoint.
2. Download `install.cmd` from the latest release.
3. Double click it. No administrator password is needed; it installs for you only.
4. Open Excel and click **tANk** at the right of the Home tab.

`install.cmd` is the one to use. It copies a single file into your AppData and writes one
registry value, using tools that already ship with Windows. There is no program to install and
nothing runs from a temporary folder, so the security features that block unknown installers have
nothing to object to.

`tANk-Setup.exe` is also on the release page and does exactly the same three things through a
normal setup wizard. Prefer it if you like a wizard, but see the warning below: on some machines
Windows refuses to run it at all.

![The Add-ins menu on the Home tab, with tANk listed under Developer Add-ins](docs/where-to-find-tank.png)

If the button is not there yet, open **Add-ins** on the same tab and pick **tANk** under
Developer Add-ins, as in the picture above. Office can take a restart or two after installing
before it adds the button itself. The pane opens on the right and stays open while you work.

## If Windows blocks the exe

Two different things can happen, and only one of them can be clicked past.

**"Windows protected your PC", a blue box with More info and Run anyway.** That is SmartScreen.
Click **More info**, then **Run anyway**.

**"Smart App Control blocked an app that may be unsafe", or "Error 4551: An Application Control
policy has blocked this file".** That is Smart App Control, and there is no way past it. It refuses
any program whose publisher it cannot verify, and it does not offer an override.

Both happen because the installer is not code signed. Signing means buying a certificate and
renewing it every year, and this is a free project, so it does not have one. Neither message says
anything about what is inside the file; Windows shows them for every unsigned program.

Use `install.cmd` instead. Smart App Control has nothing to block there, because nothing is being
installed as a program. There is also a PowerShell version if you prefer it:

```powershell
irm https://anjan-NZ.github.io/tANk/install.ps1 -OutFile install.ps1
powershell -ExecutionPolicy Bypass -File install.ps1
```

If you do use the exe and want to check it is the one built from this repository, rather than
something swapped in later, verify the provenance GitHub records at build time:

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
3. Run `uninstall.cmd`. If you installed with the exe instead, use Windows Settings, Apps, find
   tANk, Uninstall.

Step 1 matters. The keys are held by the pane inside Office's own web storage, which an
uninstaller cannot pick apart without wiping what every other add-in on the PC has saved, so it
does not try.

Everything else goes on uninstall: tANk writes one registry value and one folder in your AppData,
and both are removed. It touches nothing else on the machine.

## If something goes wrong

**tANk is not in the Add-ins list.** Office was open during the install. Close all Office apps
and run the installer again.

**No tANk button on the ribbon yet.** Office adds it when it next rebuilds the ribbon, which can
take a restart or two. Until then, Home > Add-ins > tANk opens the same pane.

**The pane is blank.** Check your internet connection. The pane loads from the web.

**Answers pause, then carry on.** A provider hit its per minute limit and tANk moved to another
model by itself. It does not interrupt to say so; the transcript shows a "Show model switches"
link if you want to see what it did. Adding a second key gives it somewhere to go.

**A model stops working.** Open Settings and press Refresh models. Providers retire models often.

## Note

tANk is a personal project shared for learning purposes. It comes with no warranty and is not
connected to Microsoft or to any of the AI providers. Check whatever it writes before you rely on
it for real work.
