# Altium Library Loader — "PCBLib object not set" / settings file keeps corrupting (root cause & fix)

A guide to a set of failures with the SamacSys **Altium Library Loader** VBScript, where importing an ECAD model crashes or the settings file (`AltiumLL.txt`) keeps getting corrupted — especially **on the second import** in a session.

**Confirmed working on Altium Designer 26.8.1.**

Typical errors you may see:

- `PCBLib` — *Object not set* (at `PCBServer.GetCurrentPCBLibrary`)
- *No PCB library focused — the .PcbLib document did not open*
- `CurrentLib` — *object required* (at `SCHServer.GetCurrentSchDocument`)
- `fso.CreateTextFile(...)` — **Permission denied**
- **SchLib / PcbLib paths in Settings disappear**, often right after importing a second component

---

## TL;DR

Every one of these symptoms traces back to a **corrupted `AltiumLL.txt` settings file**. The loader expects exactly **10 lines** in a fixed order; when the file is short, shifted, or unwritable, the library paths are read as garbage, the libraries never open, and the script crashes trying to use a library that isn't there.

There are **three** distinct triggers, and you may hit more than one:

1. **Bad initial file** — the file is short/shifted, so paths load as garbage.
2. **Unwritable file** — `AltiumLL.txt` is read-only (or the folder is locked/synced), so `UpdateTXT` fails with *Permission denied* and can never save your fix.
3. **Second-import corruption (the big one)** — the import flow leaves the settings text boxes in a bad state, and the next `UpdateTXT` writes that garbage over your good paths. This is why it works once, then breaks on the second component.

The durable fix is to **guard `UpdateTXT` so it refuses to write an invalid file**, plus keep the file writable and correctly formatted. Details below.

---

## The settings file format

The loader stores its settings here:

```
%USERPROFILE%\Documents\AltiumLL\AltiumLL.txt
```

It is a **fixed 10-line file**, one value per line, read by line number (`ReadTXT(n)`) and rewritten in full (`UpdateTXT`):

| Line | Setting | Example |
|------|---------|---------|
| 1 | Username (email) | `your_email@example.com` |
| 2 | Password | `your_password` |
| 3 | Downloads folder | `C:\AltiumLL\Downloads` |
| 4 | **SchLib path** | `C:\AltiumLL\MyLib.SchLib` |
| 5 | **PcbLib path** | `C:\AltiumLL\MyLib.PcbLib` |
| 6 | Use proxy (`True`/`False`) | `False` |
| 7 | Proxy address | *(blank)* |
| 8 | Proxy port | *(blank)* |
| 9 | Show instruction (`True`/`False`) | `False` |
| 10 | Altium symbols (`True`/`False`) | `False` |

A correct file (`AltiumLL.example.txt` is included alongside this doc):

```
your_email@example.com
your_password
C:\AltiumLL\Downloads
C:\AltiumLL\MyLib.SchLib
C:\AltiumLL\MyLib.PcbLib
False


False
False
```

> Lines 7 and 8 are intentionally **blank**. Lines 4 and 5 must be full paths to **existing** `.SchLib` and `.PcbLib` files.

---

## Symptoms → what they mean

| Symptom | What it means |
|---------|---------------|
| Crash at `Set PCBLib = PCBServer.GetCurrentPCBLibrary` | The `.PcbLib` never opened → `GetCurrentPCBLibrary` returned `Nothing` → dereferencing it crashes. |
| "No PCB library focused" | Same, but caught by a guard: line 5 (PcbLib path) is wrong/blank. |
| `CurrentLib is required` at `SCHServer.GetCurrentSchDocument` | The `.SchLib` never opened: line 4 (SchLib path) is wrong/blank. |
| `CreateTextFile` — Permission denied | `AltiumLL.txt` is read-only or the folder is locked → `UpdateTXT` can't save. |
| Paths vanish, especially after a 2nd import | The import flow corrupted the text boxes and `UpdateTXT` wrote the garbage back. |

---

## Root cause #1 — a short / shifted file

If the file has **fewer than 10 lines** or values are **shifted**, then:

- Lines 4/5 no longer hold valid library paths.
- `Client.OpenDocument("PcbLib", <bad path>)` **silently fails** — no library opens.
- `PCBServer.GetCurrentPCBLibrary` / `SCHServer.GetCurrentSchDocument` return **`Nothing`**.
- The next line dereferences that `Nothing` → **"object not set / required"** crash.

Example of a real broken file (only 6 lines, values shifted):

```
user@example.com   <- line 1  username        OK
mypassword         <- line 2  password        OK
0                  <- line 3  downloads       (should be a folder)
80                 <- line 4  SchLib path     GARBAGE
0                  <- line 5  PcbLib path     GARBAGE
0                  <- line 6  proxy flag
                       line 7  MISSING
                       line 8  MISSING
                       line 9  MISSING
                       line 10 MISSING
```

Leftover proxy values (`80`, `0`) landed in the SchLib/PcbLib slots, and lines 7–10 didn't exist, so `ReadTXT(7..10)` failed and those boxes blanked out.

---

## Root cause #2 — the file can't be written (Permission denied)

`UpdateTXT` does:

```vbscript
Set f = fso.createTextfile(TXTFileName, True)   ' True = overwrite
```

`CreateTextFile(..., True)` **cannot overwrite a read-only file** — it raises **Permission denied**. So if `AltiumLL.txt` is marked Read-only (or the folder is locked / synced by OneDrive and locked at that moment), every save fails and your corrected file is never persisted.

**Do not** mark `AltiumLL.txt` as Read-only to "protect" it — that's what causes this error. Keep it writable; the guard in the fix below is what protects it instead.

---

## Root cause #3 — the second-import corruption (the main bug)

This is the one that makes the file "mess up again" after it was fixed.

`UpdateTXT` blindly writes **whatever is currently in the settings text boxes** to disk. During an import, `ProcessCB` reads from and reassigns those boxes and shared variables (`username`, `password`, `SchLib`, `PcbLib`, `partID`, grid rows, etc.).

- On **import #1**, the boxes are still clean, so any save writes good values.
- On **import #2**, leftover state from the first run (a box that got touched, a stale variable, the selected grid row) is still populated. When any handler then fires `UpdateTXT`, it writes that **wrong state** back into `AltiumLL.txt`, shifting/overwriting your good paths.

Result: it works once, then corrupts on the second component — exactly the pattern observed. Nothing external is touching the file; **the app is saving bad in-memory state to disk.**

---

## The fix

### 1. Restore a correct, writable file

1. **Close** the Library Loader **and** Altium Designer.
2. Make sure `AltiumLL.txt` is **NOT Read-only** (right-click → Properties → untick Read-only).
3. Keep it in a **normal, writable, non-synced** folder. The loader reads from `%USERPROFILE%\Documents\AltiumLL\`, so if your Documents is on OneDrive, either pause sync or set that folder to "Always keep on this device".
4. Replace the contents with the correct **10 lines** (use `AltiumLL.example.txt`, put your real paths on lines 4 and 5).

### 2. Guard `UpdateTXT` so it can never write a bad file (the durable fix)

This stops the second-import corruption at the source: the file is only ever written when the SchLib/PcbLib boxes actually look valid. Replace the existing `UpdateTXT` with:

```vbscript
Function UpdateTXT()
   ' Refuse to save if critical fields are blank/invalid.
   ' Prevents the import flow from overwriting good paths with garbage
   ' (the "settings corrupt on 2nd import" bug).
   Dim s, p
   s = Trim(txt_SchLib.Text)
   p = Trim(txt_PcbLib.Text)

   If s = "" Or p = "" Then Exit Function
   If InStr(LCase(s), ".schlib") = 0 Then Exit Function
   If InStr(LCase(p), ".pcblib") = 0 Then Exit Function

   TXTFileName = InstalledDir & "\AltiumLL.txt"
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.createTextfile(TXTFileName, True)
   With f
      .Writeline txt_Username.Text
      .Writeline txt_Password.Text
      .Writeline txt_DnldsFldr.Text
      .Writeline txt_SchLib.Text
      .Writeline txt_PcbLib.Text
      .Writeline chk_Proxy.Checked
      .Writeline txt_Address.Text
      .Writeline txt_Port.Text
      .Writeline chk_ShowInstruction.Checked
      .Writeline chk_AltiumSymbols.Checked
      .Close
   End With
   Set fso = Nothing
End Function
```

Now, if the import ever leaves the path boxes blank or wrong, `UpdateTXT` **quietly refuses to write** instead of corrupting the file. Your good paths survive every import.

### 3. (Optional) harden the library lookups so they never crash silently

These turn a "document didn't open" situation into a clear message, and fix a latent missing-`Set` bug in `AddPcbLib`.

**`AddPcbLib`:**

```vbscript
Function AddPcbLib(footprint)
    Dim PCBLib, j, Component
    On Error Resume Next

    Client.StartServer("PCB")          ' make sure the PCB server is loaded

    If PCBServer Is Nothing Then
       MsgBox "PCBServer is Nothing - the PCB server did not start.", vbSystemModal, "Altium Library Loader"
       AddPcbLib = False
       Exit Function
    End If

    Set PCBLib = PCBServer.GetCurrentPCBLibrary   ' NOTE: 'Set' (original was missing it)
    If PCBLib Is Nothing Then
       MsgBox "No PCB library is focused - the .PcbLib document did not open. Check the PcbLib path in Settings.", vbSystemModal, "Altium Library Loader"
       AddPcbLib = False
       Exit Function
    End If

    On Error GoTo 0
    For j = 0 To PCBLib.ComponentCount - 1
        Set Component = PCBLib.GetComponent(j)
        If Component.Name = footprint Then
           AddPcbLib = False
           Exit Function
        End If
    Next
    AddPcbLib = True
End Function
```

**`AddSchLib`:**

```vbscript
Function AddSchLib(component)
    Dim CurrentLib, LibraryIterator, LibComp, LibCompNameNext, LibCompNamePrev
    Set CurrentLib = SCHServer.GetCurrentSchDocument
    If CurrentLib Is Nothing Then
       MsgBox "No schematic library is focused - the .SchLib document did not open. Check the SchLib path in Settings.", vbSystemModal, "Altium Library Loader"
       AddSchLib = False
       Exit Function
    End If

    Set LibraryIterator = CurrentLib.SchLibIterator_Create
    ' ... rest of the original function unchanged
End Function
```

---

## Verifying the fix (Altium 26.8.1)

1. Apply the guarded `UpdateTXT` (and optionally the two lookup guards).
2. Put a correct, **writable**, non-read-only `AltiumLL.txt` in place.
3. Open the loader → **Settings** → confirm both SchLib and PcbLib paths show **before** doing anything else.
4. Import component #1.
5. **Import component #2** — this is the case that used to corrupt the file. With the guard, the paths stay intact.
6. Close and reopen the loader → Settings → confirm the paths are still there.

Tested and confirmed working on **Altium Designer 26.8.1**.

---

## Why the fix works

The file was never corrupted by an external force — the app was writing bad in-memory state to disk. On import #1 the state is clean; on import #2 it isn't. By guarding the **save** (only write when the paths are valid), the corruption chain is broken permanently, regardless of what the import flow leaves in the text boxes. Keeping the file writable (not read-only, not locked by sync) ensures the good state can actually persist.

---

## How to avoid it in the future

- Set library paths with the **browse buttons**, not by typing — browsing commits the save.
- Point SchLib/PcbLib at **real, existing files**.
- Keep `AltiumLL.txt` **writable** — not Read-only, not on a locked/offline-synced drive.
- Apply the **guarded `UpdateTXT`** so multi-import sessions can't corrupt the file.
- If settings ever misbehave, close the loader, restore the correct 10-line file, reopen, and verify in Settings before importing.

---

## Security note

`AltiumLL.txt` stores your login in **plain text**. **Never commit a real `AltiumLL.txt`, email, or password to a public repository** — use placeholder values (as in `AltiumLL.example.txt`). Deleted commits still live in git history. If your real credentials have ever been exposed (pasted in chats, screenshots, issues), **change that account password**.
