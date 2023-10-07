# WrapNAPS
Wrapper for NAPS2 (Not Another PDF Scanner 2)

# Background
Previously a tech at my last job utilized NAPS2 and a batch script to execute NAPS2.console.exe to perform a one-shot scan using a premade profile. One day I needed to set this up for a user with zero documentation on NAPS2, or the batch script. Several users had problems with the app not working, and would never provide any feedback to the user. Techs were extracting this app and manually editing the batch script to scan to a temp file then move it to a new file using time/date system variables to name the file. I wrote this wrapper to be a drop in replacment for that batch script.

# Features
- During installation mode, creates "Scan to PDF" to desktop, and prompts the user for default scan destination folder.
- Error feedback.
- Select scanner via dropdown.
- Select document source (Duplex, Flatbed/Glass, Document Feeder)
- Detects scanner's capabilities (For supported document sources)
- Option to remember last settings.
- Option to keep open for batch scans.
- Option to launch full NAPS2.
- Update checking.
- Post scan options like open folder/scanned file.

# Download
You can download the latest release in the releases section of this repo. The release is a 7-Zip Self Extracting archive containing `Init.NAPS2.exe, App, Data, and LICENSE` The 7z SFX will extract to the current user's profile under `AppData\Local\Programs\NAPS2`, then run `Init.NAPS2.exe ~!Install`.

# Usage
1. After the installation finishes and you click the "Scan to PDF" link on your desktop.
2. If you have a scanner available it will appear in the top dropdown. If there is only one, it will be automatically selected, and conversely if you have no scanners attached you will be unable to select a scanner and must restart the program after one is connected. (In the future the program will detect when scanners are connected/disconnected).
3. After you select a scanner you will need to select a Document Source. Flatbed, Glass, or Duplex. (Depending on the scanner or driver used some of these options may be unavailable).
4. By default the `Remember options` will be selected, if you dont want to save the changes made during this session simply uncheck this option.
5. By default WrapNAPS will exit after you start a scan. Check `Stay Open (Batch Scan)` to keep WrapNAPS open for additional scans.
6. Finally click `Scan` to start the scan. You will be presented with with a tooltip showing "Scanning..."
- Click `NAPS2` to run the full NAPS2 (Not another PDF Scanner) application.

# Command line arguments and function
- `~!Install`  Creates a shortcut on the desktop named "Scan to PDF" that points to Init.NAPS2.exe. If an existing shortcut is found it is deleted. The user will be prompted for a scan destination, which defaults to the user's `OneDrive\Documents\Scans` folder and Registry keys are created with the user's preferred scan destination at `HKCU\Software\InfinitySys\Apps\WrapNAPS`.
- `~!Update` is intended to only be called internally, it will backup the existing `Init.NAPS2.exe` to `Snaps\WrapNAPS2_v1.YYMM.DDHH.mmSS.exe` if it doesnt already exist, copy itself over `Init.NAPS2.exe` and then run `Init.NAPS2.exe`.
- `~!Recovery` can be called manually if the current version is not functioning as expected. It will roll back the current `Init.NAPS2.exe` to a previous version. You will be presented with a list of available versions. The function depepnds on the `Snaps` directory. **THIS FEATURE IS NOT CURRENTLY IMPLEMENTED.**

# About the update checker
- Currently the update function will check the `VERSION` file in this repo against the interval version, and then it will download `Init.NAPS2.exe`.
- The current version and all previous versions are stored in AppData/Local/Programs/NAPS2/Snaps/WrapNAPS2_v1.YYMM.DDHH.mmSS.exe

# To Install
- Download and run the installer from the Releases page.

# To Build
- Visit https://www.autoitscript.com/site/autoit/downloads/
- Downlaod and Install the latest version of AutoIt.
- Then goto https://www.autoitscript.com/site/autoit-script-editor/downloads/
- Download and Install `SciTE4AutoIt3.exe` (SciTE Editor & Utilities)
- Then open `_.Sources\Init.NAPS2.au3` in this repo.
- Press `F7` or goto Tools -> Build to compile.


# Screenshot
![image](https://github.com/BiatuAutMiahn/WrapNAPS/assets/6149596/f4ab9ada-050a-4252-b810-e1cbbf4860f5)


