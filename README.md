# WrapNAPS
Wrapper for NAPS2 (Not Another PDF Scanner 2)

Background:
Previously a tech at my last job utilized NAPS2 and a batch script to execute NAPS2.console.exe to perform a one-shot scan using a premade profile. One day I needed to set this up for a user with zero documentation on NAPS2, or the batch script. Several users had problems with the app not working, and would never provide any feedback to the user. Techs were extracting this app and manually editing the batch script to scan to a temp file then move it to a new file using time/date system variables to name the file. I wrote this wrapper to be a drop in replacment for that batch script.

Features:
  -During installation mode, creates "Scan to PDF" to desktop, and prompts the user for default scan destination folder.
  -Error feedback.
  -Select scanner via dropdown.
  -Select document source (Duplex, Flatbed/Glass, Document Feeder)
  -Detects scanner's capabilities (For supported document sources)
  -Option to remember last settings
  -Option to keep open for batch scans.
  -Option to launch full NAPS2.
  -Update checking.

About the update checker:
  -Currently the update function will check the VERSION file in this repo against the interval version, and then it will download Init.NAPS2.
  -The current version and all previous versions are stored in AppData/Local/Programs/NAPS2/Snaps/WrapNAPS2_v1.YYMM.DDHH.mmSS.exe

Screenshot:

![image](https://github.com/BiatuAutMiahn/WrapNAPS/assets/6149596/0cb677a7-b63f-4ed5-90b4-650bc3c3063f)


