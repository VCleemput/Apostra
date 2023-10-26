# Apostra
**APO** ge**ST**ructureerde **RA**pporten - Flemish implementation in AutoHotkey
## Introduction
This repository aims to provide an easy approach to roll out structured reporting in the context of pathological anatomy in nearly any environment that accepts copy-paste commands.
To do so, it leverages the functionalities of [AutoHotkey](https://github.com/AutoHotkey/AutoHotkey) in order to show easy to navigate GUI's that fill in and paste templates. Where available, these templates aim to follow the directives of the Belgian Society for Pathology.

## Running the script
Just copy the repository and unzip, and open the Apostra.ahk file with the included AutoHotkey executable (either by dragging it on top of the executable, or by context menu "open with"). No installation required.
Some information about the antibodies or equipment used in your lab can be entered in "lab-variables.ini"

## How to use
After running the script, just type the hotkey or hotstring in the application of your choice.
### Currently implemented
At the moment, the script contains the following hotkeys
- windows+b: backup: pastes the previous text again if something went wrong
- *69pb : breast punction core needle biopsy
- *qs: "quickscore": hormone-receptors, HER2 and ki67. After generating the text, Windows+s optionally pastes the "short version".
- *88: melanoma
- *41: colonresectie
- *pd: proposition PD-L1 reporting according to the minimal requirements of the cancer registry. Fill in your lab-specifica parameters in the lab-variables.ini-file.


## Contribution
If you'd like to contribute in any way or have suggestions for improvements (both code-wise or regarding the content), just start an issue or open a pull request.

## Disclaimer
This code is supplied as-is. You yourself remain responsible for the reports you produce.

