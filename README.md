# What Is This For?
A tool used for communicating with an Allen Bradley/Rockwell ControlLogix PLC to read and write tag values for their PlantPAX Library.

# Disclaimer
I am not responsible for anything bad that happens as a result of your use of this. You work in controls and automation, you should know this by now. 

# Why?
The Excel based PlantPAX Online Configure Tool provided by Rockwell only works if you have a PAID license for RSLinx Enterprise (or FactoryTalk Gateway) to use the OPC/DDE connections. 

# How does it work?
It uses LibreOffice's scripting environment to call some python scripts use a well written python driver/library that interacts with the processor. 

## How well does it work?
Pretty well! It's a bit slow with the massive number of parameters and can cause Libreoffice to hang, but it still seems to read/write correctly. 


# Development/Testing Notes
This was built and tested with the following, no idea if it will work in the past or future versions of anything:
   1. Operating System: [Ubuntu 24.04 LTS](https://releases.ubuntu.com/) (running in VirtualBox)
   1. [LibreOffice 24.2.4](https://www.libreoffice.org/download/release-notes/)
      1. [APSO - Alternative Script Organizer for Python](https://extensions.libreoffice.org/en/extensions/show/apso-alternative-script-organizer-for-python)
   2. [Python 3.12.4](https://www.python.org/download/releases/)
   3. [Pycomm3 1.2.14](https://docs.pycomm3.dev/en/latest/)
   4. ControlLogix 1756-L81E
      1. [Firmware 35.01.00](https://compatibility.rockwellautomation.com/Pages/ProductReplacement.aspx?crumb=101&restore=1&vid=55729)
   5. [PlantPAX 4.10.06](https://compatibility.rockwellautomation.com/Pages/ProductReplacement.aspx?crumb=101&restore=1&vid=55212) AOIs


# Getting Started

## Setup Environment
1. Install python 
2. Intall pycomm3 `pip install pycomm3 --break-system-packages` (or use proper environment packages)
   1. Recommend testing pycomm3 works prior to trying the Spreadsheet
3. Install libreoffice
4. Open LibreOffice Calc
5. Navigate to the tab for the PAX object you want to read
6. Click 'Find Tags'
7. Click 'Read Tags'
8. Make changes to calues
9. Click 'Write Tags'
---
Not sure if this is required for use, maybe development only
1. Install APSO `apt install libreoffice-script-provider-python`
2. Make directory for the python scripts to be located `~/.config/libreoffice4/user/Scripts/python`
   1. Open Libreoffice and add this location
      1. Tools > Macros > Organize Macros > Organizer > Add Library
----

# Credit
I couldn't make anything like this if it wasnt for Free Open Source Software (FOSS). I owe much of my career to the resources found online from the thousands of people who share their knowledge and passion.
Please pay it forward and release anything you make, for free, forever. Donate to FOSS projecst you enjoy, be kind to the people who make things, don't expect that they can answer all your complaints. These are passion projects for most people. And document it!

