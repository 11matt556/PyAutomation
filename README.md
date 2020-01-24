# **HOW DO I USE THIS?**

#### INSTALLATION
1. [Download this project](https://gitlab.com/11matt556/PyAutomation/-/archive/master/PyAutomation-master.zip)
2. Intall [Python 3.x](https://www.python.org/downloads/)
    - **Make sure to check  `Add to PATH` in the installer for easier usage**
3. Download and place [chromedriver.exe](https://sites.google.com/a/chromium.org/chromedriver/) into the same directory as run_me.py.
4. Open cmd or powershell as an admin
5. Run `pip install selenium`

#### USAGE
1. Open `input.csv`
2. In column A, enter the hostnames of the computers
3. In column B, enter the action the script should take on each hostname. Valid actions are "restock", "repair_isc", and "decommission"
4. Close and save `input.csv`
5. Click on `run_me.py`
6. ServiceNow may prompt you to login. If it does, the program will wait for youto login before continuing.
7. After the program completes all the tikets it will create two files - `output.csv` and `review.csv`
8. `output.csv` is used to print the labels for the computer. If you have Tim's labeling script then all you need to do is enter your name and press ctrl+alt+p to print the labels.
9. `review.csv` contains any hostnames the program wasunable to complete tickets for. For example, there may have been no reclaim found for the hostname.
10. Verify and/or complete any tickets in `review.csv`


# TROUBLSHOOTING
#### Unable to start run_me.py
* Verify python is installed correctly by typing `python --version` in the command line.
* Try running `run_me.py` from the command line instead of clicking on it
* Try running `run_me.py` from an admin command prompt
* Make sure chromedriver.exe is in the directory and that Google Chrome is installed on your computer

#### Blank browser screen after program starts, or program gets stuck on ServiceNow homepage
* There is a known issue on some computers that cause the program to freeze shortly after opening ServiceNow. 
* The only known solution is reimaging your computer or trying a different computer
