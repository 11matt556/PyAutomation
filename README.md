# **HOW DO I USE THIS?**

#### INSTALLATION
1. Download this project
2. Intall [Python 3.x](https://www.python.org/downloads/)
    - **Make sure to check  `Add to PATH` in the installer for easier usage**
3. Download and place [chromedriver.exe](https://sites.google.com/a/chromium.org/chromedriver/) into the same directory as main.py.
4. Open up an **admin** command prompt
5. Run `pip install selenium`

#### USAGE
1. Open `input.csv` and enter the hostnames of the computers you want to work on
    - The same task will be performed on all entries!
2. Close and save `input.csv`
3. Click on `main.py`
    - If this doesn't work, open up an **admin** command prompt and run `python main.py`
4. You should now be prompted to enter the type of task you want to do. Enter the task and press enter.
5. After the program completes it will create an `output.csv` and an `review.csv` file.
    - "output.csv" will contain the RITM for each computer in the format Tim's script is expecting
    -  **"output.csv" is cleared each time the program is run**
6. (Optional) If the program encountered any errors, it will print out the items that caued the errors. It will also save a copy as review.csv.
    - Please manually review these tickets to make sure they are completed properly
    - **"review.csv" is cleared each time the program is run**

Old versions of output.csv and review.csv can be found in the 'logs' folder. They are in the format of `Y-M-D_HM`. 

These logs are not automatically deleted at any point.

Please report any unusual behavior to me as it is very possible there are still some bugs
