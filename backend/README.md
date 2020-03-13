# To Run the program

1. Double click RunWOJournalEntry.bat file.

2. Fill in the AiM login username and password in the command line window, note that the password won't be shown on the screen. Click enter and the program will continue.

![CMD login screenshots](images/screenshots1.PNG)

# Install the program (for developers)

1. Install [Miniconda (Python3.7)](https://docs.conda.io/en/latest/miniconda.html)  on your local machine. You can access conda via the console, to make sure it's properly installed please run `conda -V` to display the version.

2. Create new virtual environment. Open Anaconda Prompt and run:

    ```conda create --name selenium_env python==3.7 –y```

3.  Once it's created you can activate it by running: ```conda activate selenium_env```

4. Go to the current directory, run ```pip install -r requirements.txt``` under the activated virtual environment

5. Setup environment variables for Windows 10 System. Typical path is :C:\Users\user_ID\AppData\Local\Continuum\miniconda3\envs\selenium_env

6. To show hidden files in Windows 10:

![show hidden files](images/screenshots2.PNG)

7. Run the run_demo.bat, it will call WO_journal_entry(demo).py, and check if everything works.
8. Maybe need to download chrome webdriver, check user's ***chrome version***, and download it. Added to environmental variables.

# Important Notes for Users

1. If there are any questions or improvement suggestions, feel free to contact Tim.



# Reference

[Blog: Accessing Google Spreadsheet with Python](https://towardsdatascience.com/accessing-google-spreadsheet-data-using-python-90a5bc214fd2)