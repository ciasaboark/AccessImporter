# Installation instructions

## Prerequisites

### Microsoft Access Driver
The Access Database driver is required for the python odbc connection. This driver
is installed as part of Access, but can be installed separately.

If you already have Microsoft Office installed you can check which bit version is installed by opening one of the Office applications (Word, Excel, etc...).  Click on `File`, then click on `Account`.  In the window that opens click on the `About` button.

<img src="https://raw.githubusercontent.com/ciasaboark/AccessImporter/master/images/excel1.png" width="400px"/>

At the top of the window that opens there will be a line showing the application version and bit version.

<img src="https://raw.githubusercontent.com/ciasaboark/AccessImporter/master/images/excel2.png" width="400px"/>

The bit version (32 bit/64 bit) of the Access driver and the bit version of the Python interpreter must match.

If Office is not installed a standalone driver can be downloaded from https://www.microsoft.com/en-US/download/details.aspx?id=13255.



### Python 
Python 3.3 or later is required. The scripts were tested against Python 3.8.3.
Download the latest Python 3 version at https://www.python.org/downloads/windows/.

Choose the Python bit version that matches the Microsoft Access driver version.

By default Python will install to the users application folder and the interpreter will not be added to the environment path.

On the initial install screen check `Add Python to PATH` then click `Customize Installation`. On the next screen leave all options selected and click `Next`.

<img src="https://raw.githubusercontent.com/ciasaboark/AccessImporter/master/images/python1.png" width="400px"/>


On the Advanced Options screen check `Install for all users`.  This will change the install location to 'C:\Program Files\Python--version--'.  Click `Install`

<img src="https://raw.githubusercontent.com/ciasaboark/AccessImporter/master/images/python2.png" width="400px"/>

### Git
Git is not required for installation, but can be used to easily install the script and upgrade later if changes are made.

Download and install the latest version of Git from https://git-scm.com/download/win. The default option are recommended.

### Post Install
Open a new cmd.exe window.

Type `python --version`. Python will report back the version that was installed if it was added to the PATH.

```
C:\Users\My User>python --version
Python 3.8.3
```

Type `git --version`. Git will respond with the installed version if it has been added to the PATH.

```
C:\Users\My User>git --version
git version 2.27.0.windows.1
```

## Install the script

The Access Importer script can be installed either by downloading a zip file or by cloning the repository using Git.

If Git was not installed open https://github.com/ciasaboark/AccessImporter. Click on `Clone or download` and choose `Download ZIP`.  The zip file will contain a folder called AccessImporter-master. Extract this folder wherever you want to install the program.

If Git was installed open a cmd.exe window. Type `git clone https://github.com/ciasaboark/AccessImporter.git` to clone a fresh copy of the repository.

```
C:\Users\My User>cd Apps

C:\Users\My User\Apps>git clone https://github.com/ciasaboark/AccessImporter.git
Cloning into 'AccessImporter'...
remote: Enumerating objects: 41, done.
remote: Counting objects: 100% (41/41), done.
remote: Compressing objects: 100% (24/24), done.
remote: Total 41 (delta 17), reused 38 (delta 14), pack-reused 0
Unpacking objects: 100% (41/41), 19.59 KiB | 185.00 KiB/s, done.
```

### Install Python Modules
The script uses a few Python libraries. To install the libraries open a cmd.exe window as an Administrator. Change to the installation directory for the script.

Type `pip install -r requirements.txt` to install the required packages.

```
c:\Users\My User\Apps\AccessImporter>pip install -r requirements.txt
Collecting watchdog (from -r requirements.txt (line 1))
  Using cached https://files.pythonhosted.org/packages/73/c3/ed6d992006837e011baca89476a4bbffb0a91602432f73bd4473816c76e2/watchdog-0.10.2.tar.gz
Collecting pyodbc (from -r requirements.txt (line 2))
  Using cached https://files.pythonhosted.org/packages/00/3d/fcd43247944a14aa36c69ea8ad5993238f89c4ef214150cef26fdfff03fa/pyodbc-4.0.30-cp38-cp38-win_amd64.whl
Collecting configargparse (from -r requirements.txt (line 3))
  Using cached https://files.pythonhosted.org/packages/bb/79/3045743bb26ca2e44a1d317c37395462bfed82dbbd38e69a3280b63696ce/ConfigArgParse-1.2.3.tar.gz
Collecting xlrd (from -r requirements.txt (line 4))
  Using cached https://files.pythonhosted.org/packages/b0/16/63576a1a001752e34bf8ea62e367997530dc553b689356b9879339cf45a4/xlrd-1.2.0-py2.py3-none-any.whl
Collecting pywin32 (from -r requirements.txt (line 5))
  Using cached https://files.pythonhosted.org/packages/d6/ef/db6c352b19eee9f944c03a6ae304ac7c5f96ef3b8429ec905b3e2c64a4af/pywin32-228-cp38-cp38-win_amd64.whl
Requirement already satisfied: pathtools>=0.1.1 in c:\program files\python38\lib\site-packages (from watchdog->-r requirements.txt (line 1)) (0.1.2)
Installing collected packages: watchdog, pyodbc, configargparse, xlrd, pywin32
  Running setup.py install for watchdog ... done
  Running setup.py install for configargparse ... done
Successfully installed configargparse-1.2.3 pyodbc-4.0.30 pywin32-228 watchdog-0.10.2 xlrd-1.2.0
```


### Install the Script as a Windows Service
The script will run in the background as a Windows service. To install the service open a cmd.exe window as an administrator and change to the installation directory for the script.

```
c:\Users\My User\Apps\AccessImporter>python watcher.py install
Installing service container_tracking_importer
Service installed
```

### Create the Required Directories
The script will need three separate directories to hold the data.

- An import directory to watch to watch for newly created Excel files.
- An archive directory to hold successfully imported Excel files. This must not be the same directory as the import directory.
- An archive error directory to hold Excel files that could not be imported. This must not be the same directory as the import directory.

These directories will need to be created before the script can run. The location of each will be configured in the next step.

The script will also need to know where the Access Database is located as well as where to save log files. The location of each will be configured in the next step.

### Configure the Script Settings
In the cmd.exe window type `mmc Services.msc` to open the Windows Services window.

The script was installed as `Container Tracking Importer`. Scroll down to this script and double click on the service name.

<img src="https://raw.githubusercontent.com/ciasaboark/AccessImporter/master/images/service1.png" width="400px" />

Click on `Start` to start the service. You will see a error dialog saying that Windows could not start the service. This is fine. The initial run will populate configuration options in the Windows registry.

Type `regedit` in the cmd.exe window to open the Windows Registry. In the registry editor navigate to `HKEY_LOCAL_MACHINE\SOFTWARE\Container Tracking\`.

There should be five configuration options to set. Change each to point to the directory or file required.

- `watch`: The directory to watch for new Excel files. These files will be automatically imported into the Access database.
- `archive`: The directory to hold successfully imported Excel files. This must not be the same directory as the `watch` directory, but may be a sub directory of the `watch` directory.
- `errors`: The directory to hold Excel files that could not be imported. This must not be the same directory as the `watch` directory, but may be a sub directory of the `watch` directory.
- `log_file`: The name of the log file. This must be a full path with the file name (i.e. `C:\logs\importer.log`). The log file will be rotated every night and the previous seven days of log files will be kept. This file should not be in the `watch` directory or it will trigger excessive logging.
- `database`: The full path of the Access database to import the Excel data into.

### Configure the Service
By default the service will be set to run manually. If desired the service can be configured to start automatically.

Switch back to the Windows Services window and double click on the `Container Tracking Importer` service.

In the `General` tab select startup type `Automatic`.

By default the script will run using the Local System account. This may be an issue if network drives were selected in the [Configure the Script Settings](#configure-the-script-settings) section.

In the `Log On` tab select `This account`. Click on `Browse` and type in the login name of the user to use. Click on `Check Names` and the login information should fill in. Click on OK. In the previous window type in the users current password in both the `Password` and `Confirm password` boxes.


>**Warning!**
>
>If the windows account password is changed this step must be repeated to update the service with the correct password. The script will continue to supply the old password and may trigger an account lock out if it is started too often with an old password.


In the `Recovery` tab, change both `First failure` and `Second failure` to `Restart the Service`.

Click on OK to save the settings.


### Run the Service
In the Windows Services window double click on the `Container Tracking Importer` service. Click on `start` again. If the configuration options were set correctly the service should start.

Drag and drop an Excel file into the `watch` directory. The file should process within 1-2 seconds and will be moved to either the `archive` or `errors` directory.


## Checking for Errors
If the service fails to start the error may have occurred before the log file was initialized.  Warnings and error messages are logged to the Windows event log. Use start->run and type in `eventvwr` or search for 'Event Viewer' in the start menu.

Expand `Windows Logs` and select `Application`. Events for all applications will display. In the `Actions` pane select `Filter Current Log...`.

In the `Event sources` option select `container_tracking_importer`. Click on OK. The most recent events will be listed at the top.

If an Excel file does not import it will either be left in the `watch` directory if the import can be attempted again later, or moved to the `errors` directory if the script can not recover from the error.

Logging information is written to the log file selected in the [Configure the Script Settings](#configure-the-script-settings) section. An exception message and stack trace will be embedded for any failed import.

Successful import
```
╭╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼
╽
╽ Beginning import of C:\import\Weekly Tracked Pallets Shipped.xlsx
╽
╰╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼
Beginning import of file C:\import\Weekly Tracked Pallets Shipped.xlsx
Sha1 hash of file: 3ae1d6c483b1e14eec7d5d709ce5cc9fed11326a
Reading records
File appears to be for DC 7.0
Read 1417 records
Writing records into database
Wrote 1417 records
Archiving file as 'C:\import\archive\Weekly Tracked Pallets Shipped.xlsx.2020-06-18-13-29-26-405622'.
╭╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼
╽
╽ Finished import of C:\import\Weekly Tracked Pallets Shipped.xlsx
╽ Status: Archived
╽
╰╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼
```

Failed import
```
╭╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼
╽
╽ Beginning import of C:\import\Weekly Tracked Pallets Shipped.xlsx
╽
╰╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼
Beginning import of file C:\import\Weekly Tracked Pallets Shipped.xlsx
Sha1 hash of file: e32dacbc685ff98a557a594703899e1d5c52676f
Reading records
File appears to be for DC 6.0
Read 188 records
Writing records into database
Error importing row {'DC Id': 6.0, 'DC Name': 'MDV - Pensacola', 'Store Id': '5850', 'Store Name': 'NAS JACKSONVILLE COMMISSARY', 'Address': '1701 ALLEGHENY RD', 'City': 'NAS JACKSONVILLE', 'State': 'FL', 'Zip': '322120042', 'Transaction Date': 43998.0, 'Container Type': 'CP', 'Container Qty': 37.0, 'Transaction Date String': '2020-06-16'}
Unable to import data from file: ('23000', '[23000] [Microsoft][ODBC Microsoft Access Driver] The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship. Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again. (-1605) (SQLExecDirectW)')
Unable to import data due to an unrecoverable SQL error.
('23000', '[23000] [Microsoft][ODBC Microsoft Access Driver] The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship. Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again. (-1605) (SQLExecDirectW)')
Traceback (most recent call last):
  File "c:\Users\My User\Dropbox\Projects\AccessImporter\watcher.py", line 372, in import_file
    importer.begin_import()
  File "c:\Users\My User\Dropbox\Projects\AccessImporter\importer.py", line 36, in begin_import
    self.insert_rows(importData)
  File "c:\Users\My User\Dropbox\Projects\AccessImporter\importer.py", line 102, in insert_rows
    raise e
  File "c:\Users\My User\Dropbox\Projects\AccessImporter\importer.py", line 86, in insert_rows
    cursor.execute("INSERT INTO [{0}] ([DC ID], [DC Name], [Store ID], [Store Name], [Address], [City], [State],"
pyodbc.IntegrityError: ('23000', '[23000] [Microsoft][ODBC Microsoft Access Driver] The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship. Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again. (-1605) (SQLExecDirectW)')
Archiving file as 'C:\import\archive\errors\Weekly Tracked Pallets Shipped.xlsx.2020-06-18-13-31-55-652917'.
╭╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼
╽
╽ Finished import of C:\import\Weekly Tracked Pallets Shipped.xlsx
╽ Status: Error
╽
╰╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼
```

## Updating to a New Version
### Update the Script
Use Git to install an updated version of the software. If updates are available git will show output similar to the following example.

```
c:\Users\My User\Apps\AccessImporter>git pull
Updating 00fa356..2bd4074
Fast-forward
 requirements.txt | 1 -
 watcher.py       | 2 +-
 2 files changed, 1 insertion(+), 2 deletions(-)
```

If updates are not available a message will print saying everything is up to date.

```
c:\Users\My User\Apps\AccessImporter>git pull
Already up to date.
```

If the script was not installed using Git a new copy of the zip file must be downloaded and extracted. Unzip the contents into the same directory as the old install, overwriting any old files.

### Updating the Service
After a new version of the script is downloaded Windows will need to know about the updated service.

Open a cmd.exe window as an administrator. Change to the script installation directory and use `python watcher.py update` to update the service.

```
c:\Users\My User\Apps\AccessImporter>python watcher.py update
Changing service configuration
Service updated
```

Updating the service may reset some configuration options. Refer to the [Configure the Service](#configure-the-service) section to reconfigure the service.


