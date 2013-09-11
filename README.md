This requires Visual studio 2010 SP1 or Visual Studio 2012 to open it.

Should run on Windows or on Linux under mono.

This application must be passed 4 arguments:
  * <username> - The username of the user requesting a calendar. (E.g. abc1d23).
  * <password> - The pass word of the user requesting a calendar.
  * <calendar> - The name of the calendar. (E.g. building 32 room 4049 would be b32r4049).
  * <writefile> - The file to write the iCal for the calendar to. (E.g. C:\Users\User\b32r4049.ics or Windows or /home/user/b32r4049.ics on Linux).


## Running Under Windows ##
The executable for the application can be found in Exchange Calendar To iCal\bin\Debug.  To run this application use cmd and after changing directory
to iCal\bin\Debug type the following command (substituting where appropriate):

  "Exchange Calendar To iCal.exe" abc1d23 PASSWORD b32r4049 C:\Users\User\b32r4049.ics


## Running under Linux ##
The executable for the application can be found in Exchange\ Calendar\ To\ iCal/bin/Debug.  Before you can run this application you will no to install 
a complete version of mono.  (On Ubuntu 12.04 the package is called mono-complete).  Once installed you will need to run the following commands to
install SSL certificates.  (If you do not have sudo run as root):

  sudo mozroots --import --ask-remove --machine

  sudo certmgr -ssl https://www.outlook.soton.ac.uk:443

After running those command change directory to Exchange\ Calendar\ To\ iCal/bin/Debug and run something like the following command (substituting 
where appropriate):

  mono Exchange\ Calendar\ To\ iCal.exe abc1d23 PASSWORD b32r4049 /home/user/b32r4049.ics
