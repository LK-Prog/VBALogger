# VBALogger
## Description
This repo contains a class *vbLogger.cls*, which supplies basic logging functionalities for VBA and can be used as a starting point or framework for building further logging related features.

## Default Functionality
- Basic logging capabilites
- Working with multiple loggers in one file
- Level based conditional logging
- Roll over strategies
- Creating log reports inside a given excel workbook
- Logging related exceptions
- Timing functions

## How To Start
Either add the cls-file to your vba-project or copy the code into an existing vba-class.

1. Start by instantiating an object of the typ *vbLogger* (or your name of the class)

    ``` vba
    Public Sub vbLoggerDemo()
        Dim logger as vbLogger
        Set logger = new vbLogger
        
        Set logger = Nothing
    End Sub
    ```

2. Make a call to the initLogger function

    ``` vba
    Public Sub vbLoggerDemo()
        Dim logger as vbLogger
        Set logger = new vbLogger
        
        logger.initLogger "path/to/logfile.txt", WARN, False
        
        Set logger = Nothing
    End Sub
    ```
    (vba does not support constructor parameters, so this "factory" is nessecary)
    
    All init-calls can be run at anytime. Should the required resources already be created, they will not recreate them.

3. Define a roll over strategie if nessecary

    ``` vba
    Public Sub vbLoggerDemo()
        Dim logger as vbLogger
        Set logger = new vbLogger
        
        logger.initLogger "path/to/logfile.txt", WARN, False
        logger.initRollOver 3, True, #1/1/2022#
        
        Set logger = Nothing
    End Sub
    ```
    
4. Log text for the definied log file

    ``` vba
    Public Sub vbLoggerDemo()
        Dim logger as vbLogger
        Set logger = new vbLogger
        
        logger.initLogger "path/to/logfile.txt", WARN, False
        logger.initRollOver 3, True, #1/1/2022#
        
        logger.logText Level:=INFO, s:="Initial log text", userName:="vbLogger", force:=True
        
        Set logger = Nothing
    End Sub
    ```
