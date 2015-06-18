# README #

## GUI Interface to Interactive Brokers API ##

* This project provides a VB.NET GUI interface to Interactive Broker TWS CSharp API. 

* I use my own automated trading systems. I needed a GUI interface to connect to IB Gateway in order to streamline the automated trading strategies, gather account data, perform risk and order management, generate/place orders and generate/email reports. I found building GUI on Visual Studio was quite straight forward and hence built this GUI interface for IB Gateway. I have chosen the CSharp API because I found it to be stable. 

* Version 1.0,  2014

**DISCLAIMER:** This software is not approved by Interactive Brokers or any of its affiliates. It comes with absolutely no warranty and the use of this software for actual trading is at your own risk.


## What to do next?

You can either clone or download the repository. When you clone the repository, you will clone a folder **IB_TWS**. This folder is self-contained, you can open  the project in Visual Studio (works seamlessly on Visual Studio 2012).

Go to: 

* Open Visual Studio 2012 ->  File -> Open Project -> Browse to *YOUR_LOCAL_DRIVE/IB_TWS/* and open **IB_TWS.vbproj**

Build/Debug:

* Open -> File -> *YOUR_LOCAL_DRIVE/IB_TWS/dlgMain.vb*
* Build or Debug

## How to run GUI application?

* Dependency: 
    * For the GUI to run correctly with all its functionalities, it needs an XML file (**/utilities/ibSet.xml**). I have stored this XML file under *IB_TWS/bin/Debug/*
    * IB API (CSharpAPI.dll). If you have successfully installed IB API then this file is available inside *D:\TWS API\source\CSharpClient\client\bin\Release* (*I have committed this file, you will find it under /IB_TWS/bin/Debug/*)
* If you want to test/debug, after you "build" just double click on **IB_TWS.exe** (inside */IB_TWS/bin/Debug/*)
* If you want to use the Release version, then copy **/utilities/ibSet.xml** to your release folder (e.g. */IB_TWS/bin/Release/*) and then run the application.
 

## Features of this GUI

* Connection Tab
    * Connection status indicator
    * Connect to IB Gateway by setting IP address, log level, port number and client id including Connect and Disconnect buttons
    * Log: Server logs, Error logs, Clear error log button
* Account Tab 
    * Account number indicator
    * Last sync time indicator
    * Subscribe and Unsubscribe functionalities and buttons
    * Summary and Portfolio tabs
* Order Tab
    * Ticker description and Order description (including a pop gui for extra Ticker attributes and Order attributes)
    * What-If, Place Orders, GLOBAL CANCEL, Clear Orders buttons
    * Next Order-ID in-sync with IB Gateway
    * Open Orders and Order Status tab
    * Cancel Order by right-click on Open Orders
* Executions tab
    * Executions request settings
    * Request Executions and Clear Executions button
    * Reports tab
* Saved Settings
    * All settings i.e. connection, ticker and order settings saved in an XML file (ibSet.xml)
    * Every time the application loads, it will retrieve settings from the settings XML file
* Exposed functionalities
    * GUI and its functionalities are all exposed for you to build extra features on top
    * *System.Data.DataSet* used to store all data from IB Gateway
    * Always in-sync with IB Gateway i.e. whenever IB Gateway delivers data, all logs and datatables are updated
    * Datatable manipulation (e.g: insertion, retrieval, find etc)
    * checkConnection functionality to check if the connection between application and IB Gateway is alive
    * You can add more tabs/buttons/whatever you want!


## Want to build features to this GUI interface?

* It will be great if you can build features on top of this implementation, please do! :-)
* If you build new/extra features on top, please do commit your work to this repository. I will approve your commits.

## Found a bug? 

* If possible, please fix the bug and commit it. That will be really helpful
* Otherwise, you can use issue-tracking on this repository to report the bug

## Comments/Queries?

* Your comments/suggestions are always welcome. Send me a message via BitBucket.

## License

90% of the source code is from Interactive Brokers. Please refer to  [Interactive Brokers](http://www.interactivebrokers.com) for any licensing issues


Use the application and make it better! :-)

Have fun!

Whacky

[The Portfolio Trader](http://www.theportfoliotrader.com)