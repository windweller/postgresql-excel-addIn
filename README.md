#PostgreSQL Excel Add-In

PostgreSQL is currently the fastest rising alternative database to commerical products like Oracle or recently aquired database MySQL. PostgreSQL implements strict SQL specifications compared to MySQL's not-so-strict implementation. In some research PostgreSQL is better at performing [certain tasks](http://www.wikivs.com/wiki/MySQL_vs_PostgreSQL#PostgreSQL) (join/subquerying) than MySQL, but to a large extent, we can consider these two databases equal.

This Excel Add-In provides automatic SQL generation with a GUI for users that are less familiar with SQL syntax.  

It can finish following SQL tasks:
* Retrieve tables within one database
* Retrieve specific columns in a table
* Search for exact match of rows
* Fuzzy search
* Custom SQL execution

This Add-In also provides a very easy way to add ODBC driver to User/System DSN (DataSource Name). 

Notice: this Add-In will not work on Macs, and if you are using Macs, congratulations, you already have very complete and strong database support. Go to Data tab and click Database from External Datasource. 

##What is DSN?

Programs use database drivers to connect to databases. Database developers for PostgreSQL and MySQL provide those drivers at their official website. For connecting to databases inside Excel, you have to download ODBC driver (there are also JDBC driver but it's for Java programs). 

This is the [ODBC Driver for PostgreSQL](http://www.postgresql.org/ftp/odbc/versions/msi/), simply download the latest version (at the very bottom). You will see two similar packages that read like `psqlodbc_09_03_0300-x64.zip` or `psqlodbc_09_03_0300.zip`. The `-x64` at the end of file name is the indication of the machine you are running (64-bit vs 32-bit). However, you must not install the `-x64` ending driver because Excel only works with 32-bit drivers (more precisely, the VBA - Visual Basic code only works with 32-bit driver; to be even more precise, it's the [ActiveX Object](http://en.wikipedia.org/wiki/ActiveX) that VBA uses to connect to database, due to its lack of continuing development and old age cannot work with 64-bit driver). 

After you download the zip file, decompress it and run the `.msi` MSI file and it will install automatically.

Then (here is why this Add-In is superior to current Windows database import option), `ctrl-j` to open this Add-In, click "List Data Source" button on the control panel. This will lead you to the correct DSN manager on your machine. It's worth mentioning you can't go to `Control Panel\System and Security\Administrative Tools` and open `DataSource (ODBC)`, because if your machine is running on Windows 64-bit, this shortcut will only lead you to 64-bit datasrouce manager, and 32-bit manager is hidden elsewhere, but don't worry, once you click "List Data Source", it will take you to the correct version.

Click "list data source" button and it will lead you to system Database Source Manager. Then click "Drivers" tab. If you have already installed the corresponding driver, it will show up (you might have to scroll down the list). Click "User DSN" tab, click "add" button, and fill the requested fields (server name, user name, password... you can find tutorials online on how to fill those fields). If you are truly uncertain, this is an example for PostgreSQL:

![DSN Screen Capture](https://github.com/windweller/postgresql-excel-addIn/blob/master/ScreenCapture/AddDSN.PNG)


