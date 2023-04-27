# Get Server Update Status

this script reads out if there are windows updates open for longer than 30days
it creates and excel output and send an email to the user/group that is set as manager for the server

## Installation 

You need PS Version 5.1 or higher

You need Powershell AD Tools
https://4sysops.com/wiki/how-to-install-the-powershell-active-directory-module/

To export in excel you need this package:
https://www.powershellgallery.com/packages/ImportExcel/7.4.1

## Usage

Enter values into XML file

``` xml
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<Values>
  <sender>sendermail</sender>
  <smtp>smtpserver</smtp>
  <excluded_servers>excludes</excluded_servers>
  <searchbase>OU of Servers you are looking for</searchbase>
  <filepath>path + filename.extension where you wanna store the list</filepath>
</Values>
```

run script 

## Contributing 

pull requests are welcome.

## License

feel free to use it, i'm happy to help.

