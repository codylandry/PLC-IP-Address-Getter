# PLC IP Address Getter

## Purpose:

This program will take a directory of .L5X files and crawl over the IO trees
in each program pulling out the name, device model number and IP Address of each 
device. The data from each program will be put in its own excel sheet. It also 
will create a sheet with all available IP addresses for each subnet found in the 
process. 

## How to use:

1. Save all of your Allen Bradley .ACD programs in .L5X format in a folder on your desktop
2. Run the main.py file and follow the prompts:
  1. Select your folder with the .L5X files
  2. Type in a name  and select a location for the excel file the program will generate.
3. You're all done! Check out your new excel file listing all of your PLC's Ethernet devices with name, device model # and IP Address.
4. Also, take a look at the last sheet.  If all of your PLC's are on the same subnet, you can use this sheet to find the unused addresses.  

## Dependencies:

- xmltodict
- easygui
- xlwt
- os
- collections

#### Note: You can use the setup.py file to make this program an executable (.exe) 
*Use:*

```
python setup.py py2exe
```
