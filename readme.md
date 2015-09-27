# PLC IP Address Getter

## Purpose:

This program will take a directory of .L5X files and crawl over the IO trees
in each program pulling out the name, device model number and IP Address of each 
device. The data from each program will be put in its own excel sheet. It also 
will create a sheet with all available IP addresses for each subnet found in the 
process. 

## Dependencies:

- xmltodict
- easygui
- xlwt
- os
- collections
