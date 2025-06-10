# For Foxconn shipping process simplification

1. send email to nvidia for delivery number

2. read email from nvidia for get delivery number and other info

3. create GR file from nvidia feedfile or existing file, then auto complete the rest parts other than scanning in Serial Number for posting Good Record

```shell
// Sample
PS C:\Users\sfcuser\Desktop\NEIS_TOOLS> python .\main.py
ERROR: The process "excel.exe" not found.
[00:00:00] outlook:     LOADING OUTLOOK
[00:00:00] outlook:     OUTLOOK LOADED
[00:00:00] main:        Please select a step to continue:
                        1. Send DN request for a GR
                        2. Check Email for new DN (in progress, do not use) 
                        3. Build GR file
                        4. Build GR file from Feedfile
                        5. Model info look up
                        6. Working order number look up
                        7. Apply ITN for a DN
                        8. Product Tracking by SN
                        9. POD generate
                        0. Create Report
                        Enter quit to Quit

```
