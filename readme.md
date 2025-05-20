# For Foxconn shipping process simplification

1. send email to nvidia for delivery number

2. read email from nvidia for get delivery number and other info

3. create GR file from nvidia feedfile or existing file, then auto complete the rest parts other than scanning in Serial Number for posting Good Record

```shell
// Sample

[00:00:00] excel:       LOADING EXCEL
[00:00:10] excel:       EXCEL LOADED
[00:00:15] outlook:     LOADING OUTLOOK
[00:00:17] outlook:     OUTLOOK LOADED
[00:00:17] main:        Please select a step to continue:
                        1. Sending DN request for a GR
                        2. Checking Email for new DN
                        3. Build GR file
                        Enter quit to Quit
                        3
[00:00:17] GR:  START BUILDING GR FILE
[00:00:29] GR:  START SCANNING
[00:00:29] GR:  PLEASE START SCANNING SN.ENTER "quit" TO QUIT. WHEN FINISH, PRESS ENTER TO CONTINUE
158xxxxxxxxxx
158xxxxxxxxxx

[00:01:15] GR:  2 UNITS, CORRECT? (PRESS ENTER TO CONTINUE, quit to Quit, ANYTHING ELSE TO SCAN MORE)

[00:01:16] GR:  SCANNING COMPLETE, START PROCESSING DATA
[00:01:18] GR:  PROCESSING COMPLETE, FINALIZING GR FILE
[00:01:19] GR:  GR FILE BUILDING COMPLETE, FIEL CREATE: E:/0520 PB-59484_2x_5202A50121.xlsx
[00:01:20] main:        GR file ready
[00:01:20] main:        Please select a step to continue:
                        1. Sending DN request for a GR
                        2. Checking Email for new DN
                        3. Build GR file
                        Enter quit to Quit
                        1
[00:03:49] excel:       LOADING GR FILE
[00:03:49] excel:       GR FILE LOADED
[00:03:49] excel:       READING SN
[00:03:50] excel:       READING COMPLETE, 2 ROWS READ
[00:03:50] excel:       BUILDING SN STT TABLE
[00:03:50] excel:       TABLE BUILT
[00:03:51] main:        Email for sent
[00:03:51] main:        Please select a step to continue:
                        1. Sending DN request for a GR
                        2. Checking Email for new DN
                        3. Build GR file
                        Enter quit to Quit
                        1
[00:10:47] excel:       LOADING GR FILE
[00:10:48] excel:       GR FILE LOADED
[00:10:48] excel:       READING SN
[00:10:48] excel:       READING COMPLETE, 4 ROWS READ
[00:10:48] excel:       BUILDING SN STT TABLE
[00:10:48] excel:       TABLE BUILT
[00:10:48] main:        Email for sent
[00:10:48] main:        Please select a step to continue:
                        1. Sending DN request for a GR
                        2. Checking Email for new DN
                        3. Build GR file
                        Enter quit to Quit
                        quit
[00:12:06] main:        Program closing
[00:12:06] excel:       CLEAR UP START
[00:12:06] excel:       CLEAN UP FINISH
```
