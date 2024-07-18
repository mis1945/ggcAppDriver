Attribute VB_Name = "modHistory"
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'Project    : ggcAppDriver
'Author     : GGC SEG/SSG
'Copyright  : Guanzon Group of Companies (GGC)
'             Copyright(c) 2007 and Beyond
'             All Rights Reserved.
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'01.31.2007    09:55pm     XerSys
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Reminder:
'  + Show Warning form.
'x + Update neterror of project and syserror of user inside xImportLogError
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'02.01.2007    11:00am     KaLYPTuS
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'  + Reminder regarding neterrors is now okay.
'02.27.2007    09:00am     KaLYPTuS
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'  + Change pxeCharWidth to xeCharWidth
'  + Remove the function that displays modification information of a record.
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'03.08.2007    09:00am     KaLYPTuS
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'  + Added nUserRght from the SQL statement in the setMDIMain module of
'    clsAppDriver.
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'07.03.2008    04:35pm     Jheff
' Skin Color
' Default Color = All Control
'     wb0;et0
' Buttons
'     loControl.BackColor = p_oAppDrivr.getColor("HB1")
'     loControl.BackColorDown = p_oAppDrivr.getColor("HB4")
'     loControl.BorderColorFocus = p_oAppDrivr.getColor("BC0")
'     loControl.BorderColorHover = p_oAppDrivr.getColor("BC1")
'     loControl.ForeColor = p_oAppDrivr.getColor("ET0")
' Skin Names
'     nColorWT0 = Window Text Color
'     nColorWB0 = Window Background Color
'     nColorFT0 = Frame Text Color
'     nColorFB0 = Frame Banckground Color
'     nColorET0 = Entry Text Color
'     nColorEB0 = Entry Back Color
'     nColorHT0 = Highlight Text Color 0
'     nColorHB0 = Highlight Back Colot 0
'     nColorHT1 = Highlight Text Color 1
'     nColorHB1 = Highlight Back Colot 1
'     nColorHT2 = Highlight Text Color 2
'     nColorHB2 = Highlight Back Colot 2
'     nColorHT3 = Highlight Text Color 3
'     nColorHB3 = Highlight Back Colot 3
'     nColorHT4 = Highlight Text Color 4
'     nColorHB4 = Highlight Back Colot 4
'     nColorBC0 = Back Color 0
'     nColorBC1 = Back Color 1
'     nColorBC2 = Back Color 2
'     nColorTC0 = Text Color 0
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'04.03.2018    02:57pm     KaLYPTuS
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'  + Added Encrypt, and Decrypt function(s) to the clsAppDriver object.
'  + Added HexToString, StringToHex function(s) to the clsMainModules object.

