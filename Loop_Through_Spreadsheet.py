#import libraries
import arcpy
import xlrd
import os, itertools

#import modules
from xlrd import open_workbook
from arcpy import env

#set location of spreadsheet
book = xlrd.open_workbook('file:///C:\Users\mensahd\Documents\figure_number_change.xlsx')
sheet = sheet.name("sheet1")

#set location of MXDs
MXDs = r"M:\New_York_City\NYSERDA_Offshore\Maps\MXD\Masterplan_figures\3_Marine_Mammals\Seasonal_Maps"

#create list of MXDs
all_MXDs = [MXDs]

#loops through spreadsheet and obtains figure numbers in column A
for row in range(sheet.nrows):
    if row == 0:
        continue
    Fignum = sheet.cell(row, 1)

    #iterate through the list of MXDs in folder
    for mxd in all_MXDs:

        #setting the input mxd workspace
        arcpy.env.workspace = ws

        #getting all the mxds in the workspace and put in list
        all_MXDs = arcpy.ListFiles("*.mxd")

        print "Opening mxds in CURRENT WORKSPACE: " + str(ws)
        print "mxd_list: " + str(len(mxd_list)) + " items in this folder" + '\n'

        #loop through the list of mxds in the list
        for mxd in all_MXDs:

            #get the file path of the current mxd
            current_mxd = arcpy.mapping.MapDocument(str(ws) + '\\' + str(mxd))
            print str(ws) + '\\' + str(mxd)
            for elm in arcpy.mapping.ListLayoutElements(mxd, Fignum):
                if elm.text == Fignum:
                    elm.text = Fignum
            mxd.save()
        
            #close the mxd
            del mxd
  
print "Done"

    
    
    

