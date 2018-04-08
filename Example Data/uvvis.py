#-----------------------------------MODULE CHECKS------------------------------

# Check for modules, try to exit gracefully if not found
import sys
import imp
try:
    imp.find_module('numpy')
    foundnp = True
except ImportError:
    foundnp = False
try:
    imp.find_module('matplotlib')
    foundplot = True
except ImportError:
    foundplot = False
try:
    imp.find_module('pandas')
    foundpd = True
except ImportError:
    foundplot = False
if not foundnp:
    print("Numpy is required. Exiting")
    sys.exit()
if not foundplot:
    print("Matplotlib is required. Exiting")
    sys.exit()
if not foundpd:
    print("Pandas is required. Exiting")
    sys.exit()

#-------------------------------------------------------------------------------

import os
import glob

#Stop message from appearing
import warnings
warnings.filterwarnings("ignore",".*GUI is implemented.*")

#Find relevant CSVs in folder
path = os.getcwd()
extension = 'csv'
os.chdir(path)
csvresult = [i for i in glob.glob('*.{}'.format(extension))]
print("Plotting the following:")
print(csvresult)

import numpy as np
import matplotlib.pyplot as plt
#import matplotlib.patches as patches
import pandas as pd

#Make x-axis
t = np.linspace(325, 1100, 776)
#define color for lines
color_code=['b','r','g','k','c','m', 'y', 'lime', 'crimson', 'teal', 'aqua']
#define determinant calculation
def det(a, b):
    return a[0] * b[1] - a[1] * b[0]

#set initial values (for later use)
k=0
horiz=.04
calcexport=[]
dataexport=[]
#iterate through CSVs
while k < len(csvresult):
    exportcalc=[]
    exportdata=[]
    path = os.getcwd()
    path = path + "/" + csvresult[k]
    uvvis=pd.read_csv(path, delimiter=",", skiprows = 16)
    
    #Extract Absorbance column
    abscoltemp=uvvis.values[:,1]
    #Convert all elements from strings to floats so they can be math. manipulated
    abscol = [ float(x) for x in abscoltemp ]
    exportdata.append('Abs')
    exportdata.append(abscol)
    #Make normalized abs. data
    x=max(abscol)
    normabscol = [ i / x for i in abscol ]
    exportdata.append('Norm Abs')
    exportdata.append(normabscol)
    #assign attributes to plot
    colour = color_code[k]
    plotlabel = csvresult[k]
    plotlabel = plotlabel[:-4]
    calcexport.append(plotlabel)
    dataexport.append(plotlabel)
    dataexport.append(exportdata)
    #plot graph
    plt.plot(t, normabscol, colour, label= plotlabel) # plotting t, a separately
    plt.draw()
    
    #Point Clicks and Intersections
    print('>> Please choose two points for first line')
    line1 = plt.ginput(2) # it will wait for two clicks
    print('>> Please choose two points for second line')
    line2 = plt.ginput(2)
    #Find intersection
    xdiff = (line1[0][0] - line1[1][0], line2[0][0] - line2[1][0])
    ydiff = (line1[0][1] - line1[1][1], line2[0][1] - line2[1][1])
    div = det(xdiff, ydiff)
        #if div == 0:
           #raise Exception('lines do not intersect')
    d = (det(*line1), det(*line2))
    
    #calculate onset
    x = det(d, xdiff) / div
    exportcalc.append(x)
    #calculate band gap
    bndgp = 1240/x
    exportcalc.append(bndgp)
    #round these values
    x = round(x,2)
    bndgp = round(bndgp,2)
    #make these strings instead of integers
    x = str(x)
    bndgp = str(bndgp)
    #print message on command prompt
    print('------------')
    print("Absorbance onset:")
    print(x + " nm")
    print('------------')

    #find location of max absorbance
    indx=np.argmax(abscol)
    #use location to find max lambda
    lmbdamax=t[indx]
    #fill excel export lists
    exportcalc.append(lmbdamax)
    calcexport.append(exportcalc)
    lmbdamax = str(lmbdamax)
    #write texts to put on chart
    lmbdamaxtxt = '$\lambda$' + "$_{max}$" + ' - ' + lmbdamax + 'nm'
    onsettxt = '$\epsilon$' '$_{onset}$' + ' - ' + x + "nm"
    bandtxt = 'Band Gap - ' + bndgp + "eV"
    #change initial value and make first text label
    plt.gca().set_position((.1, .35, .8, .6)) # to make a bit of room for extra text
    vert = 0.2
    plt.figtext(horiz,vert,plotlabel,style='italic')
    #plt.figure().add_subplot(111).plot(range(10), range(10))
    txtlist = [lmbdamaxtxt,onsettxt,bandtxt]
    i=0
    #iterate through text list to make text boxes of values
    while i < len(txtlist):
        vert=vert-0.05
        curr=txtlist[i]
        plt.figtext(horiz,vert,curr)
        i=i+1
        continue
    k=k+1
    horiz = horiz + 0.25
    #case where if you reach your last CSV, then break the loop
    if k == len(csvresult):
        plt.title('UV/VIS',weight='bold')
        plt.legend(loc='best')
        break
    continue
#----------------------------------EXCEL EXPORT---------------------------------

#---------> CALCULATIONS DATA

# Create a Pandas dataframe title for data.
df0 = pd.DataFrame({'Data': ['Onset of Absorbance (nm)', 'Band Gap (eV)','Lambda Max (nm)']})

# Create a Pandas Excel writer using XlsxWriter as the engine.
folder = 'Processed UVVIS Data'
if not os.path.exists(folder):
    os.makedirs(folder)
writer = pd.ExcelWriter(os.path.join(folder,'uvviscalcs.xlsx'), engine='xlsxwriter')

#Loop through data from total export list
j=0
index=1
while j < len(calcexport):
    plotlabel=calcexport[j]
    data=calcexport[j+1]
    df = pd.DataFrame({plotlabel: data})
    df.to_excel(writer, sheet_name='Sheet1', startcol=index, index=False)
    j=j+2
    index=index+1
    continue

# Convert the dataframe to an XlsxWriter Excel object.
df0.to_excel(writer, sheet_name='Sheet1', startrow=1, header = False, index=False)

# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['Sheet1']

# Set the column width and format.
worksheet.set_column('A:A', 20)
worksheet.set_column('B:B', 15)
worksheet.set_column('C:C', 15)
worksheet.set_column('D:D', 15)

# Close the Pandas Excel writer and output the Excel file.
writer.save()
print('UVVIS Excel Calculations exported ---->')

#---------> DATA TO CREATE GRAPH IN EXCEL

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter(os.path.join(folder,'uvvisdataset.xlsx'), engine='xlsxwriter')

#Create wavelength column
df = pd.DataFrame({'Wavelength': t})
df.to_excel(writer, sheet_name='Sheet1', startcol=0, startrow=1, index=False)

#Loop through data from total export list
i=0
j=1
colindex=1
titlindex=1
while i < len(dataexport):
    k=0
    #Loop through to make abs and norma abs columns
    while k < 3:
        plotlabel=dataexport[j][k]
        plotdata=dataexport[j][k+1]
        df = pd.DataFrame({plotlabel: plotdata})
        df.to_excel(writer, sheet_name='Sheet1', startcol=colindex, startrow=1, index=False)
        colindex=colindex+1
        k=k+2
        continue
    #Write name of compound
    titlelabel=dataexport[i]
    title=pd.DataFrame({titlelabel: []})
    title.to_excel(writer, sheet_name='Sheet1', startcol=titlindex, startrow=0, index=False)
    i=i+2
    j=j+2
    titlindex=titlindex+2
    continue


# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['Sheet1']

# Set the column width and format.
worksheet.set_column('A:A', 11)
worksheet.set_column('B:B', 8)
worksheet.set_column('C:C', 11)
worksheet.set_column('D:D', 8)
worksheet.set_column('E:E', 11)

# Close the Pandas Excel writer and output the Excel file.
writer.save()
print('UVVIS Excel Dataset exported --------->')

#-------------------------------------------------------------------------------

#Bold axis numbers and change font sizes
ax=plt.gca()
for tick in ax.xaxis.get_major_ticks():
    tick.label1.set_fontsize(9)
    tick.label1.set_fontweight('bold')
for tick in ax.yaxis.get_major_ticks():
    tick.label1.set_fontsize(9)
    tick.label1.set_fontweight('bold')
#Axis Ranges
plt.xlim([325,1100])
plt.ylim([0.0, 1.1])
#Labels
plt.xlabel('Wavelength ($\lambda$) / nm',weight='bold')
plt.ylabel('Normalized Absorbance ($\epsilon$)',weight='bold')
plt.draw()
#Graph Finished Message
sep=" "
name = os.getlogin()
name = name.split(sep, 1)[0]
msg = 'Hey ' + name + ', your graph has finished processing.'
print(msg)
plt.savefig(os.path.join(folder,'UVVIS.png'), bbox_inches='tight')
plt.show()
