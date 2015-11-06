"""
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Silas Toms, 1/10/2010 GIS Technician, East Bay Regional Parks District.
This s a GUI based project that converts shapefiles into KMLs.
It uses a SpatiaLite backend and a Tkinter frontend, tied together using
Python 2.6.4. GUI2Exe was used to convert the script into an executable.
This script depends on the initsql script for important data. Make sure
to download that script and put it on your path or next to the script.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
"""

import os, base64, time   
from Tkinter import *
import tkFileDialog, tkColorChooser
from pysqlite2 import dbapi2 as sql
from osgeo import ogr
from initsql import pushpin, polystyle, closekml, smallpin, pinkml, polykml
from initsql import leaf, linekml, linestyle
from dbf.dbfreader import dbfreader 

import psyco
psyco.full()


# These global variables are used to set the default kml output colors and other settings. They can be adjusted as needed
global polycolor
global linecolor 
global highlight
global icon
global width
global scale
global indicator
global trans
global fill

indicator = 0
linecolor = '00cc66'
polycolor = '008040'
highlight = 'ffff00'
scale = '1.0'
width = '1.0'
outlined = '1'
icon = 'pushpin.png'
trans = 'b2'
fill = '1'
highlight3 = '#ff80ff'

#The following creates the necessary SQLite/SpatiaLite database and tables
if not os.path.exists('Geo.sqlite'):
    from initsql import init
    fo = open('init.sql','w')
    fo.write(init)
    fo.close()
    del fo
    sql_connection = sql.Connection('Geo.sqlite')     # This is the path to the database. If it doesn't exist, it will be created, though, of course, without the require tables
    cursor = sql.Cursor(sql_connection)
    sql_connection.enable_load_extension(1)
    sql_connection.load_extension('libspatialite-2.dll')                 # Here, to be able to programmatically interact with the Spatialite program, an extension must be called. It resides in the C:/WINDOWS/System32 folder, with some associated DLLs that must also be present
    os.system('spatialite Geo.sqlite < init.sql')
    cursor.execute('Create table FavoriteSRID("srid","ref_sys_name")')
    sql_connection.commit()
    del init
    os.remove('init.sql')

# The following creates the connection with the database    
sql_connection = sql.Connection('Geo.sqlite')     # This is the path to the database. If it doesn't exist, it will be created, though, of course, without the require tables
cursor = sql.Cursor(sql_connection)
sql_connection.enable_load_extension(1)
sql_connection.load_extension('libspatialite-2.dll')                 # Here, to be able to programmatically interact with the Spatialite program, an extension must be called. It resides in the C:/WINDOWS/System32 folder, with some associated DLLs that must also be present


    
# This function is used to convert the coordinates from a Well Known Text format into a KML-ready format
def coordinates(poly):
    multi = poly.split('), (')
    multilist = []
    for single in multi:
        poly = single.replace('MULTI','')
        poly = poly.replace('LINESTRING','')
        poly = poly.replace('POINT','')
        poly = poly.replace('(((','')
        poly = poly.replace(')))','')
        poly = poly.replace('POLYGON','')
        poly = poly.replace(')','')
        poly = poly.replace('(','')
        poly = poly.replace('((','')
        poly = poly.replace('))','')
        coords = ''       
        polylist = poly.split(',')
        for coord in polylist:
                coord = coord.split()
                longitude = coord[0]
                latitude = coord[1]
                coords = coords + str(longitude) + ',' + str(latitude) + ',0 '
        multilist.append(coords)
    return multilist


#This function uses the tkFileDialog to select an icon for point data. The default is the pushpin.png file
def symbolchoose():
    global master2
    global icon
    symbol = tkFileDialog.askopenfile(parent=master2,
                                      mode='rb',
                                      title='Choose a image file',
                                      filetypes=[('Image Files','*.png *.jpg *.bmp *.ico *.gif')])
    try:
        icon = symbol.name
    except:
        pass

# Chose the fill color for polygons
def colorchoose():
    global polycolor
    #rgb_tuple = tkColorChooser.askcolor()[0]
    polycolor = tkColorChooser.askcolor()[1]
    polycolor = polycolor.strip('#')

# Chose the outline/line color
def colorOfLine():
    global linecolor
##    rgb_tuple = tkColorChooser.askcolor()[0]
    linecolor = tkColorChooser.askcolor()[1]
    linecolor = linecolor.strip('#')

# Chose the highlight color
def colorOfHighlight():
    global highlight
    
##    rgb_tuple = tkColorChooser.askcolor()[0]
    highlight = tkColorChooser.askcolor()[1]
    highlight = highlight.strip('#')

# links to the EPSG site listing all srids
def sridhelp():
    try:
        os.startfile('http://spatialreference.org/')
    except:
        pass
# Wikipedia's Shapefile description
def shapehelp():
    try:
        os.startfile('http://en.wikipedia.org/wiki/Shapefile')
    except:
        pass

# Google's KML description
def kmlhelp():
    try:
        os.startfile('http://code.google.com/apis/kml/documentation/whatiskml.html')
    except:
        pass

# Window README text
def credits():
    about = Tk()
    about.geometry('+400+100')
    about.title('KML Kreator README')
    about.wm_attributes("-topmost", 1)
    about.iconbitmap('leaf.ico')
    credits_label = Label(about,
                          text=''''KML Konverter Program
By Silas Toms. 1.10.2010
GIS Technician, Land Acquisition,
East Bay Regional Parks District.
---- stoms@ebparks.org ----
This program converts shapefiles
into Google Earth KMLs using a
SpatiaLite® database backend with
a Python® and Tkinter® GUI frontend.
The script was converted to an
executable using GUI2Exe v0.3.
SpatiaLite Version = 2.3.1
Python version = 2.6.4''')
    credits_label.pack()

# This function allows the user to reorganize the list of SRIDs (ORDER by SRID)
def sridatize():
    global listbox

    cursor.execute('select srid, ref_sys_name from FavoriteSRID')
    favsrids = cursor.fetchall()
    favcounter = 0
    for srid in favsrids:
        favcounter += 1

    cursor.execute('select srid, ref_sys_name from spatial_ref_sys')
    sridlist = cursor.fetchall()
    sridcounter = 0
    for srid in sridlist:
        sridcounter+=1
    del sridlist, favsrids
    listbox.delete(favcounter+3, sridcounter+5)
    cursor.execute('select srid, ref_sys_name from spatial_ref_sys ORDER BY srid')
    sridandname = cursor.fetchall()
    
    for srid in sridandname:
                id = str(srid[0])
                name = srid[1]
                listitem = ' ' + id +'                       ' + name
                listbox.insert(END,listitem)
    del sridandname

# This function alphabetizes the list of SRIDs
def alphabetize():
    global listbox

    cursor.execute('select srid, ref_sys_name from FavoriteSRID')
    favsrids = cursor.fetchall()
    favcounter = 0
    for srid in favsrids:
        favcounter += 1

    cursor.execute('select srid, ref_sys_name from spatial_ref_sys')
    sridlist = cursor.fetchall()
    sridcounter = 0
    for srid in sridlist:
        sridcounter+=1
    del sridlist, favsrids
    listbox.delete(favcounter+3, sridcounter+5)
    cursor.execute('select srid, ref_sys_name from spatial_ref_sys ORDER BY ref_sys_name')
    sridandname = cursor.fetchall()
    
    for srid in sridandname:
                id = str(srid[0])
                name = srid[1]

                listitem = ' ' + id +'                       ' + name
                listbox.insert(END,listitem)
    del sridandname

# This allows the user to clear an entry from the list of favorite SRIDs       
def clearfav():
    global listbox
    global spatialref
    try:
        cursor.execute('select srid, ref_sys_name from FavoriteSRID')
        favsrids = cursor.fetchall()
        favcounter = 0
        for srid in favsrids:   # Count the SRIDs that the User has designated 'Favorites'
            favcounter += 1
        favcounter = favcounter +3
        spatialrefnum = listbox.curselection()[0]  #'Get' the index number of the selected SRID
        spatialref = listbox.get(spatialrefnum)
        spatialref = spatialref.split()[0]
        if spatialrefnum != '0':
            if spatialref != 'Reference':
                if spatialref !='-----------Favorites---------':
                    if int(spatialrefnum) < favcounter and spatialref != "-----------SRID":
                        cursor.execute('DELETE FROM FavoriteSRID WHERE srid LIKE %s'% spatialref)
                        sql_connection.commit()
                        listbox.delete(ANCHOR)
    except:
        pass

# This function allows the User to 'save' an SRID to the list of Favorites, making it easy to access
def addfav():
        global listbox
        try:
            spatialrefnum = listbox.curselection()[0]  #'Get' the index number of the selected SRID
            global spatialref
            if spatialrefnum != '0':
                spatialref = listbox.get(spatialrefnum)
                spatialref = spatialref.split()[0]
                cursor.execute('select srid, ref_sys_name from spatial_ref_sys WHERE srid LIKE %s' % spatialref)
                sridandname = cursor.fetchone()
                srid = str(sridandname[0])
                name = sridandname[1]
                cursor.execute('select srid, ref_sys_name from FavoriteSRID WHERE srid LIKE %s' % spatialref)
                sridandname = cursor.fetchone()
                if sridandname == None:
                    cursor.execute('INSERT INTO FavoriteSRID VALUES("%s","%s")' % (srid,name))
                    sql_connection.commit()
                    listitem = ' ' + srid +'                       ' + name
                    listbox.insert(2,listitem)
        except:
            pass

#This function is where the shapefile to be converted is selected
def shapegrab():
    shape = tkFileDialog.askopenfile(parent=root,mode='rb',
                                     title='Choose a shapefile',
                                     filetypes=[('Shapefiles','*.shp')])
    try:
        
        global name
        name = shape.name
        global nameorig
        nameorig = name
        name = name.rsplit('.',1)[0]
        global tablename
        tablename = name.rsplit('/',1)[1]
        labelpath.set(tablename)
        masterWindow()
    except:
        pass

# This function determines the feature type of the shapefile, i.e. point/line/polygon/multipolygon (an ogr Geometry Type of '6' == SpatiaLite multipolygon)
def getFeatureType(shp):
    shp =  str(shp)
    open = ogr.Open(shp)
    shplyr = shp.split('.')[0].split('/')[-1]
    lyr = open.GetLayerByName(shplyr)
    lyr.ResetReading()
    f = lyr.GetNextFeature()
    geom = f.GetGeometryRef()
    geom_type = geom.GetGeometryType()
    global type
    dic = {1:'POINT', 2:'LINESTRING', 3:'POLYGON', 6:'POLYGON'}
    del open, f, geom, lyr, shp
    type = dic[geom_type]
    return type
  

# This is the second window of the GUI; it houses the list of SRIDs
def masterWindow():
    global root
    root.destroy()
    
    global master
    global listbox
    
    master = Tk()
    master.wm_attributes("-topmost", 1)
    master.title('Select Spatial Reference')
    master.maxsize(500,500)
    master.geometry('+100+100')
    master.iconbitmap('leaf.ico')
    menubar = Menu(master)
    menubar.add_command(label="SRID Help", command=sridhelp)
    menubar.add_separator()
    menubar.add_command(label="Add To Favorites", command=addfav)
    menubar.add_command(label="Clear From Favorites", command=clearfav)
    menubar.add_separator()
    #menubar.add_command(label="Alphabetical SRIDs", command=alphabetize)
    filemenu = Menu(menubar, tearoff=0)
    filemenu.add_command(label="Alphabetical SRIDs", command=alphabetize)
    filemenu.add_command(label="Order By SRIDs", command=sridatize)
    menubar.add_cascade(label ='Change Case', menu=filemenu)
    master.config(menu=menubar)

    frame2 = Frame(master, height=6, bd=7, relief=SUNKEN)
    frame2.pack(fill=X, padx= 2, pady=2)

    frame5 = Frame(frame2, bd=7)
    frame5.pack()
    
    scrollbar = Scrollbar(frame5)
    scrollbar.pack(side=RIGHT, fill= Y)

    listbox = Listbox(frame5,height = 20, width= 400,
                      yscrollcommand=scrollbar.set,
                      xscrollcommand=scrollbar.set)
    listbox.insert('active', " Reference Code     Reference System")
    cursor.execute('select srid, ref_sys_name from FavoriteSRID')
    sridfavs = cursor.fetchall()
    listbox.insert(1, "-----------Favorites---------")
    if sridfavs != []:
        for item in sridfavs:
            item = ' ' + str(item[0]) +'                       ' + item[1]
            listbox.insert(END, item)
    cursor.execute('select srid, ref_sys_name from spatial_ref_sys')
    srids = cursor.fetchall()
    listbox.insert(END, "-----------SRID List---------")
    for item in srids:
        item = ' ' + str(item[0]) +'                       ' + item[1]
        listbox.insert(END, item)
    listbox.pack()
    scrollbar.config(command=listbox.yview)
    stepButton = Button(frame5, text= 'Next Step',
                         font = 'Gill_Sans_MT -12 bold' ,
                         bg = 'dark green', fg = 'white',
                         bd = 10, width=12,
                         command=nextStep)
    stepButton.pack(side=TOP, pady =5,padx=4)
    master.mainloop()


#this is an intermediate function that links the masterWindow function with the LastGUI function; it checks if the User has selected an SRID
def nextStep():
    try:
        spatialrefnum = listbox.curselection()[0]
        global spatialref
        spatialref = listbox.get(spatialrefnum)
        spatialref = spatialref.split()[0]
        if spatialrefnum != '0' and spatialrefnum != '1' and spatialref != "-----------SRID":
            master.destroy()
            LastGUI()
    except:
        pass

#This function generates a GUI that can be expanded to allow the user to adjust the labeling, scaling and coloring defaults
def LastGUI():
    global master2
    global type
    global widthscale
    global transcale
    global highlight
    global filelabels
    global popups
    global dbfheaders
    global usefields
    
    master2 = Tk()
    master2.wm_attributes("-topmost", 1)
    master2.title('Create KML')
    master2.minsize(300,100)
    master2.geometry('+100+100')
    master2.iconbitmap('leaf.ico')
    type  = getFeatureType(nameorig)

    frame2 = Frame(master2,height=6, bd=7, width = 20)
    fieldframe = Frame(master2,height=6, bd=7, width = 20)
    separator = Frame(height=2, bd=1, relief=SUNKEN)
    separator2 = Frame(height=2, bd=1, relief=SUNKEN)
    fieldframe1 = Frame(fieldframe,height=6, bd=7, width = 20)
    fieldframe2 = Frame(fieldframe,height=6, bd=7, width = 20)
    def changeframe():
        frame2.pack(side = BOTTOM)
        separator.pack(fill=X, padx=5, pady=5, side=BOTTOM) 
    def fields():
        fieldframe.pack(side=BOTTOM)
        separator2.pack(fill=X, padx=5, pady=5, side=BOTTOM)

        fieldframe1.pack()

        fieldframe2.pack()

        scrollbar.pack(side=RIGHT)        
    menubar = Menu(master2)
    menubar.add_command(command = changeframe, label = 'Change Defaults')
    menubar.add_command(command = fields, label = 'Adjust Labeling')
    master2.config(menu=menubar) 
    dbfile = open(name+'.dbf','rb')
    dbfheaders = list(dbfreader(dbfile))[0]
    dbfile.close()   
  
    if type != "POINT":
        if type == "POLYGON":
            colorbutton = Button(frame2, text = 'Fill Color',
                                 font = 'Gill_Sans_MT -12 bold' ,
                                 bg = 'dark green', fg = 'white',
                                 bd = 5, width = 15,height = 1,
                                 command=colorchoose)
            colorbutton.pack(pady = 3)

        linecolorbutton = Button(frame2, text = 'Line Color',
                             font = 'Gill_Sans_MT -12 bold' ,
                             bg = 'dark green', fg = 'white',
                             bd = 5, width = 15,height = 1,
                             command=colorOfLine)
        linecolorbutton.pack(pady = 3)
        highlightbutton = Button(frame2, text = 'Highight Color',
                             font = 'Gill_Sans_MT -12 bold' ,
                             bg = 'dark green', fg = 'white',
                             bd = 5, width = 15,height = 1,
                             command=colorOfHighlight)
        highlightbutton.pack(pady = 3)
        transcale = Scale(frame2,label = 'Transparency', from_=10, to=100, resolution=5, orient=HORIZONTAL)
        transcale.set(70)
        transcale.pack()

        widthscale = Scale(frame2,label = 'Line Width', from_=0, to=10, resolution=.1, orient=HORIZONTAL)
        widthscale.set(2)
        widthscale.pack()
        

    else:
        imagebutton = Button(frame2, text = 'Choose Image',
                             font = 'Gill_Sans_MT -12 bold' ,
                             bg = 'dark green', fg = 'white',
                             bd = 5, width = 12,height = 1,
                             command=symbolchoose)
        imagebutton.pack()
        widthscale = Scale(frame2,
                           label='Scale of Icon',
                           from_=1, to=7,
                           resolution=.2,
                           orient=HORIZONTAL)
        widthscale.set(1)
        widthscale.pack()
      
    scrollbar1 = Scrollbar(fieldframe1)
    
    filelabels = Listbox(fieldframe1,exportselection= 0,
                         height =3, width=25,
                         yscrollcommand=scrollbar1.set)
    scrollbar1.config(command=filelabels.yview)
    scrollbar1.pack(side=RIGHT, fill=Y)
 
    for field in dbfheaders:
        filelabels.insert(END, field)
    filelabels.pack()
    filelabels.focus_set()
    la= StringVar()
    la.set('Select Label for Features')
    Label (fieldframe1, textvariable=la, font = 'Helvetica -11 bold').pack()
    
    scrollbar = Scrollbar(fieldframe2)


    popups = Listbox(fieldframe2,exportselection= 0,
                     height =5, width=25,selectmode=MULTIPLE,
                     yscrollcommand=scrollbar.set)
    scrollbar.config(command=popups.yview)
    scrollbar.pack(side=RIGHT, fill=Y)
    for field in dbfheaders:
        popups.insert(END, field)
    popups.pack()
    popups.focus_set()
    la= StringVar()
    la.set('Select Pop-Up Fields')
    Label (fieldframe2, textvariable=la, font = 'Helvetica -11 bold').pack()        
    
    kmlButton = Button(master2, text='Create KML',
                         font = 'Gill_Sans_MT -18 bold' ,
                         bg = 'dark green', fg = 'white',
                         bd = 7, width = 15,height = 3,
                         command=kmlcreate)
    kmlButton.pack(side=TOP, pady = 10,padx=4)


#This is the meat of the script; it adds the shapefile to a SpatiaLite table as a virtual table; then creates a table
#to which the relevant data is copied. This could be adjusted to only copy specified records if desired. The tables
#are dropped at the end of the function so that the database does not get too large.
def kmlcreate():
    global tablename
    global name
    global spatialref
    global nameorig
    global type
    global polycolor
    global scale
    global linecolor
    global widthscale
    global transcale
    global width
    global highlight
    global fill
    global filelabels
    global popups
    global dbfheaders


    try:

        newtable = tablename + '_vtable'

        virtual_table = "Create VIRTUAL TABLE '%s' USING VirtualShape('%s','%s',%s)" %\
                        (tablename, name, 'CP1252', str(spatialref) )

        cursor.execute(virtual_table)
        columns = cursor.execute('PRAGMA table_info(%s)' % tablename)
        columnames = ''

        columnlist = []
        column_set = columns.fetchall()
        for column in column_set:
            columnlist.append(column[1])
            if column[1] != 'Geometry':
                columnames = columnames + '"'+column[1] + '"' + ', '

        columnames = columnames.rsplit(',',1)[0]
        select_sql = "CREATE TABLE '%s'(%s) "% (newtable, columnames)
        
        cursor.execute(select_sql)
        feature_type  = getFeatureType(nameorig)
        addgeom = "SELECT AddGeometryColumn('%s', 'Geometry', %s,'%s',2)" %\
                  (newtable, str(spatialref), feature_type)

        cursor.execute(addgeom)
        sql_connection.commit()
        columnameswithgeom = ''
        for column in columnlist:
            columnameswithgeom = columnameswithgeom + '"'+ column + '"'+ ', '
        columnameswithgeom = columnameswithgeom.rsplit(',',1)[0]

        insert_sql = 'INSERT INTO "%s"(%s) SELECT %s FROM "%s"' %\
                     (newtable, columnameswithgeom,columnameswithgeom, tablename)

        cursor.execute(insert_sql)
        sql_connection.commit()

    except:

        newtable = tablename + '_vtable'
        droptable = 'DROP TABLE %s' % tablename
        cursor.execute(droptable)
        droptable = 'DROP TABLE %s' % newtable
        cursor.execute(droptable)
        sql_connection.commit()

        virtual_table = "Create VIRTUAL TABLE '%s' USING VirtualShape('%s','%s',%s)" %\
                        (tablename, name, 'CP1252', str(spatialref) )

        cursor.execute(virtual_table)
        columns = cursor.execute('PRAGMA table_info(%s)' % tablename)
        columnames = ''

        columnlist = []
        column_set = columns.fetchall()
        for column in column_set:
            columnlist.append(column[1])
            if column[1] != 'Geometry':
                columnames = columnames + '"'+column[1] + '"' + ', '

        columnames = columnames.rsplit(',',1)[0]
        select_sql = "CREATE TABLE '%s'(%s) "% (newtable, columnames)
        
        cursor.execute(select_sql)
        addgeom = "SELECT AddGeometryColumn('%s', 'Geometry', %s,'%s',2)" %\
                  (newtable, str(spatialref), 'MULTIPOLYGON')

        cursor.execute(addgeom)
        sql_connection.commit()
        columnameswithgeom = ''

        for column in columnlist:
            columnameswithgeom = columnameswithgeom + '"'+ column + '"'+ ', '
    
        columnameswithgeom = columnameswithgeom.rsplit(',',1)[0]

        insert_sql = 'INSERT INTO "%s"(%s) SELECT %s FROM "%s"' %\
                     (newtable, columnameswithgeom,columnameswithgeom, tablename)
        
        cursor.execute(insert_sql)
        sql_connection.commit()
    
    popupfieldslist = []
    columnames = ''
    for num in popups.curselection():
        columnheader = popups.get(num).title()
        popupfieldslist.append(columnheader)
        columnames = columnames + columnheader + ','

    columnames = columnames.rsplit(',',1)[0]

    if columnames =='':
        for header in dbfheaders:
            columnames = columnames + header + ', '
        columnames = columnames.rsplit(',',1)[0]

    featurelabel = ''    
    for num in filelabels.curselection():
            featurelabel = filelabels.get(num)


    grab_geometry = 'SELECT AsText(Transform(Geometry, 4326)), %s FROM %s' %\
                   (columnames, newtable)
    cursor.execute(grab_geometry)

    kmlname = tkFileDialog.asksaveasfilename(filetypes=[('Google Earth KML','*.kml')])
    kmlname = kmlname.rsplit('.')[0]
    kmlfileopen = open(kmlname +'.kml', 'w')
    if type != "POINT":
        trans = transcale.get()
        trans = int((trans * .01) * 255)
        trans = '%02x' % trans
        polycolor = trans + polycolor
        linecolor = 'ff' + linecolor
        highlight = 'ff' + highlight
        width = widthscale.get()
        if type == 'LINESTRING':
            styling = linestyle % (linecolor, width, highlight, width)
            #styling = linestyle % (name, linecolor, width, highlight, width)
        else:
            styling = polystyle % (kmlname,highlight,polycolor,fill,linecolor, width, polycolor, fill)

    else:
        scale = widthscale.get()
        styling = smallpin % (kmlname, scale, icon, scale, scale, icon, scale)
    
    kmlfileopen.write(styling)
    kmlfileopen.close()

    shpdata = cursor.fetchone()
    while shpdata != None:
        geometry = shpdata[0]
        
        labelname = ''
        datastring = ''
        DATACOUNTER = 1

        if popupfieldslist == []:
            popupfieldslist = dbfheaders
        for columnheader in popupfieldslist:
            if columnheader.title() == featurelabel.title():
                labelname = str(shpdata[DATACOUNTER])
                labelname = labelname.replace('&','and')
                labename = labelname.title()
            data = str(shpdata[DATACOUNTER])
            data = data.replace('&','and')
            columnheader = columnheader.title()
            data = data.title()
            datastring =  columnheader + ' = '+ data+ '''
''' +  datastring
            DATACOUNTER += 1
        if labelname=='':
            labelname = kmlname.rsplit('/')[-1]
            
        geomlist = coordinates(geometry)
        for geom in geomlist:
            kmlfileopen = open(kmlname+'.kml','a')


            if type != 'POINT':
                if type != 'POLYGON':
                    kmlcoords = linekml % (labelname, datastring, geom)
                else:
                    kmlcoords = polykml % (labelname, datastring, geom)
            
            else:
                kmlcoords = pinkml % (datastring, geom)
            kmlfileopen.write(kmlcoords)

        shpdata = cursor.fetchone()
    kmlfileopen.write(closekml)
    kmlfileopen.close()
    os.startfile(kmlname+'.kml')
    master2.destroy()
    cursor.execute('DROP Table %s' % tablename)
    cursor.execute('DROP Table %s' % (tablename + '_vtable'))
    sql_connection.commit()


#The next two sections create images for use by the kmls and the program itself
ico1 = base64.b64decode(pushpin)   #decode the string
temp_file = "pushpin.png"         #create a temporary file
fout = open(temp_file,"wb")      # open the file in binary format
fout.write(ico1)                 #write the converted data to the temp file
fout.close()                     # make sure to close the file!!!!
del pushpin



ico1 = base64.b64decode(leaf)   #decode the string
temp_file = "leaf.ico"         #create a temporary file
fout = open(temp_file,"wb")      # open the file in binary format
fout.write(ico1)                 #write the converted data to the temp file
fout.close()                     # make sure to close the file!!!!
del leaf



    
# This is the base of the script; 
global root
root =Tk()
root.geometry('+100+100')
root.title('KML Konverter')
#root.maxsize(400,700)
root.wm_attributes("-topmost", 1)  
root.iconbitmap('leaf.ico')
frame = Frame(height=6, bd=7, relief=SUNKEN)
frame.pack(fill=X, padx= 2, pady=2)

menubar = Menu(root)
filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(command = shapehelp, label = 'Shapefile Description')
filemenu.add_command(label="KML Description", command=kmlhelp)
menubar.add_cascade(label ='SHP/KML Info', menu=filemenu)
menubar.add_command(command = credits, label = 'Program Info')

root.config(menu=menubar) 
frame2 = Frame(frame,height=6, bd=7)
frame2.pack()
labelpath = StringVar()
labelpath.set('''Select Shapefile to Start''')
Label (frame2, textvariable=labelpath, font = 'Helvetica -20 bold').pack(side=TOP)

boundButton = Button(frame, text='Select Shapefile',
                     font = 'Gill_Sans_MT -12 bold' ,
                     bg = 'dark green', fg = 'white',
                     bd = 10, width=20, height=2,
                     command=shapegrab)
boundButton.pack(pady =5)

root.mainloop()




