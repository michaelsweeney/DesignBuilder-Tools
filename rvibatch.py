import os
import glob
import sqlite3
import shutil
import re
import xlwings as xw

# assign each dsb and building within dsb to its respective sim manager number
filelist = []
bnamelist = []
simnumlist = []

for file in glob.glob("*.dsb"):
    filelist.append(file)

for name in filelist:
    conn = sqlite3.connect("C:/ProgramData/DesignBuilder/JobServer/DBJobServer.db")
    c = conn.cursor()
    c.execute("SELECT * FROM job_instances")
    rows = c.fetchall()

    for row in rows:
        if name in row[4]:
            simnum = row[2]
            simnumlist.append(simnum)
            bname = row[4]
            bnamelist.append(bname)

cleanbnamelist = []

for x in bnamelist:
    y = re.sub('[^A-Za-z0-9]+', '', x)
    z = re.sub('dsb', '', y)
    cleanbnamelist.append(z)

mydict = dict(zip(cleanbnamelist, simnumlist))

# make directories and copy E+ files from sim manager folder to project folder
idflist = []
for simname, simdir in mydict.items():
    simname = "Outputs_"+simname
    if not os.path.exists(simname):
        os.makedirs(simname)

    csvout = "C:/ProgramData/DesignBuilder/JobServer/Users/User/" + str(simdir) + "/" + "eplustbl.csv"
    esoout = "C:/ProgramData/DesignBuilder/JobServer/Users/User/" + str(simdir) + "/" + "eplusout.eso"
    idfout = "C:/ProgramData/DesignBuilder/JobServer/Users/User/" + str(simdir) + "/" + "in.idf"
    rddfile = "C:/ProgramData/DesignBuilder/JobServer/Users/User/" + str(simdir) + "/" + "eplusout.rdd"
    #sqlout = "C:/ProgramData/DesignBuilder/JobServer/Users/User/" + str(simnumstr) + "/" + "eplusout.sql"

    shutil.copy(csvout, simname+"/"+"_"+simname+".csv")
    shutil.copy(esoout, simname+"/"+"eplusout.eso")
    shutil.copy(idfout, simname+"/"+"in.idf")
    shutil.copy(rddfile, simname+"/"+"eplusout.rdd")
    #shutil.copy(sqlout, simname+"/"+"eplusout.sql")

    wb = xw.Book((simname+"/"+"_"+simname+".csv"))
    wb.close()
    currentidf = simname+"/"+"in.idf"
    idflist.append(currentidf)

# make rvi template file
for i in idflist:

    with open(i) as f:
        hrvarstring = []
        flist = f.readlines()
        for line in flist:
            if "Output:Variable, " in line and "hourly" in line:
                hrvarstring.append(line)
        hrvarlist = []

        for h in hrvarstring:
            splitter = h.split(",")
            hrvarlist.append(splitter)

        rvilist = []
        for m in hrvarlist:
            rvilist.append(m[2])

        rviset = set(rvilist)
        rvilist = rviset
        rviliststrip = []
        for r in rvilist:
            rviliststrip.append(r[1:])
        rvilist = rviliststrip
        rvilist = sorted(rvilist)

        rviname = i.strip("in.idf")+"eplusout.rvi_template"

        f = open(rviname, "w")
        f.write("eplusout.eso")
        f.write("\n")
        f.write(i.strip("/in.idf").strip("_Outputs")+"_HourlyResults.csv")
        f.write("\n")
        f.write("\n".join(str(x) for x in rvilist))
        f.write("\n")
        f.write("0")
        f.close()

# make BAT file
for i in idflist:
    with open(i) as b:
        batname = i.strip("in.idf")+"readvars.bat"
        b = open(batname, "w")
        b.write("@echo off")
        b.write("\n")
        b.write("set post_proc=C:\\EnergyPlusV8-6-0\\PostProcess\\")
        b.write("\n")
        b.write("IF EXIST eplusout.eso DEL eplusout.csv")
        b.write("\n")
        b.write("@echo .")
        b.write("\n")
        b.write("@echo =====  Extracting Results")
        b.write("\n")
        b.write("@echo .")
        b.write("\n")
        b.write("IF EXIST \"eplusout.eso\" %post_proc%ReadVarsESO.exe \"eplusout.rvi\" hourly unlimited")
        b.write("\n")
        b.write("@echo .")
        b.write("\n")
        b.close()