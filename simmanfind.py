'''
Looks at all dsb files in project folder, finds all buildings in each.
Finds the sim files on the C:\ within the base install folder.
Brings them into the project folder.

'''

import os
import glob
import sqlite3
import shutil
import re
import xlwings as xw


def simmanfind(eso=False):
    '''
    User runs function.
    optional: ESO.
    Prefers sql output in order to facilitate reports outputting.
    '''
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
        z = re.sub('dsb','',y)
        cleanbnamelist.append(z)
        
    mydict = dict(zip(cleanbnamelist,simnumlist))
    
    
    for simname, simdir in mydict.items():   
        print ('Starting to copy files from ' + str(simdir) + " / " + str(simname))
        csvout = "C:/ProgramData/DesignBuilder/JobServer/Users/User/" + str(simdir) + "/" + "eplustbl.csv"
        if eso:
            esoout = "C:/ProgramData/DesignBuilder/JobServer/Users/User/" + str(simdir) + "/" + "eplusout.eso"
        
        idfout = "C:/ProgramData/DesignBuilder/JobServer/Users/User/" + str(simdir) + "/" + "in.idf"
        #rddfile = "C:/ProgramData/DesignBuilder/JobServer/Users/User/" + str(simdir) + "/" + "eplusout.rdd"
        
        try:
            sqlout = "C:/ProgramData/DesignBuilder/JobServer/Users/User/" + str(simdir) + "/" + "eplusout.sql"
            sqlexcept = False
            
        except:
            sqlexcept = True
            print ('sql output not requested. consider requesting from e+ run...')
            input()
            
        errfile = "C:/ProgramData/DesignBuilder/JobServer/Users/User/" + str(simdir) + "/" + "eplusout.err"
    	  	
        shutil.copy(csvout,"_"+simname+".csv")
        if eso:
            shutil.copy(esoout,"_"+simname+".eso")
        shutil.copy(idfout,"_"+simname+".idf")
        #shutil.copy(rddfile,"_"+simname+".rdd")
        if not sqlexcept:
            shutil.copy(sqlout,"_"+simname+".sql")
            
        shutil.copy(errfile,"_"+simname+".err")
        
    	
        wb = xw.Book(("_"+simname+".csv"))
        wb.close() 
    
        with open(errfile, 'r') as f:
            err = f.readlines()
    	
        print('\n'.join(map(str, err)))
        print ('\nSummary of Errors and Warnings Above...')



if __name__ == "__main__":
    simmanfind()
    