'''
Update unit measurement pgb names
'''
from __future__ import annotations
import os
import sys

if __name__ == '__main__':
    print(sys.version)
    #Ensure right working directory
    executePath = os.path.abspath(__file__)
    executeFolder = os.path.dirname(executePath)
    os.chdir(executeFolder)
    sys.path.append(executeFolder)
    print(executeFolder)

from typing import List

if __name__ == '__main__':
    from execute_pscad import connectPSCAD

import mhi.pscad

def updateUMs(pscad : mhi.pscad):
    projectLst = pscad.projects()
    for prjDic in projectLst:
        if prjDic['type'].lower() == 'case':
            project = pscad.project(prjDic['name'])
            print('Updating unit measurements in project: {}'.format(project))
            ums : List[mhi.pscad.UserCmp]= project.find_all(Name_='$ALIAS_UM_9124$')
            for um in ums:
                print('\t{}'.format(um))
                canvas : mhi.pscad.Canvas = um.canvas()
                umParams = um.parameters()
                alias = umParams['alias']
                pgbs = canvas.find_all('master:pgb')
                for pgb in pgbs:
                    print('\t\t{}'.format(pgb))
                    pgbParams = pgb.parameters()
                    pgb.parameters(Name = alias + '_' + pgbParams['Group'])

def main():
    pscad = connectPSCAD()    
    updateUMs(pscad)

if __name__ == '__main__':
    main()

