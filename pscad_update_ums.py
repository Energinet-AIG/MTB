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

def updateUMs(pscad : mhi.pscad.PSCAD, verbose : bool = False) -> None:
    projectLst = pscad.projects()
    for prjDic in projectLst:
        if prjDic['type'].lower() == 'case':
            project = pscad.project(prjDic['name'])
            print(f'Updating unit measurements in project: {project}')
            ums : List[mhi.pscad.UserCmp]= project.find_all(Name_='$ALIAS_UM_9124$') #type: ignore
            for um in ums:
                print(f'\t{um}')
                canvas : mhi.pscad.Canvas = um.canvas()
                umParams = um.parameters() #type: ignore
                alias = umParams['alias'] #type: ignore
                pgbs = canvas.find_all('master:pgb') #type: ignore
                for pgb in pgbs:
                    if verbose:
                        print(f'\t\t{pgb}')
                    pgbParams = pgb.parameters()
                    pgb.parameters(Name = alias + '_' + pgbParams['Group']) #type: ignore

            '''
            pll_seq_def = project.definition('PLL_seq')
            if pll_seq_def:
                pll_seq_def.name = 'PLL_seq_9124'
            pll_adaptive = project.definition('PLL_ADAPTIVE')
            if pll_adaptive:
                pll_adaptive.name = 'PLL_ADAPTIVE_9124'
            FFT_TRACKING = project.definition('FFT_TRACKING')
            if FFT_TRACKING:
                FFT_TRACKING.name = 'FFT_TRACKING_9124'
            '''

def main():
    pscad = connectPSCAD() #type: ignore 
    updateUMs(pscad)
    print()

if __name__ == '__main__':
    main()

