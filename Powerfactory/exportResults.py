from os import remove
import pandas as pd
from types import SimpleNamespace

def exportResults(app, name, path, refName, refScale):
    fixHeaderAndReference = False
    project = app.GetActiveProject()
    networkData = app.GetProjectFolder('netdat')
    if not project:
        raise Exception('No project activated')

    export = networkData.SearchObject('PP-MTB\\exportResults\\graphExport.ComWr')     
    board = app.GetFromStudyCase('SetDesktop')
    plots = board.GetContents('*.GrpPage',1)   
    comRes = app.GetFromStudyCase('ComRes')
    elmRes = app.GetFromStudyCase('ElmRes')

    for p in plots:
        p.Show()
        p.DoAutoScaleX()
        p.DoAutoScaleY()
        filePath =  '{}\\{}_{}.png'.format(path, name, p.GetAttribute('e:loc_name'))
        export.SetAttribute('e:f', filePath)
        export.Execute()

    csvFileName = '{}\\{}.csv'.format(path, name)
    comRes.SetAttribute('pResult', elmRes)
    comRes.SetAttribute('iopt_exp', 6)
    comRes.SetAttribute('iopt_sep', 0)
    comRes.SetAttribute('ciopt_head', 1)
    comRes.SetAttribute('dec_Sep', ',')
    comRes.SetAttribute('col_Sep', ';')
    comRes.SetAttribute('f_name', csvFileName)
    comRes.Execute()
    
    if fixHeaderAndReference:
        df = pd.read_csv(csvFileName, sep = ';', decimal = ',', header = 1)
        rnmDict ={
            'b:tnow in s'         : 't[s]',
            'm:u:bus2:A in p.u.'  : 'uA[pu]',
            'm:u:bus2:B in p.u.'  : 'uB[pu]',
            'm:u:bus2:C in p.u.'  : 'uC[pu]', 
            'm:u1:bus2 in p.u.'   : 'u1[pu]',
            'm:u2:bus2 in p.u.'   : 'u2[pu]',
            'm:phiu1:bus2 in deg' : 'phi[deg]',
            'm:i:bus2:A in p.u.'  : 'iA[pu]',
            'm:i:bus2:B in p.u.'  : 'iB[pu]',
            'm:i:bus2:C in p.u.'  : 'iC[pu]',
            'm:i1:bus2 in p.u.'   : 'i1[pu]',
            'm:i2:bus2 in p.u.'   : 'i2[pu]',
            'm:i1P:bus2 in p.u.'  : 'i1d[pu]',
            'm:i1Q:bus2 in p.u.'  : 'i1q[pu]',
            'm:i2P:bus2 in p.u.'  : 'i2d[pu]',
            'm:i2Q:bus2 in p.u.'  : 'i2q[pu]',
            'm:cosphisum:bus2'    : 'cosphi[-]',
            'm:fehz in Hz'        : 'f[hz]',
            's:p in p.u.'         : 'p[pu]',
            's:q in p.u.'         : 'q[pu]',
            's:p2 in p.u.'        : 'p2[pu]',
            's:q2 in p.u.'        : 'q2[pu]'
            }
        
        leftOutSignals = list(set(df.columns) - set(rnmDict.keys()))
        
        if len(leftOutSignals) == 1:
            df = df.rename(columns={leftOutSignals[0] : refName})
            df[refName] = df[refName]/refScale
        elif len(leftOutSignals) > 1:
            raise Exception('Unkown signals in resultfile')

        df = df.rename(columns=rnmDict)
        remove(csvFileName)
        df.to_csv(csvFileName, sep=';', decimal=',', index=False)

if __name__ == "__main__":
    import powerfactory as PF # type: ignore

    app = PF.GetApplication()
    if not app:
        raise Exception('No connection to powerfactory application')

    thisScript = app.GetCurrentScript()
    name : str = str(thisScript.GetInputParameterString('name')[1])
    path : str = str(thisScript.GetInputParameterString('path')[1])
    refName : str = str(thisScript.GetInputParameterString('refName')[1])
    refScale : float = float(thisScript.GetInputParameterDouble('refScale')[1])

    exportResults(app, name, path, refName, refScale)