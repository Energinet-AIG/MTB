"""
Information concerning the use of Power Plant and Model Test Bench

Energinet provides the Power Plant and Model Test Bench (PP-MTB) for the purpose of developing a prequalification test bench for production facility and simulation performance which the facility owner may use in its own simulation environment in order to pre-test compliance with the applicable technical requirements for simulation models. The PP-MTB are provided under the following considerations:
-	Use of the PP-MTB and its results are indicative and for informational purposes only. Energinet may only in its own simulation environment perform conclusive testing, performance and compliance of the simulation models developed and supplied by the facility owner.
-	Downloading the PP-MTB and updating the PP-MTB must only be done through a Energinet provided link. Users of the PP-MTB must not share the PP-MTB with other facility owners. The facility owner should always use the latest version of the PP-MTB in order to get the most correct results. 
-	Use of the PP-MTB are at the facility owners and the users own risk. Energinet is not responsible for any damage to hardware or software, including simulation models or computers.
-	All intellectual property rights, including copyright to the PP-MTB remains at Energinet in accordance with applicable Danish law.
"""
from types import SimpleNamespace

def setupPlots(app,eventPlot,faultPlot,phasePlot,uLim,ctrlMode,Qmode,symSim,inputBlock,inputSignal,inputScaling):
    project = app.GetActiveProject()
    networkData = app.GetProjectFolder('netdat')
    if not project:
        raise Exception('No project activated')

    grid = SimpleNamespace()
    grid.measurement = networkData.SearchObject('PP-MTB\\meas.ElmSind')
    grid.PQmeasurement = networkData.SearchObject('PP-MTB\\pq.StaPqmea')
    grid.poc = networkData.SearchObject('PP-MTB\\pcc.ElmTerm') 

    res = app.GetFromStudyCase('ElmRes')
    res.Load()

    board = app.GetFromStudyCase('SetDesktop')
    plots = board.GetContents('*.SetVipage',1)
    
    board.SetAttribute('e:auto_xscl', 2)

    for p in plots:
        p.Close()
        p.Delete()

    if eventPlot:
        evPage = board.GetPage('PQuf', 1, 'SetVipage')
        evPage.SetStyle('Paper')
        evPage.SetAttribute('iopt_add', 'hor')

        #Event - F/angle plot
        if ctrlMode == 2 or ctrlMode == 4:
            Faplot = evPage.GetOrInsertPlot('F' if ctrlMode == 4 else 'Angle',  'VisPlot', 1)
            Faplot.AddResVars(res, grid.poc if ctrlMode == 4 else grid.measurement, 'm:fehz' if ctrlMode == 4 else 'm:phiu1:bus2')
            Faplot.SetAttribute('e:auto_yscl', 2)

        #Event - P plot
        Pplot = evPage.GetOrInsertPlot('PQ' if Qmode == 2 else 'P',  'VisPlot', 1)      
        if ctrlMode == 0 and inputBlock is not None:
            Pplot.AddResVars(res, inputBlock, inputSignal)
            #Pplot.SetAttribute('e:userdesc:0', 'Reference')
            Pplot.SetAttribute('e:dIsNom:0', 1)
            Pplot.SetAttribute('e:dValNom:0', inputScaling)
        Pplot.AddResVars(res, grid.PQmeasurement, 's:p')
        if Qmode == 2:
            Pplot.AddResVars(res, grid.PQmeasurement, 's:q')    
        if not symSim:
            Pplot.AddResVars(res, grid.PQmeasurement, 's:p2')
            if Qmode == 2:
                Pplot.AddResVars(res, grid.PQmeasurement, 's:q2')
        Pplot.SetAttribute('e:auto_yscl', 2)
            
        #Event - Q plot
        if Qmode == 2:
            Qplot = evPage.GetOrInsertPlot('cos(phi)',  'VisPlot', 1)
            Qplot.AddResVars(res, grid.measurement,'m:cosphisum:bus2')
        else:
            Qplot = evPage.GetOrInsertPlot('Q',  'VisPlot', 1)
            if ctrlMode == 1 and Qmode == 0 and inputBlock is not None:
                Qplot.AddResVars(res, inputBlock, inputSignal)
                #Qplot.SetAttribute('e:userdesc:0', 'Reference')
                Qplot.SetAttribute('e:dIsNom:0', 1)
                Qplot.SetAttribute('e:dValNom:0', inputScaling)
            Qplot.AddResVars(res, grid.PQmeasurement, 's:q')
            if not symSim:
                Qplot.AddResVars(res, grid.PQmeasurement, 's:q2')
        Qplot.SetAttribute('e:auto_yscl', 2)

        #Event - U plot
        Uplot = evPage.GetOrInsertPlot('U',  'VisPlot', 1)
        if ctrlMode == 1 and Qmode == 1 and inputBlock is not None:
            Uplot.AddResVars(res, inputBlock, inputSignal)
            #Uplot.SetAttribute('e:userdesc:0', 'Reference')
            Uplot.SetAttribute('e:dIsNom:0', 1)
            Uplot.SetAttribute('e:dValNom:0', inputScaling)
        Uplot.AddResVars(res, grid.measurement, 'm:u1:bus2')
        if not symSim:
            Uplot.AddResVars(res, grid.measurement, 'm:u2:bus2')
        Uplot.SetAttribute('e:auto_yscl', 2)

        evPage.DoAutoScaleX()
        evPage.DoAutoScaleY()

    volCol = res.FindColumn(grid.measurement, 'm:u1:bus2')
    minVoltage = res.FindMinInColumn(volCol)[1]
   
    # Fault plot
    if faultPlot or minVoltage <= uLim:
        fPage = board.GetPage('Idq', 1, 'SetVipage')
        fPage.SetStyle('Paper')
        fPage.SetAttribute('iopt_add', 'hor')
        
        Uplot = fPage.GetOrInsertPlot('U',  'VisPlot', 1)
        Uplot.AddResVars(res, grid.measurement, 'm:u1:bus2')
        if not symSim:
            Uplot.AddResVars(res, grid.measurement, 'm:u2:bus2')
        Uplot.SetAttribute('e:auto_yscl', 2)

        iQplot = fPage.GetOrInsertPlot('Iq',  'VisPlot', 1)
        iQplot.AddResVars(res, grid.measurement, 'm:i1Q:bus2')
        if not symSim:
            iQplot.AddResVars(res, grid.measurement, 'm:i2Q:bus2')
        iQplot.SetAttribute('e:auto_yscl', 2)

        iPplot = fPage.GetOrInsertPlot('Id',  'VisPlot', 1)
        iPplot.AddResVars(res, grid.measurement, 'm:i1P:bus2')
        if not symSim:
            iPplot.AddResVars(res, grid.measurement, 'm:i2P:bus2')
        iPplot.SetAttribute('e:auto_yscl', 2)

        fPage.DoAutoScaleX()
        fPage.DoAutoScaleY()

    # Phase voltage and current plot
    if phasePlot or minVoltage <= uLim:
        pPage = board.GetPage('UI', 1, 'SetVipage')
        pPage.SetStyle('Paper')
        pPage.SetAttribute('iopt_add', 'hor')

        Uplot = pPage.GetOrInsertPlot('U',  'VisPlot', 1)
        Uplot.AddResVars(res, grid.measurement, 'm:u:bus2:A')
        Uplot.AddResVars(res, grid.measurement, 'm:u:bus2:B')
        Uplot.AddResVars(res, grid.measurement, 'm:u:bus2:C')
        Uplot.SetAttribute('e:auto_yscl', 2)

        Iplot = pPage.GetOrInsertPlot('I',  'VisPlot', 1)
        Iplot.AddResVars(res, grid.measurement, 'm:i:bus2:A')
        Iplot.AddResVars(res, grid.measurement, 'm:i:bus2:B')
        Iplot.AddResVars(res, grid.measurement, 'm:i:bus2:C')        
        Iplot.SetAttribute('e:auto_yscl', 2)

        pPage.DoAutoScaleX()
        pPage.DoAutoScaleY()

if __name__ == "__main__":
    import powerfactory as PF # type: ignore

    app = PF.GetApplication()
    if not app:
        raise Exception('No connection to powerfactory application')

    thisScript = app.GetCurrentScript()
    eventPlot : bool = bool(thisScript.GetInputParameterInt('eventPlot')[1])
    faultPlot : bool = bool(thisScript.GetInputParameterInt('faultPlot')[1])
    phasePlot : bool = bool(thisScript.GetInputParameterInt('phasePlot')[1])
    uLim : float = float(thisScript.GetInputParameterDouble('uLim')[1])
    ctrlMode : int = int(thisScript.GetInputParameterInt('ctrlMode')[1])
    Qmode : int = int(thisScript.GetInputParameterInt('Qmode')[1])
    symSim : bool = bool(thisScript.GetInputParameterInt('symSim')[1])
    inputSignal : str = str(thisScript.GetInputParameterString('inputSignal')[1])
    inputScaling : float = float(thisScript.GetInputParameterDouble('inputScaling')[1])
    inputBlock : PF.DataObject = thisScript.GetExternalObject('inputBlock')[1]

    setupPlots(app, eventPlot, faultPlot, phasePlot, uLim, ctrlMode, Qmode, symSim, inputBlock, inputSignal, inputScaling)