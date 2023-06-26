from types import SimpleNamespace

def setupPlots(app,eventPlot,faultPlot,phasePlot,uLim,Qmode,symSim,sigP,scaleP,sigQ,scaleQ,Ph,F,blockP,blockQ):
    project = app.GetActiveProject()
    networkData = app.GetProjectFolder('netdat')
    if not project:
        raise Exception('No project activated')
    
    events = (True if sigP != '' else False, 
              True if sigQ != '' else False,
              False, # Slot for U
              True if Ph == 1 else False,
              True if F == 1 else False)

    grid = SimpleNamespace()
    grid.measurement = networkData.SearchObject('PP-MTB\\meas.ElmSind')
    grid.PQmeasurement = networkData.SearchObject('PP-MTB\\pq.StaPqmea')
    grid.poc = networkData.SearchObject('PP-MTB\\pcc.ElmTerm') 

    res = app.GetFromStudyCase('ElmRes')
    res.Load()

    board = app.GetFromStudyCase('SetDesktop')
    plots = board.GetContents('*.GrpPage',1)
    
    for p in plots:
        p.RemovePage()

    # Event plot
    if eventPlot:
        evPage = board.GetPage('PQuf', 1,'GrpPage')

        #Event - F plot
        if events[4]:
            Fplot = evPage.GetOrInsertCurvePlot('F')
            (Fplot.GetDataSeries()).AddCurve(grid.poc, 'm:fehz')

        #Event - Angle plot
        if events[3]:
            Phplot = evPage.GetOrInsertCurvePlot('Angle')
            (Phplot.GetDataSeries()).AddCurve(grid.measurement, 'm:phiu1:bus2')

        #Event - P plot
        Pplot = evPage.GetOrInsertCurvePlot('PQ' if Qmode == 2 else 'P')
        PplotDS = Pplot.GetDataSeries()  
        if events[0] and blockP is not None:
            PplotDS.AddCurve(blockP, sigP)
            PplotDS.SetAttribute('e:enableDataTrafo', 1)
            PplotDS.SetAttribute('e:curveTableNormalise:0', 1)
            PplotDS.SetAttribute('e:curveTableNormValue:0', scaleP)
        PplotDS.AddCurve(grid.PQmeasurement, 's:p')
        if Qmode == 2:
            PplotDS.AddCurve(grid.PQmeasurement, 's:q')
        if not symSim:
            PplotDS.AddCurve(grid.PQmeasurement, 's:p2')
            if Qmode == 2:
                PplotDS.AddCurve(grid.PQmeasurement, 's:q2')
       
        #Event - Q plot
        if Qmode == 2:
            Qplot = evPage.GetOrInsertCurvePlot('cos(phi)')
            QplotDS = Qplot.GetDataSeries()
            QplotDS.AddCurve(grid.measurement,'m:cosphisum:bus2')
        else:
            Qplot = evPage.GetOrInsertCurvePlot('Q')
            QplotDS = Qplot.GetDataSeries()
            if events[1] and Qmode == 0 and blockQ is not None:
                QplotDS.AddCurve(blockQ, sigQ)
                QplotDS.SetAttribute('e:enableDataTrafo', 1)
                QplotDS.SetAttribute('e:curveTableNormalise:0', 1)
                QplotDS.SetAttribute('e:curveTableNormValue:0', scaleQ)
            QplotDS.AddCurve(grid.PQmeasurement, 's:q')
            if not symSim:
                QplotDS.AddCurve(grid.PQmeasurement, 's:q2')

        #Event - U plot
        Uplot = evPage.GetOrInsertCurvePlot('U')
        UplotDS = Uplot.GetDataSeries()
        if events[1] and Qmode == 1 and blockQ is not None:
            UplotDS.AddCurve(blockQ, sigQ)
            UplotDS.SetAttribute('e:curveTableNormalise:0', 1)
            UplotDS.SetAttribute('e:curveTableNormValue:0', scaleQ)
        UplotDS.AddCurve(grid.measurement, 'm:u1:bus2')
        if not symSim:
            UplotDS.AddCurve(grid.measurement, 'm:u2:bus2')

    volCol = res.FindColumn(grid.measurement, 'm:u1:bus2')
    minVoltage = res.FindMinInColumn(volCol)[1]
   
    # Fault plot
    if faultPlot or minVoltage <= uLim:
        fPage = board.GetPage('Idq', 1, 'GrpPage')

        #Fault - U plot
        Uplot = fPage.GetOrInsertCurvePlot('U')
        UplotDS = Uplot.GetDataSeries()
        UplotDS.AddCurve(grid.measurement, 'm:u1:bus2')
        if not symSim:
            UplotDS.AddCurve(grid.measurement, 'm:u2:bus2')

        #Fault - iQ plot
        iQplot = fPage.GetOrInsertCurvePlot('Iq')
        iQplotDS = iQplot.GetDataSeries()
        iQplotDS.AddCurve(grid.measurement, 'm:i1Q:bus2')
        if not symSim:
            iQplotDS.AddCurve(grid.measurement, 'm:i2Q:bus2')

        #Fault - iP plot
        iPplot = fPage.GetOrInsertCurvePlot('Id')
        iPplotDS = iPplot.GetDataSeries()
        iPplotDS.AddCurve(grid.measurement, 'm:i1P:bus2')
        if not symSim:
            iPplotDS.AddCurve(grid.measurement, 'm:i2P:bus2')

    # Phase voltage and current plot
    if phasePlot or minVoltage <= uLim:
        pPage = board.GetPage('UI', 1, 'GrpPage')

        #Phase - U plot
        Uplot = pPage.GetOrInsertCurvePlot('U')
        UplotDS = Uplot.GetDataSeries()
        UplotDS.AddCurve(grid.measurement, 'm:u:bus2:A')
        UplotDS.AddCurve(grid.measurement, 'm:u:bus2:B')
        UplotDS.AddCurve(grid.measurement, 'm:u:bus2:C')

        #Phase - I plot
        Iplot = pPage.GetOrInsertCurvePlot('I')
        IplotDS = Iplot.GetDataSeries()
        IplotDS.AddCurve(grid.measurement, 'm:i:bus2:A')
        IplotDS.AddCurve(grid.measurement, 'm:i:bus2:B')
        IplotDS.AddCurve(grid.measurement, 'm:i:bus2:C')
    
    # Plot scaling and settings
    for page in board.GetContents('*.GrpPage',1):
        for plot in page.GetContents('*.PltLinebarplot',1):
            (plot.GetTitleObject()).SetAttribute('e:showTitle', 0)
            (plot.GetAxisX()).SetAttribute('e:scaleOnDataChange', 1)
            (plot.GetAxisY()).SetAttribute('e:scaleOnDataChange', 1)
            (plot.GetAxisY()).SetAttribute('e:limitMinimumToOrigin', 0)
        page.SetAttribute('e:autoLayoutMode', 1)
        page.DoAutoScale()


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
    Qmode : int = int(thisScript.GetInputParameterInt('Qmode')[1])
    symSim : bool = bool(thisScript.GetInputParameterInt('symSim')[1])
    sigP : str = str(thisScript.GetInputParameterString('sigP')[1])
    scaleP : float = float(thisScript.GetInputParameterDouble('scaleP')[1])
    sigQ : str = str(thisScript.GetInputParameterString('sigQ')[1])
    scaleQ : float = float(thisScript.GetInputParameterDouble('scaleQ')[1])
    Ph : int = int(thisScript.GetInputParameterInt('Ph')[1])
    F : int = int(thisScript.GetInputParameterInt('F')[1])
    blockP : PF.DataObject = thisScript.GetExternalObject('blockP')[1]
    blockQ : PF.DataObject = thisScript.GetExternalObject('blockQ')[1]

    setupPlots(app, eventPlot, faultPlot, phasePlot, uLim, Qmode, symSim, sigP, scaleP, sigQ, scaleQ, Ph, F, blockP, blockQ)