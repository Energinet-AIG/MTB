"""
Information concerning the use of Power Plant and Model Test Bench

Energinet provides the Power Plant and Model Test Bench (PP-MTB) for the purpose of developing a prequalification test bench for production facility and simulation performance which the facility owner may use in its own simulation environment in order to pre-test compliance with the applicable technical requirements for simulation models. The PP-MTB are provided under the following considerations:
-	Use of the PP-MTB and its results are indicative and for informational purposes only. Energinet may only in its own simulation environment perform conclusive testing, performance and compliance of the simulation models developed and supplied by the facility owner.
-	Downloading the PP-MTB and updating the PP-MTB must only be done through a Energinet provided link. Users of the PP-MTB must not share the PP-MTB with other facility owners. The facility owner should always use the latest version of the PP-MTB in order to get the most correct results. 
-	Use of the PP-MTB are at the facility owners and the users own risk. Energinet is not responsible for any damage to hardware or software, including simulation models or computers.
-	All intellectual property rights, including copyright to the PP-MTB remains at Energinet in accordance with applicable Danish law.
"""
import math

def setupCase(app, subScripts, options, plantInfo, grid, activeGrids, case, studyCaseSet, studyCaseFolder, baseVar, baseStage, varFolder, taskAuto):
    # Check case
    nrOfSelectedModes = sum([case.PrefCtrl, case.QrefCtrl, case.VSPhaseCtrl, case.VSVoltCtrl, case.VSFreqCtrl])
    if nrOfSelectedModes > 1:
        app.PrintWarn('Atmost one control-mode can be activated at one time. Case {} dropped'.format(case.Rank))
        return -1

    # Set-up studycase, variation and balance      
    caseName = '{}_{}_{}'.format(str(case.Rank).zfill(len(str(case.maxRank))), case.ION, case.TestType).replace('.', '')
    newStudycase = studyCaseFolder.CreateObject('IntCase', caseName)
    newStudycase.Activate()      
    newStudycase.SetStudyTime(case.studyTime)

    # Activate the relevant networks
    for g in activeGrids:
        g.Activate()

    if baseVar and baseStage:
        baseVar.Activate()
        baseStage.Activate()

    newVar = varFolder.CreateObject('IntScheme', caseName)
    newStage = newVar.CreateObject('IntSstage', caseName)
    newStage.SetAttribute('e:tAcTime', case.studyTime - 1)
    newVar.Activate()
    newStage.Activate()

    # Reactive power dispatch
    if case.internalQmode == 0:
        if not options.QmodeVar is None:
            options.QmodeVar.Activate() 
        case.Q0 = case.InitValue * plantInfo.PN # Mvar
    elif case.internalQmode == 1:
        if not options.QUmodeVar is None:
            options.QUmodeVar.Activate() 
        case.Q0 = 0 # Mvar
    else:
        if not options.QPFmodeVar is None:
            options.QPFmodeVar.Activate() 
        case.Q0 =  case.P0 * math.sqrt(1/case.InitValue**2 - 1) * plantInfo.PN # Mvar

    grid.qCtrl.SetAttribute('qsetp', case.Q0) 
    grid.qCtrl.SetAttribute('imode', 1) # Distribute Q demand according to rated power
    grid.qCtrl.SetAttribute('consQdisp', 0) 
    grid.qCtrl.SetAttribute('iQorient', 0)     # 0 = +Q
    grid.qCtrl.SetAttribute('p_cub', grid.cub)  
    grid.qCtrl.SetAttribute('i_ctrl', 1)
        
    # Active power dispatch 
    grid.pCtrl.SetAttribute('psetp', case.P0 * plantInfo.PN)    # P0 at PCC 
    grid.pCtrl.SetAttribute('rembar', grid.poc)
    grid.pCtrl.SetAttribute('imode', 0) # Distribute P demand according to rated power    

    #Setup grid  
    Rgrid = case.GridImped * plantInfo.VN * plantInfo.VN/(plantInfo.SCR*plantInfo.PN)/math.sqrt(plantInfo.XRRATIO*plantInfo.XRRATIO+1) # ohm
    Xgrid = case.GridImped * Rgrid * plantInfo.XRRATIO # ohm
    Lgrid = Xgrid/2/50/math.pi # H 

    grid.impedance.SetAttribute('rrea', Rgrid)
    grid.impedance.SetAttribute('lrea', Lgrid*1000)  # mH
    grid.voltageSource.SetAttribute('usetp', case.U0)     
    grid.sigGen.SetAttribute('e:outserv', True)
    grid.sigGen.SetAttribute('pelm:0', None)
    grid.sigGen.SetAttribute('signal:0', '')

    # Create event folder
    eventFolder = newStudycase.CreateObject('IntEvt')

    # Define fault-event
    symSim = True # Simulations should be symmetrical

    if case.FaultType != 0:
        faultType = 3 - math.ceil(case.FaultType/3) 
        faultStart = 3
        faultStop = faultStart + case.FaultPeriod

        case.FaultDepth = min(case.FaultDepth, case.U0 - 0.001)
        Rf = case.FaultDepth/(case.U0 - case.FaultDepth) * Rgrid
        Xf = case.FaultDepth/(case.U0 - case.FaultDepth) * Xgrid

        scEvent = eventFolder.CreateObject('EvtShc', 'fault') 
        scEvent.SetAttribute('p_target', grid.poc)
        scEvent.SetAttribute('time', faultStart)
        scEvent.SetAttribute('R_f', Rf)
        scEvent.SetAttribute('X_f', Xf)

        if faultType == 0:
            faultTypeName = 'ABC-G'
            scEvent.SetAttribute('i_shc', 0)
            scEvent.SetAttribute('loc_name', faultTypeName)
        elif faultType == 1:
            faultTypeName = 'AB-G'
            scEvent.SetAttribute('i_shc', 3)
            scEvent.SetAttribute('i_p2pgf', 0)
            scEvent.SetAttribute('loc_name', faultTypeName)
            symSim = False
        else:
            faultTypeName = 'A-G'
            scEvent.SetAttribute('i_shc', 2)
            scEvent.SetAttribute('i_pspgf', 0) # A-G
            scEvent.SetAttribute('loc_name', faultTypeName)
            symSim = False

        cEvent = eventFolder.CreateObject('EvtShc', '{} clear'.format(faultTypeName))  
        cEvent.SetAttribute('p_target', grid.poc)
        cEvent.SetAttribute('time', faultStop)
        cEvent.SetAttribute('i_shc', 4)  

    # Define step/ramp event
    ctrlMode = -1
    inputScaling = 1
    inputSignal = ''
    inputBlock = None
    refName = 'nan' 

    if nrOfSelectedModes > 0:
        ctrlMode = case.QrefCtrl + 2 * case.VSPhaseCtrl + 3 * case.VSVoltCtrl + 4 * case.VSFreqCtrl
        sigGenEvent = True

        if ctrlMode == 0:
            inputBlock = options.PCtrl
            inputSignal = options.PspInputName
            inputScaling = options.PspScale
            refName = 'Pref[pu]'
        elif ctrlMode == 1:
            if case.internalQmode == 0:
                inputBlock = options.QCtrl
                inputSignal = options.QspInputName
                inputScaling = options.QspScale
                refName = 'Qref[pu]'
            elif case.internalQmode == 1:
                inputBlock = options.QUCtrl
                inputSignal = options.QUspInputName
                inputScaling = options.QUspScale
                refName = 'QUref[pu]'
            else:
                inputBlock = options.QPFCtrl
                inputSignal = options.QPFspInputName
                inputScaling = options.QPFspScale
                refName = 'QPref[-]'
        else:
            inputBlock = grid.voltageSource
  
        if ctrlMode == 2:
            inputSignal = 'dphiu'
            inputScaling = math.pi/180
            refName = 'Phiref[deg]'         
        elif ctrlMode == 3:
            inputSignal = 's:u0'
            inputScaling = 1
            minimumVoltage = math.inf
            refName = 'Vref[pu]'
        elif ctrlMode == 4:
            inputSignal = 'F0Hz'
            inputScaling = 1
            refName = 'Fref[Hz]'
        else:
            blockSignals = (inputBlock.GetAttribute('t:sInput')[0]).split(',')
            blockParams = (inputBlock.GetAttribute('e:parameterNames')[0]).split(',')
            if inputSignal.split(':')[-1] in blockParams:
                sigGenEvent = False
                inputSignal = 'c:{}'.format(inputSignal.split(':')[-1])
            elif inputSignal.split(':')[-1] in blockSignals:
                inputSignal = 's:{}'.format(inputSignal.split(':')[-1])
            else:
                #Unknown signal
                pass
        
        #Conect block
        if sigGenEvent:
            grid.connector.SetAttribute('pelm:1', inputBlock) 
            grid.sink.SetAttribute('sInput', [inputSignal.split(':')[-1]])
            grid.sigGen.SetAttribute('e:outserv', False)
            grid.sigGen.SetAttribute('scale', inputScaling)

        for event in case.events:
            evStart = event[0]
            evSp = event[1]
            evRp = event[2]
            if sigGenEvent:
                # State event
                evSpEvent = eventFolder.CreateObject('EvtParam','signalGenState')
                evSpEvent.SetAttribute('p_target', grid.sigGen)
                evSpEvent.SetAttribute('time', evStart)
                evSpEvent.SetAttribute('variable', 's:x')
                evSpEvent.SetAttribute('value', str(evSp))         
                # Slope event
                evRpEvent = eventFolder.CreateObject('EvtParam','signalGenSlope')
                evRpEvent.SetAttribute('p_target', grid.sigGen)
                evRpEvent.SetAttribute('time', evStart)
                evRpEvent.SetAttribute('variable', 'c:slope')
                evRpEvent.SetAttribute('value', str(evRp))
                if ctrlMode == 3: 
                    minimumVoltage = min(minimumVoltage, evSp)
            else:
                evSpEvent = eventFolder.CreateObject('EvtParam','signalGenState')
                evSpEvent.SetAttribute('p_target', inputBlock)
                evSpEvent.SetAttribute('time', evStart)
                evSpEvent.SetAttribute('variable', inputSignal)
                evSpEvent.SetAttribute('value', str(inputScaling * evSp))  


    symSim = symSim and not options.asymSim

    # setup simulation and loadflow
    inc = app.GetFromStudyCase('ComInc')
    ldf = app.GetFromStudyCase('ComLdf')
    sim = app.GetFromStudyCase('ComSim')

    ldf.SetAttribute('iopt_lim', 1)
    ldf.SetAttribute('iopt_apdist', 1)
    ldf.SetAttribute('iPST_at', 1)
    ldf.SetAttribute('iopt_at', 1)
    ldf.SetAttribute('iopt_asht', 1)
    ldf.SetAttribute('iopt_plim', 1)
    ldf.SetAttribute('iopt_lim', 1)
    ldf.SetAttribute('iopt_net', not symSim)
    inc.SetAttribute('iopt_net', 'sym' if symSim else 'rst')
    inc.SetAttribute('iopt_show', 1)
    inc.SetAttribute('dtgrd', 0.001)
    inc.SetAttribute('dtgrd_max', 0.01)
    inc.SetAttribute('tstart', 0)
    inc.SetAttribute('iopt_sync', 1)
    inc.SetAttribute('syncperiod', 0.001)
    inc.SetAttribute('iopt_adapt', options.varStep)
    inc.SetAttribute('iopt_lt', 0) #A-stable per element

    sim.SetAttribute('tstop', case.Tstop) 

    # Add resultvariables 
    res = app.GetFromStudyCase('ElmRes')

    # POC voltages
    res.AddVariable(grid.measurement, 'm:u:bus2:A')
    res.AddVariable(grid.measurement, 'm:u:bus2:B')
    res.AddVariable(grid.measurement, 'm:u:bus2:C')
    res.AddVariable(grid.measurement, 'm:u1:bus2')
    res.AddVariable(grid.measurement, 'm:u2:bus2')
    res.AddVariable(grid.measurement, 'm:phiu1:bus2')

    # POC currents
    res.AddVariable(grid.measurement, 'm:i:bus2:A')
    res.AddVariable(grid.measurement, 'm:i:bus2:B')
    res.AddVariable(grid.measurement, 'm:i:bus2:C')
    res.AddVariable(grid.measurement, 'm:i1:bus2')
    res.AddVariable(grid.measurement, 'm:i2:bus2')
    res.AddVariable(grid.measurement, 'm:i1P:bus2')
    res.AddVariable(grid.measurement, 'm:i1Q:bus2')
    res.AddVariable(grid.measurement, 'm:i2P:bus2')
    res.AddVariable(grid.measurement, 'm:i2Q:bus2')

    # Powers
    res.AddVariable(grid.PQmeasurement, 's:p')
    res.AddVariable(grid.PQmeasurement, 's:q')        
    res.AddVariable(grid.PQmeasurement, 's:p2')
    res.AddVariable(grid.PQmeasurement, 's:q2')        

    # Other
    res.AddVariable(grid.poc, 'm:fehz')
    res.AddVariable(grid.measurement, 'm:cosphisum:bus2')
    if not inputBlock is None:
        res.AddVariable(inputBlock, inputSignal)

    # Setup plot setup script
    subScripts.setupPlots.SetInputParameterInt('eventPlot', ctrlMode != -1 or options.eventPlot)
    subScripts.setupPlots.SetInputParameterInt('faultPlot', case.FaultType != 0 or options.faultPlot)
    subScripts.setupPlots.SetInputParameterInt('phasePlot', case.FaultType != 0 or options.phasePlot)
    subScripts.setupPlots.SetInputParameterInt('ctrlMode', ctrlMode)
    subScripts.setupPlots.SetInputParameterInt('Qmode', case.internalQmode)
    subScripts.setupPlots.SetInputParameterInt('symSim', symSim)
    subScripts.setupPlots.SetInputParameterDouble('PN', 1)
    subScripts.setupPlots.SetInputParameterString('inputSignal', inputSignal)
    subScripts.setupPlots.SetInputParameterDouble('inputScaling', inputScaling)
    subScripts.setupPlots.SetExternalObject('inputBlock', inputBlock)

    # Setup export plots
    if options.nameByRank:
        exportName = '{}_{}'.format(plantInfo.PROJECTNAME, case.Rank)
    else:
        exportName = caseName

    subScripts.exportResults.SetInputParameterString('name', exportName)
    subScripts.exportResults.SetInputParameterString('path', options.fpath)
    subScripts.exportResults.SetInputParameterString('refName', refName)
    subScripts.exportResults.SetInputParameterDouble('refScale', inputScaling)

    # Add to taskautomation
    taskAuto.AppendStudyCase(newStudycase)
    taskAuto.AppendCommand(inc, -1)
    taskAuto.AppendCommand(sim, -1)
    taskAuto.AppendCommand(subScripts.setupPlots, -1)
    newStudycase.Deactivate()
    studyCaseSet.AddRef(newStudycase)