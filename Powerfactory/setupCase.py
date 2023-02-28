import math
import powerfactory as PF # type: ignore
from types import SimpleNamespace

def createFault(app : PF.DataObject, case : SimpleNamespace, grid : SimpleNamespace, plantInfo : SimpleNamespace) -> bool:
    if case.FaultType != 0:
        currentStudycase = app.GetActiveStudyCase()
        if not currentStudycase:
            exit('No studycase active. Cannot create faultevent.')

         # Check if eventfolder exists
        eventFolder = app.GetFromStudyCase('IntEvt')
        if not eventFolder:
            eventFolder = currentStudycase.CreateObject('IntEvt')

        faultType = 3 - math.ceil(case.FaultType/3) # Map from EMT-lile fault numbering to PF numbering
        faultStart = 3 # All faults happen at t = 3s
        faultStop = faultStart + case.FaultPeriod

        case.FaultDepth = min(case.FaultDepth, case.U0 - 0.001)
        Rgrid = case.GridImped * plantInfo.VN * plantInfo.VN/(plantInfo.SCR*plantInfo.PN)/math.sqrt(plantInfo.XRRATIO*plantInfo.XRRATIO+1) # ohm
        Xgrid = case.GridImped * Rgrid * plantInfo.XRRATIO # ohm
        Rf = case.FaultDepth/(case.U0 - case.FaultDepth) * Rgrid
        Xf = case.FaultDepth/(case.U0 - case.FaultDepth) * Xgrid

        scEvent = eventFolder.CreateObject('EvtShc', 'fault') 
        scEvent.SetAttribute('p_target', grid.poc)
        scEvent.SetAttribute('time', faultStart)
        scEvent.SetAttribute('R_f', Rf)
        scEvent.SetAttribute('X_f', Xf)

        symFault = False

        if faultType == 0:
            faultTypeName = 'ABC-G'
            scEvent.SetAttribute('i_shc', 0)
            scEvent.SetAttribute('loc_name', faultTypeName)
            symFault = True
        elif faultType == 1:
            faultTypeName = 'AB-G'
            scEvent.SetAttribute('i_shc', 3)
            scEvent.SetAttribute('i_p2pgf', 0)
            scEvent.SetAttribute('loc_name', faultTypeName)
        else:
            faultTypeName = 'A-G'
            scEvent.SetAttribute('i_shc', 2)
            scEvent.SetAttribute('i_pspgf', 0) # A-G
            scEvent.SetAttribute('loc_name', faultTypeName)

        cEvent = eventFolder.CreateObject('EvtShc', '{} clear'.format(faultTypeName))  
        cEvent.SetAttribute('p_target', grid.poc)
        cEvent.SetAttribute('time', faultStop)
        cEvent.SetAttribute('i_shc', 4)
        return symFault
    else:
        return True

def createStepRamp(app : PF.DataObject, case  : SimpleNamespace, grid : SimpleNamespace, options : SimpleNamespace) -> tuple:
    ctrlN = case.PrefCtrl + case.QrefCtrl + case.VSPhaseCtrl + case.VSVoltCtrl + case.VSFreqCtrl
    if ctrlN == 1:
        # Define step/ramp event
        currentStudycase = app.GetActiveStudyCase()
        if not currentStudycase:
            exit('No studycase active. Cannot create step/ramp event.')

        # Check if eventfolder exists
        eventFolder = app.GetFromStudyCase('IntEvt')
        if not eventFolder:
            eventFolder = currentStudycase.CreateObject('IntEvt')

        inputScaling = 1
        inputConvSp = None
        inputConvRp = None
        inputSig = ''
        inputBlc = None
        refName = 'nan' #Currently not in use
      
        ctrlMode = case.QrefCtrl + 2 * case.VSPhaseCtrl + 3 * case.VSVoltCtrl + 4 * case.VSFreqCtrl
        sigGenEvent = True

        if ctrlMode == 0:
            inputBlc = options.PCtrl
            inputSig = options.PspInputName
            inputScaling = options.PspScale
            inputConvSp = None
            inputConvRp = None
            refName = 'Pref[pu]'
        elif ctrlMode == 1:
            if case.internalQmode == 0:
                inputBlc = options.QCtrl
                inputSig = options.QspInputName
                inputScaling = options.QspScale
                refName = 'Qref[pu]'
            elif case.internalQmode == 1:
                inputBlc = options.QUCtrl
                inputSig = options.QUspInputName
                inputScaling = options.QUspScale
                refName = 'QUref[pu]'
            else:
                inputBlc = options.QPFCtrl
                inputSig = options.QPFspInputName
                inputScaling = options.QPFspScale
                if options.QPFmode == 1:
                    inputConvSp = lambda x : math.acos(abs(x)) if x >= 0 else -math.acos(abs(x))
                refName = 'QPref[-]'
        else:
            inputBlc = grid.voltageSource

        inputBlcIsDsl = inputBlc.GetClassName() == 'ElmDsl'
        if ctrlMode == 2:
            inputSigFixed = inSigVar = 'dphiu'
            inputScaling = math.pi/180
            refName = 'Phiref[deg]'         
        elif ctrlMode == 3:
            inputSigFixed = inSigVar = 'u0'
            refName = 'Vref[pu]'
        elif ctrlMode == 4:
            inputSigFixed = inSigVar = 'F0Hz'
            refName = 'Fref[Hz]'
        else:
            inSigVar = inputSig.split(':')[-1]
            
            blcSigsCmd = inputBlc.GetAttribute('t:sInput')
            if inputBlcIsDsl:
                blcParamsCmd = inputBlc.GetAttribute('e:parameterNames')
            else: # if composite frame
                blcParamsCmd = []

            if len(blcSigsCmd) > 0:
                blcSigs = blcSigsCmd[0].split(',')
            else:
                blcSigs = []

            if len(blcParamsCmd) > 0:
                blcParams = blcParamsCmd[0].split(',')
            else:
                blcParams = []     

            if inSigVar in blcParams:
                sigGenEvent = False
                inputSigFixed = 'c:{}'.format(inSigVar)
            elif inSigVar in blcSigs:
                inputSigFixed = 's:{}'.format(inSigVar)
                if options.paraEventsOnly:
                    sigGenEvent = False
            else:
                #Check if inputBlc is dll based
                if inputBlcIsDsl:
                    inputBlcTyp = inputBlc.GetAttribute('typ_id')
                    inputBlcModTyp = inputBlcTyp.GetAttribute('iopt_modtyp')
                    if inputBlcModTyp == 1 and blcSigs == []:
                        app.PrintWarn('t:sInput is empty and signalname is not in parameterlist. DSL model is compiled. Assuming {} is signal.'.format(inSigVar))
                        inputSigFixed = 's:{}'.format(inSigVar)
                        if options.paraEventsOnly:
                            sigGenEvent = False                    
                    else:
                        exit('Signal {} not found on inputblock: {}'.format(inSigVar, inputBlc.GetFullName(0))) 
                else:
                    exit('Signal {} not found on inputblock: {}'.format(inSigVar, inputBlc.GetFullName(0)))
   
        #Conect block
        if sigGenEvent:
            grid.connector.SetAttribute('pelm:1', inputBlc) 
            grid.sink.SetAttribute('sInput', [inSigVar])
            grid.sigGen.SetAttribute('e:outserv', False)
            grid.sigGen.SetAttribute('scale', inputScaling)
        elif not inputBlcIsDsl:
            app.PrintWarn('Parameter event cannot be applied to composite-frame.')

        for event in case.events:
            evStart = event[0]
            evSp = event[1]
            evSpConv = evSp if inputConvSp is None else inputConvSp(evSp)
            evRp = event[2]
            evRpConv = evRp if inputConvRp is None else inputConvRp(evRp)
            if sigGenEvent:
                # State event
                evSpEvent = eventFolder.CreateObject('EvtParam','signalGenState')
                evSpEvent.SetAttribute('p_target', grid.sigGen)
                evSpEvent.SetAttribute('time', evStart)
                evSpEvent.SetAttribute('variable', 's:x')
                evSpEvent.SetAttribute('value', str(evSpConv))         
                # Slope event
                evRpEvent = eventFolder.CreateObject('EvtParam','signalGenSlope')
                evRpEvent.SetAttribute('p_target', grid.sigGen)
                evRpEvent.SetAttribute('time', evStart)
                evRpEvent.SetAttribute('variable', 'c:slope')
                evRpEvent.SetAttribute('value', str(evRpConv))
            else:
                evSpEvent = eventFolder.CreateObject('EvtParam','signalGenState')
                evSpEvent.SetAttribute('p_target', inputBlc)
                evSpEvent.SetAttribute('time', evStart)
                evSpEvent.SetAttribute('variable', inputSigFixed)
                evSpEvent.SetAttribute('value', str(inputScaling * evSpConv))  
        
        return (inputBlc, inputSigFixed, inputScaling, ctrlMode)
    elif ctrlN > 1:
        exit('Atmost one control-mode can be activated at one time')
    else:
        return (None, None, None, None)   
            
def setDynQmode(case : PF.DataObject, options : SimpleNamespace) -> None:
    # Reactive power dispatch
    if case.internalQmode == 0:
        if not options.QmodeVar is None:
            options.QmodeVar.Activate() 
    elif case.internalQmode == 1:
        if not options.QUmodeVar is None:
            options.QUmodeVar.Activate() 
    else:
        if not options.QPFmodeVar is None:
            options.QPFmodeVar.Activate() 

def staticDispatch(case : SimpleNamespace, options : SimpleNamespace, grid : SimpleNamespace, activeGrids : SimpleNamespace, plantInfo : SimpleNamespace, generatorSet : PF.DataObject) -> None:
    # Statgen dispatch
    if options.autoGen:
        generatorSet.Clear()
        for g in activeGrids:
            generatorSet.AddRef(g.GetContents('*.ElmGenstat', 1))
  

    for gen in generatorSet.All():
        gen.SetAttribute('pgini', 0)
        gen.SetAttribute('qgini', 0)
        gen.SetAttribute('c_pstac', grid.qCtrl)          
        gen.SetAttribute('c_psecc', grid.pCtrl)  

    # Reactive power dispatch
    if case.internalQmode == 0:
        Q0 = case.InitValue * plantInfo.PN # Mvar
    elif case.internalQmode == 1:
        Q0 = 0 # Mvar
    else:
        Q0 =  case.P0 * math.sqrt(1/case.InitValue**2 - 1) * plantInfo.PN # Mvar

    grid.qCtrl.SetAttribute('qsetp', Q0) 
    grid.qCtrl.SetAttribute('imode', 1) # Distribute Q demand according to rated power
    grid.qCtrl.SetAttribute('consQdisp', 0) 
    grid.qCtrl.SetAttribute('iQorient', 0)     # 0 = +Q
    grid.qCtrl.SetAttribute('p_cub', grid.cub)  
    grid.qCtrl.SetAttribute('i_ctrl', 1)
        
    # Active power dispatch 
    grid.pCtrl.SetAttribute('psetp', case.P0 * plantInfo.PN)    # P0 at PCC 
    grid.pCtrl.SetAttribute('rembar', grid.poc)
    grid.pCtrl.SetAttribute('imode', 0) # Distribute P demand according to rated power    

def setupResFile(app : PF.DataObject, grid : SimpleNamespace, inputBlock : PF.DataObject, inputSignal : str) -> None:
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

def setupStaticCalc(app : PF.DataObject, options : SimpleNamespace, symSim : bool) -> None:
    inc = app.GetFromStudyCase('ComInc')
    ldf = app.GetFromStudyCase('ComLdf')
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

def setupGrid(case : SimpleNamespace, grid : SimpleNamespace, plantInfo : SimpleNamespace) -> None:
    #Setup grid  
    Rgrid = case.GridImped * plantInfo.VN * plantInfo.VN/(plantInfo.SCR*plantInfo.PN)/math.sqrt(plantInfo.XRRATIO*plantInfo.XRRATIO+1) # ohm
    Xgrid = case.GridImped * Rgrid * plantInfo.XRRATIO # ohm
    Lgrid = Xgrid/2/50/math.pi # H 

    grid.impedance.SetAttribute('rrea', Rgrid)
    grid.impedance.SetAttribute('lrea', Lgrid*1000)  # mH
    grid.impedance.SetAttribute('ucn',plantInfo.VN)
    grid.impedance.SetAttribute('Sn', plantInfo.PN)
    for term in grid.terminals:
        term.SetAttribute('uknom', plantInfo.VN)
    grid.voltageSource.SetAttribute('usetp', case.U0)
    grid.voltageSource.SetAttribute('Unom', plantInfo.VN)
    grid.voltageSource.SetAttribute('phisetp', 0) 
    grid.voltageSource.SetAttribute('contbar', grid.poc)  
    grid.measurement.SetAttribute('ucn', plantInfo.VN)
    grid.measurement.SetAttribute('Sn', plantInfo.PN)   

    grid.sigGen.SetAttribute('e:outserv',True) 

def setupCase(app : PF.DataObject,
     subScripts : SimpleNamespace, 
     options : SimpleNamespace,
     plantInfo : SimpleNamespace,
     grid : SimpleNamespace,
     activeGrids : list,
     activeVars : list,
     case : SimpleNamespace,
     generatorSet : PF.DataObject,
     studyCaseSet : PF.DataObject,
     studyCaseFolder : PF.DataObject,
     varFolder : PF.DataObject,
     taskAuto : PF.DataObject) -> None:

    # Set-up studycase, variation and balance      
    caseName = '{}_{}_{}'.format(str(case.Rank).zfill(len(str(case.maxRank))), case.ION, case.TestType).replace('.', '')
    newStudycase = studyCaseFolder.CreateObject('IntCase', caseName)
    newStudycase.Activate()      
    newStudycase.SetStudyTime(case.studyTime)

    # Activate the relevant networks
    for g in activeGrids:
        g.Activate()

    if(not options.consolidate):
        # Activate the relevant variations 
        for v in activeVars:
            v.Activate()

    newVar = varFolder.CreateObject('IntScheme', caseName)
    newStage = newVar.CreateObject('IntSstage', caseName)
    newStage.SetAttribute('e:tAcTime', case.studyTime)
    newVar.Activate()
    newStage.Activate()

    setupGrid(case, grid, plantInfo)
    symFault = createFault(app, case, grid, plantInfo)
    symSim = symFault and not options.asymSim
    inputBlock, inputSignal, inputScaling, ctrlMode = createStepRamp(app, case, grid, options)
    refName = ''
    setupStaticCalc(app, options, symSim)
    staticDispatch(case, options, grid, activeGrids, plantInfo, generatorSet)
    setDynQmode(case, options)
    setupResFile(app, grid, inputBlock, inputSignal)

    # setup simulation and loadflow
    inc = app.GetFromStudyCase('ComInc')
    sim = app.GetFromStudyCase('ComSim')
    sim.SetAttribute('tstop', case.Tstop) 

    # Setup plot setup script
    subScripts.setupPlots.SetInputParameterInt('eventPlot', ctrlMode is not None or options.eventPlot)
    subScripts.setupPlots.SetInputParameterInt('faultPlot', case.FaultType != 0 or options.faultPlot)
    subScripts.setupPlots.SetInputParameterInt('phasePlot', case.FaultType != 0 or options.phasePlot)
    subScripts.setupPlots.SetInputParameterInt('ctrlMode', -1 if ctrlMode is None else ctrlMode)
    subScripts.setupPlots.SetInputParameterInt('Qmode', case.internalQmode)
    subScripts.setupPlots.SetInputParameterInt('symSim', symSim)
    subScripts.setupPlots.SetInputParameterDouble('PN', 1)
    subScripts.setupPlots.SetInputParameterString('inputSignal', '' if inputSignal is None else inputSignal)
    subScripts.setupPlots.SetInputParameterDouble('inputScaling', 1 if inputScaling is None else inputScaling)
    subScripts.setupPlots.SetExternalObject('inputBlock', inputBlock)

    # Setup export plots
    if options.nameByRank:
        exportName = '{}_{}'.format(plantInfo.PROJECTNAME, case.Rank)
    else:
        exportName = caseName

    subScripts.exportResults.SetInputParameterString('name', exportName)
    subScripts.exportResults.SetInputParameterString('path', options.fpath)
    subScripts.exportResults.SetInputParameterString('refName', refName)
    subScripts.exportResults.SetInputParameterDouble('refScale', 1 if inputScaling is None else inputScaling)

    # Add to taskautomation
    taskAuto.AppendStudyCase(newStudycase)
    taskAuto.AppendCommand(inc, -1)
    taskAuto.AppendCommand(sim, -1)
    taskAuto.AppendCommand(subScripts.setupPlots, -1)
    newStudycase.Deactivate()
    studyCaseSet.AddRef(newStudycase)