import math
import powerfactory as PF # type: ignore
from types import SimpleNamespace

def setupMeasFiles(app : PF.DataObject, grid : SimpleNamespace, case : PF.DataObject, options : SimpleNamespace) -> tuple:
    usedMeasFiles = []

    if all(item is 'nan' for item in case.MeasFiles):
        return (False, False, False, False, False)
    elif options.mfpath is not None:
        for i, fileName in enumerate(case.MeasFiles):
            if fileName not in 'nan':
                measFile = grid.measfiles[i]
                measFile.SetAttribute('e:outserv', False)
                if options.mfpath[len(options.mfpath)-1] is not '\\':
                    filePath = options.mfpath + '\\' + fileName
                else:
                    filePath = options.mfpath + fileName
                measFile.SetAttribute('e:f_name', filePath)

                # Set scale and offset
                afac = measFile.GetAttribute('e:afac')
                bfac = measFile.GetAttribute('e:bfac')
                
                afac[0] = options.scale_offset[i][0] # Index 0: Scale
                bfac[0] = options.scale_offset[i][1] # Index 1: Offset

                measFile.SetAttribute('e:afac', afac)
                measFile.SetAttribute('e:bfac', bfac)

                # Set simulation time
                with open(filePath, 'r') as csv:
                    content = csv.readlines()
                    measEnd = float(content[-1].split(',')[0])
                if (str(case.SimTime) in 'nan') or (case.SimTime == 0) or (measEnd > float(case.SimTime)):
                    case.SimTime = measEnd

                usedMeasFiles.append(True)
            else:
                usedMeasFiles.append(False)
        for event in case.events:
            if event[1] > case.SimTime and event[1] is not 'nan':
                case.SimTime = event[1]
    else:
        if any(item is not 'nan' for item in case.MeasFiles):
            app.PrintError('A measurement files is specified for case ' + case.Rank + ', however a path for measurement files is not defined in the execute.py script. The measurement file is ignored.')
        usedMeasFiles = [False, False, False, False, False]

    # Connect block
    if usedMeasFiles[0]:
        grid.connector.SetAttribute('pelm:1', options.PCtrl)
        grid.sinkP.SetAttribute('e:sInput', ['sigP', options.PspInputName])

    if usedMeasFiles[1]:
        if case.internalQmode == 0:
            grid.connector.SetAttribute('pelm:2', options.QCtrl)
            grid.sinkQ.SetAttribute('e:sInput', ['sigQ', options.QspInputName])
        elif case.internalQmode == 1:
            grid.connector.SetAttribute('pelm:2', options.QUCtrl)
            grid.sinkQ.SetAttribute('e:sInput', ['sigQ', options.QUspInputName])
        elif case.internalQmode == 2:
            grid.connector.SetAttribute('pelm:2', options.QPFCtrl)
            grid.sinkQ.SetAttribute('e:sInput', ['sigQ', options.QPFspInputName])

    if any([usedMeasFiles[2], usedMeasFiles[3], usedMeasFiles[4]]):
        grid.connector.SetAttribute('pelm:3', grid.voltageSource)
        inputVac = grid.sinkVac.GetAttribute('e:sInput')[0].split(',')
        if usedMeasFiles[2]:
            inputVac[1] = 'u0'
        if usedMeasFiles[3]:
            inputVac[3] = 'dphiu'
        if usedMeasFiles[4]:
            inputVac[5] = 'F0Hz'
        grid.sinkVac.SetAttribute('e:sInput',['{},{},{},{},{},{}'.format(*inputVac)])

    return (usedMeasFiles[0], usedMeasFiles[1], usedMeasFiles[2], usedMeasFiles[3], usedMeasFiles[4])

def createFault(app : PF.DataObject, case : SimpleNamespace, grid : SimpleNamespace, plantInfo : SimpleNamespace, options : SimpleNamespace) -> bool:
    currentStudycase = app.GetActiveStudyCase()
    if not currentStudycase:
        exit('No studycase active. Cannot create faultevent.')

    # Check if eventfolder exists
    eventFolder = app.GetFromStudyCase('IntEvt')
    if not eventFolder:
        eventFolder = currentStudycase.CreateObject('IntEvt')

    if any('fault' in item[0] for item in case.events) and case.VoltageMeas not in 'nan':
        app.PrintWarn('Measurement file is used to control grid voltage. Any fault events manually created in excelsheet is ignored.')

    elif any('fault' in item[0] for item in case.events): # Create fault through Events
        prevEvent = 'nan'
        for event in case.events:
            if 'fault' in event[0] or 'fault' in prevEvent: # event[0]: type
                if 'fault' in event[0]:
                    faultType = str(event[0]) 
                    prevEvent = faultType
                elif event[0] in 'nan' and event[2] is not 'nan': # event[2]: Sp or res. U
                    faultType = prevEvent
                else:
                    continue

                faultStart = float(event[1]) + plantInfo.OFFSET # event[1]: time
                faultStop = faultStart + float(event[3]) # event[3]: ramp/periode

                faultDepth = min(float(event[2]), 0.999) 
                Rgrid = case.GridImped * plantInfo.UN * plantInfo.UN/(case.SCR*plantInfo.PN)/math.sqrt(case.XRRATIO*case.XRRATIO+1) # ohm
                Xgrid = case.GridImped * Rgrid * case.XRRATIO # ohm
                Rf = faultDepth/(1 - faultDepth) * Rgrid
                Xf = faultDepth/(1 - faultDepth) * Xgrid

                scEvent = eventFolder.CreateObject('EvtShc', 'fault') 
                scEvent.SetAttribute('e:p_target', grid.poc)
                scEvent.SetAttribute('e:time', faultStart)
                scEvent.SetAttribute('e:R_f', Rf)
                scEvent.SetAttribute('e:X_f', Xf)

                if faultType == '3p fault': 
                    faultTypeName = 'ABC-G'
                    scEvent.SetAttribute('e:i_shc', 0)
                    scEvent.SetAttribute('loc_name', faultTypeName)
                elif faultType == '2p-g fault':
                    faultTypeName = 'AB-G'
                    scEvent.SetAttribute('e:i_shc', 3)
                    scEvent.SetAttribute('e:i_p2pgf', 0)
                    scEvent.SetAttribute('loc_name', faultTypeName)
                elif faultType == '1p fault':
                    faultTypeName = 'A-G'
                    scEvent.SetAttribute('e:i_shc', 2)
                    scEvent.SetAttribute('e:i_pspgf', 0) # A-G
                    scEvent.SetAttribute('loc_name', faultTypeName)

                cEvent = eventFolder.CreateObject('EvtShc', '{} clear'.format(faultTypeName))  
                cEvent.SetAttribute('e:p_target', grid.poc)
                cEvent.SetAttribute('e:time', faultStop)
                cEvent.SetAttribute('e:i_shc', 4)
            else:
                prevEvent = 'nan'
                continue

        if any('3p fault' in item[0] for item in case.events):
            return True # Contains SymFault
        else:
            return False # Doesn't contain SymFault
        
    else:
        return True

def createStepRamp(app : PF.DataObject, case  : SimpleNamespace, grid : SimpleNamespace, plantInfo : SimpleNamespace, options : SimpleNamespace) -> tuple:
    currentStudycase = app.GetActiveStudyCase()
    if not currentStudycase:
        exit('No studycase active. Cannot create step/ramp event.')

    # Check if eventfolder exists
    eventFolder = app.GetFromStudyCase('IntEvt')
    if not eventFolder:
        eventFolder = currentStudycase.CreateObject('IntEvt')
    
    usedMeasFiles = setupMeasFiles(app, grid, case, options)

    P = Q = U = Ph = F = False
        
    if any('e' in item[0] for item in case.events): # Create Step/Ramp from Event (All event types contain letter 'e' except faults and 'nan')
        prevEvent = 'nan'
        inputP = grid.sinkP.GetAttribute('e:sInput')[0].split(',')
        inputQ = grid.sinkQ.GetAttribute('e:sInput')[0].split(',')
        inputVac = grid.sinkVac.GetAttribute('e:sInput')[0].split(',')

        for eventNo, event in enumerate(case.events):
            inputScaling = 1
            inputConvSp = None
            inputConvRp = None
            inputSig = ''
            inputSel = ''
            inputBlc = None
            refName = 'nan' #Currently not in use

            # Get event type
            if event[0] not in 'nan':
                evType = str(event[0]) # event[0]: type
                prevEvent = evType
            elif prevEvent not in 'nan':
                evType = prevEvent
            else:
                exit('Event 1 must be defined with a specific type. Cannot create step/ramp event.')
            
            evStart = float(event[1]) + plantInfo.OFFSET # event[1]: time

            # Handle step setpoint
            if evType == 'dVoltage':
                evSp = case.U0 + float(event[2]) # event[2]: Sp or res. U
            else:
                if str(event[2]) not in 'nan': 
                    evSp = float(event[2])
                else:
                    evSp = str(event[2])
            
            # Handle ramp setpoint
            if str(event[3]) not in 'nan': # event[3]: ramp/periode
                evRp = event[3] 
            elif str(evSp) not in 'nan':
                evRp = 0.0
            sigGenEvent = True

            if evType == 'Pctrl ref.': 
                P = True
                if usedMeasFiles[0]: # PCtrl ref. measurement used
                    app.PrintWarn('Active power ref. measurement used. Pctrl ref. event ignored.')
                    continue
                inputBlc = options.PCtrl
                inputSig = inputP[0] = options.PspInputName
                inputSel = 'P'
                inputScaling = options.PspScale
                inputConvSp = None
                inputConvRp = None
                refName = 'Pref[pu]'
            elif evType == 'Qctrl ref.':
                Q = True
                if usedMeasFiles[1]: # QCtrl ref. measurement used
                    app.PrintWarn('Reactive power ref. measurement used. Qctrl ref. event ignored.')
                    continue
                inputSel = 'Q'
                if case.internalQmode == 0:
                    inputBlc = options.QCtrl
                    inputSig = inputQ[0] = options.QspInputName
                    inputScaling = options.QspScale
                    refName = 'Qref[pu]'
                elif case.internalQmode == 1:
                    inputBlc = options.QUCtrl
                    inputSig = inputQ[0] = options.QUspInputName
                    inputScaling = options.QUspScale
                    refName = 'QUref[pu]'
                else:
                    inputBlc = options.QPFCtrl
                    inputSig = inputQ[0] = options.QPFspInputName
                    inputScaling = options.QPFspScale
                    if options.QPFmode == 1:
                        inputConvSp = lambda x : math.acos(abs(x)) if x >= 0 else -math.acos(abs(x))
                    refName = 'QPref[-]'
            else:
                inputBlc = grid.voltageSource

            inputBlcIsDsl = inputBlc.GetClassName() == 'ElmDsl'
            if evType == 'Phase':
                Ph = True
                if usedMeasFiles[3]: # Phase measurement used
                    app.PrintWarn('Phase measurement used. Phase event ignored.')
                    continue
                inputSigFixed = inSigVar =  inputVac[2] ='dphiu'
                inputScaling = math.pi/180
                inputSel = 'Ph'
                refName = 'Phiref[deg]'
            elif str(evType) == 'Voltage':
                U = True
                if usedMeasFiles[2]: # Voltage measurement used
                    app.PrintWarn('Voltage measurement used. Voltage event ignored.')
                    continue
                inputSigFixed = inSigVar = inputVac[0] = 'u0'
                inputSel = 'U'
                refName = 'Vref[pu]'
            elif evType == 'dVoltage':
                U = True
                if usedMeasFiles[2]: # Voltage measurement used
                    app.PrintWarn('Voltage measurement used. dVoltage event ignored.')
                    continue
                inputSigFixed = inSigVar = inputVac[0] = 'u0'
                inputSel = 'U'
                refName = 'dVref[dpu]'
            elif evType == 'Frequency':
                F = True
                if usedMeasFiles[4]: # Frequency measurement used 
                    app.PrintWarn('Frequecny measurement used. Frequency event ignored.')
                    continue
                inputSigFixed = inSigVar = inputVac[4] = 'F0Hz'
                inputSel = 'F'
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

            # Signal conversion
            if str(evSp) not in 'nan':
                evSpConv = evSp if inputConvSp is None else inputConvSp(evSp)
            if str(evRp) not in 'nan':
                evRpConv = evRp if inputConvRp is None else inputConvRp(evRp)

            if sigGenEvent:
                if evType == 'Pctrl ref.':
                    grid.connector.SetAttribute('pelm:1', inputBlc)
                elif evType == 'Qctrl ref.':
                    grid.connector.SetAttribute('pelm:2', inputBlc)
                else:
                    grid.connector.SetAttribute('pelm:3', inputBlc)
                grid.sigGen.SetAttribute('e:scale{}'.format(inputSel), inputScaling)

                # State event
                if str(evSp) not in 'nan':
                    evSpEvent = eventFolder.CreateObject('EvtParam','signalGenState_' + str(eventNo+1))
                    evSpEvent.SetAttribute('e:p_target', grid.sigGen)
                    evSpEvent.SetAttribute('e:time', evStart)
                    evSpEvent.SetAttribute('e:variable', 's:x{}'.format(inputSel))
                    evSpEvent.SetAttribute('e:value', str(evSpConv))         
                if 'ref.' not in evType and str(evRp) not in 'nan': # Slope event
                    evRpEvent = eventFolder.CreateObject('EvtParam','signalGenSlope_' + str(eventNo+1))
                    evRpEvent.SetAttribute('e:p_target', grid.sigGen)
                    evRpEvent.SetAttribute('e:time', evStart)
                    evRpEvent.SetAttribute('e:variable', 'c:slope{}'.format(inputSel))
                    evRpEvent.SetAttribute('e:value', str(evRpConv))
            elif str(evSp) not in 'nan':
                evSpEvent = eventFolder.CreateObject('EvtParam','signalGenState_' + str(eventNo+1))
                evSpEvent.SetAttribute('e:p_target', inputBlc)
                evSpEvent.SetAttribute('e:time', evStart)
                evSpEvent.SetAttribute('e:variable', inputSigFixed)
                evSpEvent.SetAttribute('e:value', str(inputScaling * evSpConv))
            elif not inputBlcIsDsl:
                app.PrintWarn('Parameter event cannot be applied to composite-frame.')

        # Conect block
        grid.sinkP.SetAttribute('e:sInput',['{},{}'.format(*inputP)])
        grid.sinkQ.SetAttribute('e:sInput',['{},{}'.format(*inputQ)])
        grid.sinkVac.SetAttribute('e:sInput',['{},{},{},{},{},{}'.format(*inputVac)])
        grid.sigGen.SetAttribute('e:outserv', False) 
    
    return(P, Q, U, Ph, F)

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
        gen.SetAttribute('e:pgini', 0)
        gen.SetAttribute('e:qgini', 0)
        gen.SetAttribute('e:c_pstac', grid.qCtrl)          
        gen.SetAttribute('e:c_psecc', grid.pCtrl)  

    # Reactive power dispatch
    if case.internalQmode == 1:
        grid.qCtrl.SetAttribute('e:i_ctrl', 0)
        grid.qCtrl.SetAttribute('e:uset_mode', 1)
        grid.qCtrl.SetAttribute('e:i_droop', 1)
        grid.qCtrl.SetAttribute('e:Srated', plantInfo.PN * 0.33)
        grid.qCtrl.SetAttribute('e:ddroop', plantInfo.QUDROOP)
        grid.qCtrl.SetAttribute('e:pQmeas', grid.cubmeas)
    else:
        if case.internalQmode == 0:
            Q0 = case.Qinit * plantInfo.PN # Mvar
        else:
            Q0 =  case.P0 * math.sqrt(1/case.Qinit**2 - 1) * plantInfo.PN # Mvar   
        grid.qCtrl.SetAttribute('e:qsetp', Q0) 
        grid.qCtrl.SetAttribute('e:imode', 1) # Distribute Q demand according to rated power
        grid.qCtrl.SetAttribute('e:consQdisp', 0) 
        grid.qCtrl.SetAttribute('e:iQorient', 0)     # 0 = +Q
        grid.qCtrl.SetAttribute('e:p_cub', grid.cub)  
        grid.qCtrl.SetAttribute('e:i_ctrl', 1)
        
    # Active power dispatch 
    grid.pCtrl.SetAttribute('e:psetp', case.P0 * plantInfo.PN)    # P0 at PCC 
    grid.pCtrl.SetAttribute('e:rembar', grid.poc)
    grid.pCtrl.SetAttribute('e:imode', 0) # Distribute P demand according to rated power    

def setupResFile(app : PF.DataObject, grid : SimpleNamespace, options : SimpleNamespace, case : SimpleNamespace, events : list) -> None:
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

    if events[0]: # P ref.
        res.AddVariable(options.PCtrl, options.PspInputName)
    if events[1]: # Q ref.
        if case.internalQmode == 0: # Reactive power setpoint
            res.AddVariable(options.QCtrl, options.QspInputName)
        elif case.internalQmode == 1: # Voltage droop
            res.AddVariable(options.QUCtrl, options.QUspInputName)
        else: # Power factor
            res.AddVariable(options.QPFCtrl, options.QPFspInputName)
    if events[2]: # U 
        res.AddVariable(grid.voltageSource, 's:u0')
    if events[3]: # Ph 
        res.AddVariable(grid.voltageSource, 's:dphiu')
    if events[4]: # F
        res.AddVariable(grid.voltageSource, 's:F0Hz')

def setupStaticCalc(app : PF.DataObject, options : SimpleNamespace, symSim : bool) -> None:
    inc = app.GetFromStudyCase('ComInc')
    ldf = app.GetFromStudyCase('ComLdf')
    ldf.SetAttribute('e:iopt_lim', 1)
    ldf.SetAttribute('e:iopt_apdist', 1)
    ldf.SetAttribute('e:iPST_at', 1)
    ldf.SetAttribute('e:iopt_at', 1)
    ldf.SetAttribute('e:iopt_asht', 1)
    ldf.SetAttribute('e:iopt_plim', 1)
    ldf.SetAttribute('e:iopt_net', not symSim)
    inc.SetAttribute('e:iopt_net', 'sym' if symSim else 'rst')
    inc.SetAttribute('e:iopt_show', 1)
    inc.SetAttribute('e:dtgrd', 0.001)
    inc.SetAttribute('e:dtgrd_max', 0.01)
    inc.SetAttribute('e:tstart', 0)
    inc.SetAttribute('e:iopt_sync', options.enforcedSync)
    inc.SetAttribute('e:syncperiod', 0.001)
    inc.SetAttribute('e:iopt_adapt', options.varStep)
    inc.SetAttribute('e:iopt_lt', 0) #A-stable per element

def setupGrid(case : SimpleNamespace, grid : SimpleNamespace, plantInfo : SimpleNamespace) -> None:
    #Setup grid  
    Rgrid = case.GridImped * plantInfo.UN * plantInfo.UN/(case.SCR*plantInfo.PN)/math.sqrt(case.XRRATIO*case.XRRATIO+1) # ohm
    Xgrid = case.GridImped * Rgrid * case.XRRATIO # ohm
    Lgrid = Xgrid/2/50/math.pi # H 

    grid.impedance.SetAttribute('e:ucn',plantInfo.UN)
    grid.impedance.SetAttribute('e:Sn', plantInfo.PN)
    grid.impedance.SetAttribute('e:rrea', Rgrid)
    grid.impedance.SetAttribute('e:lrea', Lgrid*1000)  # mH
    for term in grid.terminals:
        term.SetAttribute('e:uknom', plantInfo.UN)
    grid.voltageSource.SetAttribute('e:usetp', case.U0)
    grid.voltageSource.SetAttribute('e:Unom', plantInfo.UN)
    grid.voltageSource.SetAttribute('e:phisetp', 0) 
    grid.voltageSource.SetAttribute('e:contbar', grid.poc)  
    grid.measurement.SetAttribute('e:ucn', plantInfo.UN)
    grid.measurement.SetAttribute('e:Sn', plantInfo.PN)   

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
     taskAuto : PF.DataObject,
     maxRank : int,
     studyTime : int) -> None:
    
    #Pre checks
    eventTypes = [item[0] for item in case.events]
    if options.PspInputName == '' and 'Pctrl ref.' in eventTypes:
        app.PrintWarn('Study case: {} will be ignored. PspSignal is empty and PrefCtrl == 1.'.format(case.Name))
        return 

    if options.QspInputName == '' and 'Qctrl ref.' in eventTypes and case.internalQmode == 0:
        app.PrintWarn('Study case: {} will be ignored. QspSignal is empty and Qmode is active and controlled.'.format(case.Name))
        return     

    if options.QUspInputName == '' and 'Qctrl ref.' in eventTypes and case.internalQmode == 1:
        app.PrintWarn('Study case: {} will be ignored. QUspSignal is empty and QUmode is active and controlled.'.format(case.Name))
        return 

    if options.QPFspInputName == '' and 'Qctrl ref.' in eventTypes and case.internalQmode == 2:
        app.PrintWarn('Study case: {} will be ignored. QPFspSignal is empty and QPFmode is active and controlled.'.format(case.Name))
        return 

    # Set-up studycase, variation and balance      
    caseName = '{}_{}'.format(str(case.Rank).zfill(len(str(maxRank))), case.Name).replace('.', '')
    newStudycase = studyCaseFolder.CreateObject('IntCase', caseName)
    newStudycase.Activate()      
    newStudycase.SetStudyTime(studyTime)

    # Activate the relevant networks
    for g in activeGrids:
        g.Activate()

    # Activate the relevant variations 
    for v in activeVars:
        v.Activate()

    newVar = varFolder.CreateObject('IntScheme', caseName)
    newStage = newVar.CreateObject('IntSstage', caseName)
    newStage.SetAttribute('e:tAcTime', studyTime)
    newVar.Activate()
    newStage.Activate()

    setupGrid(case, grid, plantInfo)
    symFault = createFault(app, case, grid, plantInfo, options)
    symSim = symFault and not options.asymSim
    events = createStepRamp(app, case, grid, plantInfo, options)
    refName = ''
    setupStaticCalc(app, options, symSim)
    staticDispatch(case, options, grid, activeGrids, plantInfo, generatorSet)
    setDynQmode(case, options)
    setupResFile(app, grid, options, case, events)

    # setup simulation and loadflow
    inc = app.GetFromStudyCase('ComInc')
    sim = app.GetFromStudyCase('ComSim')
    sim.SetAttribute('e:tstop', case.SimTime) 
    
    # Setup plot setup script
    subScripts.setupPlots.SetInputParameterInt('eventPlot', options.eventPlot)
    subScripts.setupPlots.SetInputParameterInt('faultPlot', not any('1p fault' in item[0] for item in case.events) or options.faultPlot)
    subScripts.setupPlots.SetInputParameterInt('phasePlot', not any('1p fault' in item[0] for item in case.events) or options.phasePlot)
    subScripts.setupPlots.SetInputParameterInt('Qmode', case.internalQmode)
    subScripts.setupPlots.SetInputParameterInt('symSim', symSim)
    subScripts.setupPlots.SetInputParameterString('sigP', str(options.PspInputName) if events[0] else '')
    subScripts.setupPlots.SetInputParameterDouble('scaleP', options.PspScale if events[0] else 1.0)
    subScripts.setupPlots.SetInputParameterString('sigQ', options.QspInputName if events[1] and case.internalQmode == 0 else 
                                                       options.QUspInputName if events[1] and case.internalQmode == 1 else 
                                                       options.QPFspInputName if events[1] and case.internalQmode == 2 else '')
    subScripts.setupPlots.SetInputParameterDouble('scaleQ', options.Qspscale if events[1] and case.internalQmode == 0 else 
                                                         options.QUspScale if events[1] and case.internalQmode == 1 else 
                                                         options.QPFspScale if events[1] and case.internalQmode == 2 else 1.0)
    subScripts.setupPlots.SetInputParameterInt('Ph', 1 if events[3] else 0)
    subScripts.setupPlots.SetInputParameterInt('F', 1 if events[4] else 0)
    subScripts.setupPlots.SetExternalObject('blockP', options.PCtrl)
    subScripts.setupPlots.SetExternalObject('blockQ', options.QCtrl if case.internalQmode == 0 else options.QUCtrl if case.internalQmode == 1 else options.QPFCtrl)

    # Setup export plots
    if options.nameByRank:
        exportName = '{}_{}'.format(plantInfo.PROJECTNAME, case.Rank)
    else:
        exportName = caseName

    subScripts.exportResults.SetInputParameterString('name', exportName)
    subScripts.exportResults.SetInputParameterString('path', options.rfpath)
    subScripts.exportResults.SetInputParameterString('refName', refName)
    subScripts.exportResults.SetInputParameterDouble('refScale', 1) # Scaling is disabled, previously: ...('refScale', 1 if inputScaling is None else inputScaling)

    # Add to taskautomation
    taskAuto.AppendStudyCase(newStudycase)
    taskAuto.AppendCommand(inc, -1)
    taskAuto.AppendCommand(sim, -1)
    taskAuto.AppendCommand(subScripts.setupPlots, -1)
    newStudycase.Deactivate()
    studyCaseSet.AddRef(newStudycase)
