import pandas as pd
import math
import sys
if getattr(sys, 'gettrace', None) is not None:
  sys.path.append('C:\\Program Files\\DIgSILENT\\PowerFactory 2022 SP3\\Python\\3.8')
import powerfactory as PF # type: ignore
import time
from datetime import datetime
from types import SimpleNamespace
from setupPlots import setupPlots
from exportResults import exportResults
from setupCase import setupCase

forceNoSetup : bool = False
forceNoRun : bool = False
forceNoPlot : bool = False
forceNoExport : bool = False

def connectPF():
  if getattr(sys, 'gettrace', None) is not None:
    app : PF.Application = PF.GetApplicationExt()
    if not app:
      raise Exception('No connection to powerfactory application')
    app.Show()
    app.ClearOutputWindow()

    project : PF.DataObject = app.GetActiveProject()
    while project is None:
        time.sleep(1)
        project = app.GetActiveProject() 
    networkData = app.GetProjectFolder('netdat')
    thisScript : PF.DataObject = networkData.SearchObject('PP-MTB\\execute.ComPython')
  else:
    app : PF.Application = PF.GetApplication()
    if not app:
      raise Exception('No connection to powerfactory application')
    app.ClearOutputWindow()
    project : PF.DataObject = app.GetActiveProject()
    if not project:
      raise Exception('No project activated')
    thisScript = app.GetCurrentScript()

  return app, project, thisScript

def readScriptOptions(thisScript) -> SimpleNamespace:
  # Read the script options
  options = SimpleNamespace()
  options.file : str =  thisScript.GetInputParameterString('file')[1]
  options.preBackup : bool = bool(thisScript.GetInputParameterInt('preBackup')[1])
  options.postBackup : bool = bool(thisScript.GetInputParameterInt('postBackup')[1])
  options.setup  : bool = bool(thisScript.GetInputParameterInt('setup')[1])
  options.run : bool = bool(thisScript.GetInputParameterInt('run')[1])
  options.plot : bool = bool(thisScript.GetInputParameterInt('plot')[1])
  options.export : bool = bool(thisScript.GetInputParameterInt('export')[1])
  options.eventPlot : bool = bool(thisScript.GetInputParameterInt('eventPlot')[1])
  options.faultPlot : bool = bool(thisScript.GetInputParameterInt('faultPlot')[1])
  options.phasePlot : bool = bool(thisScript.GetInputParameterInt('phasePlot')[1])
  options.nameByRank : bool = bool(thisScript.GetInputParameterInt('nameByRank')[1])
  options.varStep : bool = bool(thisScript.GetInputParameterInt('variableStep')[1])
  options.asymSim : bool = bool(thisScript.GetInputParameterInt('asymSim')[1])
  options.autoGen : bool = bool(thisScript.GetInputParameterInt('autoFindGen')[1])
  options.paraEventsOnly : bool = bool(thisScript.GetInputParameterInt('paraEventsOnly')[1])
  options.consolidate : bool = bool(thisScript.GetInputParameterInt('consolidate')[1]) 
  options.parallelComp : bool = bool(thisScript.GetInputParameterInt('parallelComp')[1])
  options.enforcedSync : bool = bool(thisScript.GetInputParameterInt('enforcedSync')[1])
  options.rfpath : str = thisScript.GetInputParameterString('rfpath')[1]
  options.mfpath : str = thisScript.GetInputParameterString('mfpath')[1]
  options.PspScale : float = thisScript.GetInputParameterDouble('PspScale')[1]
  options.QspScale : float = thisScript.GetInputParameterDouble('QspScale')[1]
  options.QUspScale : float = thisScript.GetInputParameterDouble('QUspScale')[1]
  options.QPFspScale : float = thisScript.GetInputParameterDouble('QPFspScale')[1]
  options.QPFmode : int = 0
  options.scaleP : float = thisScript.GetInputParameterDouble('PctrlMeasScale')[1]
  options.offsetP : float = thisScript.GetInputParameterDouble('PctrlMeasOffset')[1]
  options.scaleQ : float = thisScript.GetInputParameterDouble('QctrlMeasScale')[1]
  options.offsetQ : float = thisScript.GetInputParameterDouble('QctrlMeasOffset')[1]
  options.scaleU : float = thisScript.GetInputParameterDouble('VoltageMeasScale')[1]
  options.offsetU : float = thisScript.GetInputParameterDouble('VoltageMeasOffset')[1]
  options.scalePh : float = thisScript.GetInputParameterDouble('PhaseMeasScale')[1]
  options.offsetPh : float = thisScript.GetInputParameterDouble('PhaseMeasOffset')[1]
  options.scaleF : float = thisScript.GetInputParameterDouble('FreqMeasScale')[1]
  options.offsetF : float = thisScript.GetInputParameterDouble('FreqMeasOffset')[1]
  options.scale_offset : list = [[options.scaleP,options.offsetP],     # options.scale_offset[0]: P
                                 [options.scaleQ,options.offsetQ],     # options.scale_offset[1]: Q
                                 [options.scaleU,options.offsetU],     # options.scale_offset[2]: U
                                 [options.scalePh,options.offsetPh],   # options.scale_offset[3]: Ph
                                 [options.scaleF,options.offsetF]]     # options.scale_offset[4]: F

  # For the Pref and Qref tests
  options.PCtrl = thisScript.GetExternalObject('Pctrl')[1]
  if options.PCtrl == None or not (options.PCtrl.GetClassName() == 'ElmDsl' or options.PCtrl.GetClassName() == 'ElmComp'):
    exit('Pctrl inputblock is not a ElmDsl or ElmComp object.')
  options.QCtrl = thisScript.GetExternalObject('Qctrl')[1]
  if options.QCtrl == None or not (options.QCtrl.GetClassName() == 'ElmDsl' or options.QCtrl.GetClassName() == 'ElmComp'):
    exit('QCtrl inputblock is not a ElmDsl or ElmComp object.')
  options.QUCtrl = thisScript.GetExternalObject('QUctrl')[1]
  if options.QUCtrl == None or not (options.QUCtrl.GetClassName() == 'ElmDsl' or options.QUCtrl.GetClassName() == 'ElmComp'):
    exit('QUCtrl inputblock is not a ElmDsl or ElmComp object.')
  options.QPFCtrl = thisScript.GetExternalObject('QPFctrl')[1]
  if options.QPFCtrl == None or not (options.QPFCtrl.GetClassName() == 'ElmDsl' or options.QPFCtrl.GetClassName() == 'ElmComp'):
    exit('QPFCtrl inputblock is not a ElmDsl or ElmComp object.')
  options.PspInputName : str = thisScript.GetInputParameterString('PspSignal')[1]
  options.QspInputName : str = thisScript.GetInputParameterString('QspSignal')[1]
  options.QUspInputName : str = thisScript.GetInputParameterString('QUspSignal')[1]
  options.QPFspInputName : str = thisScript.GetInputParameterString('QPFspSignal')[1]
  options.QmodeVar = thisScript.GetExternalObject('QmodeVar')[1]
  options.QUmodeVar = thisScript.GetExternalObject('QUmodeVar')[1]
  options.QPFmodeVar = thisScript.GetExternalObject('QPFmodeVar')[1]
  options.FSMmodeVar = thisScript.GetExternalObject('FSMmodeVar')[1]

  return options

def parseGrid(app) -> SimpleNamespace:
  networkData = app.GetProjectFolder('netdat')
  grid = SimpleNamespace()
  grid.grid = networkData.SearchObject('PP-MTB.ElmNet')
  grid.impedance = networkData.SearchObject('PP-MTB\\Z.ElmSind')
  grid.voltageSource = networkData.SearchObject('PP-MTB\\vac.ElmVac')
  grid.measurement = networkData.SearchObject('PP-MTB\\meas.ElmSind')
  grid.PQmeasurement = networkData.SearchObject('PP-MTB\\pq.StaPqmea')
  grid.terminals = grid.grid.GetContents('*.ElmTerm')
  grid.poc = networkData.SearchObject('PP-MTB\\pcc.ElmTerm') 
  grid.cub = networkData.SearchObject('PP-MTB\\pcc\\cubZ.StaCubic')
  grid.cubmeas = networkData.SearchObject('PP-MTB\\pcc\\cubmeas.StaCubic')
  grid.pCtrl = networkData.SearchObject('PP-MTB\\pCtrl.ElmSecctrl')
  grid.qCtrl = networkData.SearchObject('PP-MTB\\qCtrl.ElmStactrl') 
  grid.sigGen = networkData.SearchObject('PP-MTB\\signalGenerator.ElmDsl')
  grid.connector = networkData.SearchObject('PP-MTB\\connector.ElmComp')
  grid.sinkP = networkData.SearchObject('PP-MTB\\lib\\connector\\sinkP.BlkSlot')
  grid.sinkQ = networkData.SearchObject('PP-MTB\\lib\\connector\\sinkQ.BlkSlot')
  grid.sinkVac = networkData.SearchObject('PP-MTB\\lib\\connector\\sinkVac.BlkSlot')
  grid.pctrlmeas = networkData.SearchObject('PP-MTB\\PctrlMeas.ElmFile')
  grid.qctrlmeas = networkData.SearchObject('PP-MTB\\QctrlMeas.ElmFile')
  grid.voltagemeas = networkData.SearchObject('PP-MTB\\VoltageMeas.ElmFile')
  grid.phasemeas = networkData.SearchObject('PP-MTB\\PhaseMeas.ElmFile')
  grid.freqmeas = networkData.SearchObject('PP-MTB\\FreqMeas.ElmFile')
  grid.measfiles : list = [grid.pctrlmeas,     # grid.measfiles[0]: P
                           grid.qctrlmeas,     # grid.measfiles[1]: Q
                           grid.voltagemeas,   # grid.measfiles[2]: U
                           grid.phasemeas,     # grid.measfiles[3]: Ph
                           grid.freqmeas]      # grid.measfiles[4]: F
  return grid

def loadPlantInfo(options) -> SimpleNamespace:
  pdInput = pd.read_excel(open(options.file,'rb'), sheet_name='Input', header = None, index_col = 0, usecols = 'A:B')
  plantInfo = SimpleNamespace()
  plantInfo.PROJECTNAME : str = pdInput[1]['ProjectName'] # Project name
  plantInfo.PN : float = pdInput[1]['Pn'] # Nominal active power (MW)
  plantInfo.UC : float = pdInput[1]['Uc'] # PoC nominal operating voltage (LL, RMS, kV)
  plantInfo.UN : float = pdInput[1]['Un'] # PoC nominal voltage (LL, RMS, kV)
  plantInfo.IN : float = plantInfo.PN/math.sqrt(3)/plantInfo.UN
  plantInfo.AREA : str = pdInput[1]['Area'] # DK1 or DK2
  plantInfo.SCRMIN : float = pdInput[1]['SCR min']
  plantInfo.SCRTUN : float = pdInput[1]['SCR tuning']
  plantInfo.SCRMAX : float = pdInput[1]['SCR max']
  plantInfo.XRRATIOMIN : float = pdInput[1]['X/R SCR min']
  plantInfo.XRRATIOTUN : float = pdInput[1]['X/R SCR tuning']
  plantInfo.XRRATIOMAX : float = pdInput[1]['X/R SCR max']
  plantInfo.R0 : float = pdInput[1]['R0'] # (p.u.)
  plantInfo.X0 : float = pdInput[1]['X0'] # (p.u.)
  plantInfo.DEFQMODE : str = pdInput[1]['Default Q mode'] # Q, Q(U) or PF
  plantInfo.QUDROOP : float = pdInput[1]['Q(u) droop'] # (%)
  plantInfo.OFFSET : float = pdInput[1]['Case time offset'] # (s)

  # Protection
  plantInfo.Uo1 : float = pdInput[1]['U>'] # p.u. voltage
  plantInfo.Uo1_t : float = pdInput[1]['U> time'] # s
  plantInfo.Uo2 : float = pdInput[1]['U>>'] # p.u. voltage
  plantInfo.Uo2_t : float = pdInput[1]['U>> time'] # s
  plantInfo.Uo3 : float = pdInput[1]['U>>>'] # p.u. voltage
  plantInfo.Uo3_t : float = pdInput[1]['U>>> time'] # s
  plantInfo.Uu : float = pdInput[1]['U<'] # p.u. voltage
  plantInfo.Uu_t : float = pdInput[1]['U< time'] # s
  plantInfo.Fo : float = pdInput[1]['f>'] # Hz
  plantInfo.Fo_t : float = pdInput[1]['f> time'] # s
  plantInfo.Fu : float = pdInput[1]['f<'] # Hz
  plantInfo.Fu_t : float = pdInput[1]['f< time'] # s
  plantInfo.dFpos : float = pdInput[1]['df/dt pos'] # Hz/s
  plantInfo.dFpos_t : float = pdInput[1]['df/dt pos time'] # s
  plantInfo.dFneg : float = pdInput[1]['df/dt neg'] # Hz/s
  plantInfo.dFneg_t : float = pdInput[1]['df/dt neg time'] # s

  return plantInfo

def parseSubscripts(app) -> SimpleNamespace:
  networkData = app.GetProjectFolder('netdat')
  subScripts = SimpleNamespace()
  subScripts.setupPlots = networkData.SearchObject('PP-MTB\\setupPlots.ComPython')
  subScripts.exportResults = networkData.SearchObject('PP-MTB\\exportResults.ComPython')
  return subScripts

def resetProjectUnits(app) -> None:
  # Restore the default settings for unit and time
  project = app.GetActiveProject()
  SetPrj = project.SearchObject('Settings\\Project Settings.SetPrj')
  if SetPrj:
    SetPrj.Delete()
  ComIncUnits = project.SearchObject('Settings\\Units')
  if ComIncUnits:
    ComIncUnits.Delete()
  project.Deactivate()
  project.Activate()

def readCase(app, plantInfo, pdCases, caseIndex) -> SimpleNamespace:
  # Load case
  case = SimpleNamespace()
  case.Rank : int = int(pdCases['Rank/Id'][caseIndex])
  case.Included = bool(pdCases['RMS'][caseIndex])
  case.Name : str = str(pdCases['Name'][caseIndex])

  try:
    case.U0 : float = float(pdCases['U0'][caseIndex])
  except:
    case.U0 : float = 1.0

  case.P0 : float = float(pdCases['P0'][caseIndex])
  case.FSMenabled = bool(pdCases['FSM enabled'][caseIndex])
  case.Qmode : str = str(pdCases['Qmode'][caseIndex])

  if case.Qmode == 'Default':
    case.Qmode = plantInfo.DEFQMODE

  if case.Qmode == 'Q':
    case.internalQmode = 0
  elif case.Qmode == 'Q(U)':
    case.internalQmode = 1
  elif case.Qmode == 'PF':
    case.internalQmode = 2
  else:
    app.PrintWarn('Invalid Qmode: {}. Assuming Q control mode.'.format(case.Qmode))
    case.internalQmode = 0

  case.Qinit : float = float(pdCases['Q ref. Initial'][caseIndex])
  case.SimTime : float = float(pdCases['Simulationtime'][caseIndex])
  strSCR = str(pdCases['SCR'][caseIndex])
  if strSCR == 'Ideal':
    case.GridImped = 0
    case.SCR = plantInfo.SCRMIN
    case.XRRATIO = plantInfo.XRRATIOMIN
  else:
    case.GridImped = 1
    if strSCR == 'Min':
      case.SCR = plantInfo.SCRMIN
      case.XRRATIO = plantInfo.XRRATIOMIN
    elif strSCR == 'Max':
      case.SCR = plantInfo.SCRMAX
      case.XRRATIO = plantInfo.XRRATIOMAX
    else: # strSCR == 'Tuning'
      case.SCR = plantInfo.SCRTUN
      case.XRRATIO = plantInfo.XRRATIOTUN

  case.PctrlMeas : str = str(pdCases['Pctrl meas. File'][caseIndex])
  case.QctrlMeas : str = str(pdCases['Qctrl meas. File'][caseIndex])
  case.VoltageMeas : str = str(pdCases['Voltage meas. File'][caseIndex])
  case.PhaseMeas : str = str(pdCases['Phase meas. File'][caseIndex])
  case.FreqMeas : str = str(pdCases['Frequency meas. File'][caseIndex])
  case.MeasFiles : list = [case.PctrlMeas,    # case.MeasFiles[0]: P
                           case.QctrlMeas,    # case.MeasFiles[1]: Q
                           case.VoltageMeas,  # case.MeasFiles[2]: U
                           case.PhaseMeas,    # case.MeasFiles[3]: Ph
                           case.FreqMeas]     # case.MeasFiles[4]: F

  case.events = list()

  eIndex = 0
  while True:
    typeLabel = 'type.{}'.format(eIndex)
    timeLabel = 'time.{}'.format(eIndex)
    spLabel = 'Sp or res. U.{}'.format(eIndex)
    rpLabel = 'ramp/periode.{}'.format(eIndex)
    if {typeLabel, timeLabel, spLabel, rpLabel}.issubset(pdCases.columns):
      evType = str(pdCases[typeLabel][caseIndex])
      evTime = float(pdCases[timeLabel][caseIndex])
      evSp = float(pdCases[spLabel][caseIndex])
      evRp = float(pdCases[rpLabel][caseIndex])
    else:
      break
    eIndex += 1
    if (math.isnan(evTime) and math.isnan(evSp) and math.isnan(evRp)) or evTime >= case.SimTime:
      continue
    case.events.append([evType,evTime,evSp,evRp])
  return case

def loadCases(app, options) -> pd.DataFrame:
  app.PrintPlain('Read Test Matrix: {}'.format(options.file)) 
  cases = pd.read_excel(open(options.file,'rb'), sheet_name='Cases', header = 1, index_col=None, usecols=lambda x: 'Unnamed' not in x)
  for i in range(len(cases.get('Rank/Id'))):
    if type(cases['Rank/Id'][i]) == str:
      cases.drop(i, inplace=True)
  cases.rename(columns = {'type':'type.0', 'time':'time.0', 'Sp or res. U':'Sp or res. U.0', 'ramp/periode':'ramp/periode.0'}, inplace = True)
  cases.reset_index(drop=True, inplace=True)
  return cases

def setup(app, thisScript, options, subScripts, grid, project):
  # Check if any studycase is active
  currentStudyCase = app.GetActiveStudyCase()
  if currentStudyCase is None and options.setup:
    exit('Please activate studycase.')

  # Get and check for active grids
  networkData = app.GetProjectFolder('netdat')
  grids = networkData.GetContents('.ElmNet', 1)
  activeGrids = list(filter(lambda x : x.IsCalcRelevant(), grids))

  if len(activeGrids) == 0 and options.setup:
    exit('No active grids.')

  studyTime : int = currentStudyCase.GetAttribute('iStudyTime')

  for var in (options.QmodeVar, options.QUmodeVar, options.QPFmodeVar, options.FSMmodeVar):
    if var is not None and var.GetAttribute('e:tToAc') >= studyTime:
      app.PrintWarn('The variation {} is active after or at the same time as the base case.'.format(var.GetFullName(0)))
    if not options.consolidate and var is not None:
      for stage in var.GetContents('*.IntSstage'):
        if stage.GetAttribute('e:tAcTime') == studyTime:
          exit('The variation stage {} is active at the same time as the the base case.'.format(stage.GetFullName(0)))

  recordingStage = app.GetRecordingStage()

  if not options.consolidate and recordingStage is not None and recordingStage.GetAttribute('e:tAcTime') == studyTime:
    exit('Expansionstage {} conflicts with PP-MTB setup.'.format(recordingStage.GetFullName(0)))

  # Create version pre-run
  if options.preBackup:
    project.CreateVersion('pre_PP-MTB_{}'.format(datetime.now().strftime(r'%d%m%Y%H%M%S')))

  resetProjectUnits(app)

  netFolder = app.GetProjectFolder('netmod')
  varFolder = app.GetProjectFolder('scheme')
  if varFolder is None:
    varFolder = netFolder.CreateObject('IntPrjfolder', 'Variations')
    varFolder.SetAttribute('iopt_typ', 'scheme')

  studyCaseSet = thisScript.GetChildren(0, 'studycaseSet.SetSelect')[0]
  taskAutoRef = thisScript.GetChildren(0, 'taskAutoRef.IntRef')[0]
  generatorSet = thisScript.GetChildren(0, 'generatorSet.SetSelect')[0]

  studyCaseSet.Clear()

  # Create studycase folder
  studyCaseFolder = app.GetProjectFolder('study')
  if not studyCaseFolder:
    studyCaseFolder = project.CreateObject('IntPrjfolder', 'Study Cases')
    studyCaseFolder.SetAttribute('iopt_typ', 'study')

  activeVars = app.GetActiveNetworkVariations()
  if options.consolidate:
    currentStudyCase.Consolidate()
    activeVars = []
  else:
    for var in activeVars:
      var.Deactivate()

  # Create task automation
  taskAuto = studyCaseFolder.CreateObject('ComTasks')
  taskAutoRef.SetAttribute('obj_id', taskAuto)
  taskAuto.SetAttribute('iEnableParal', options.parallelComp)
  taskAuto.SetAttribute('parMethod', 0)
  (taskAuto.GetAttribute('parallelSetting')).SetAttribute('procTimeOut', 3600)

  for var in activeVars:
    var.Activate()

  plantInfo = loadPlantInfo(options)
  cases = loadCases(app, options)
  maxRank = max(i for i in cases['Rank/Id'] if isinstance(i, int))
  
  app.EchoOff()

  for caseIndex in range(len(cases)):
    case = readCase(app, plantInfo, cases, caseIndex)
    if case.Included:
      setupCase(app, subScripts, options, plantInfo, grid, activeGrids, activeVars, case, generatorSet, studyCaseSet, studyCaseFolder, varFolder, taskAuto, maxRank, studyTime)
  
  app.EchoOn()

def plot(app, subScripts):
  eventPlot : bool = bool(subScripts.setupPlots.GetInputParameterInt('eventPlot')[1])
  faultPlot : bool = bool(subScripts.setupPlots.GetInputParameterInt('faultPlot')[1])
  phasePlot : bool = bool(subScripts.setupPlots.GetInputParameterInt('phasePlot')[1])
  uLim : float = float(subScripts.setupPlots.GetInputParameterDouble('uLim')[1])
  Qmode : int = int(subScripts.setupPlots.GetInputParameterInt('Qmode')[1])
  symSim : bool = bool(subScripts.setupPlots.GetInputParameterInt('symSim')[1])
  sigP : str = str(subScripts.setupPlots.GetInputParameterString('sigP')[1])
  scaleP : float = float(subScripts.setupPlots.GetInputParameterDouble('scaleP')[1])
  sigQ : str = str(subScripts.setupPlots.GetInputParameterString('sigQ')[1])
  scaleQ : float = float(subScripts.setupPlots.GetInputParameterDouble('scaleQ')[1])
  Ph : int = int(subScripts.setupPlots.GetInputParameterInt('Ph')[1])
  F : int = int(subScripts.setupPlots.GetInputParameterInt('F')[1])
  blockP = subScripts.setupPlots.GetExternalObject('blockP')[1]
  blockQ = subScripts.setupPlots.GetExternalObject('blockQ')[1]

  setupPlots(app, eventPlot, faultPlot, phasePlot, uLim, Qmode, symSim, sigP, scaleP, sigQ, scaleQ, Ph, F, blockP, blockQ)

def export(app, subScripts):
  name : str = str(subScripts.exportResults.GetInputParameterString('name')[1])
  path : str = str(subScripts.exportResults.GetInputParameterString('path')[1])
  refName : str = str(subScripts.exportResults.GetInputParameterString('refName')[1])
  refScale : float = float(subScripts.exportResults.GetInputParameterDouble('refScale')[1])
  exportResults(app, name, path, refName, refScale)

def main():
  app, project, thisScript = connectPF()
  options = readScriptOptions(thisScript)
  subScripts = parseSubscripts(app)
  grid = parseGrid(app)

  if options.setup and not forceNoSetup:
    setup(app, thisScript, options, subScripts, grid, project)

  studyCaseSet = thisScript.GetChildren(0, 'studycaseSet.SetSelect')[0]
  taskAutoRef = thisScript.GetChildren(0, 'taskAutoRef.IntRef')[0]

  app.EchoOn()
  if options.run and not forceNoRun:
    taskAuto = taskAutoRef.GetAttribute('obj_id')
    if taskAuto:
      taskAuto.Execute()
    else:
      exit('No setup to run.')

  app.EchoOff()
  for studycase in studyCaseSet.All():
    studycase.Activate()

    if options.plot and not forceNoPlot:
      plot(app, subScripts)

    if options.export and not forceNoExport:
      export(app, subScripts)

    studycase.Deactivate()
  app.EchoOn()

  if options.postBackup:
    project.CreateVersion('post_PP-MTB_{}'.format(datetime.now().strftime(r'%d%m%Y%H%M%S')))

if __name__ == "__main__":
  main()
