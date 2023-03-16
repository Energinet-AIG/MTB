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
  options.backup : bool = bool(thisScript.GetInputParameterInt('backup')[1]) 
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
  options.fpath : str = thisScript.GetInputParameterString('fpath')[1]
  options.PspScale : float = thisScript.GetInputParameterDouble('PspScale')[1]
  options.QspScale : float = thisScript.GetInputParameterDouble('QspScale')[1]
  options.QUspScale : float = thisScript.GetInputParameterDouble('QUspScale')[1]
  options.QPFspScale : float = thisScript.GetInputParameterDouble('QPFspScale')[1]
  options.QPFmode : int = 0 
  options.paraEventsOnly : bool = bool(thisScript.GetInputParameterInt('paraEventsOnly')[1]) 
  options.consolidate : bool = bool(thisScript.GetInputParameterInt('consolidate')[1]) 
  options.parallelComp : bool = bool(thisScript.GetInputParameterInt('parallelComp')[1])

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
  grid.pCtrl = networkData.SearchObject('PP-MTB\\pCtrl.ElmSecctrl')
  grid.qCtrl = networkData.SearchObject('PP-MTB\\qCtrl.ElmStactrl') 
  grid.sigGen = networkData.SearchObject('PP-MTB\\signalGenerator.ElmDsl')
  grid.connector = networkData.SearchObject('PP-MTB\\connector.ElmComp')
  grid.sink = networkData.SearchObject('PP-MTB\\lib\\connector\\sink.BlkSlot')
  return grid

def loadPlantInfo(options) -> SimpleNamespace:
  pdInput = pd.read_excel(open(options.file,'rb'), sheet_name='Input', header = 0, index_col=None, usecols=lambda x: 'Unnamed' not in x)
  plantInfo = SimpleNamespace()
  plantInfo.PN : float = pdInput['Pn (MW)'][0] # Nominal active power (MW)
  plantInfo.VN : float = pdInput['Vn_PoC (kV)'][0] # PoC nominal voltage (LL, RMS, kV)
  plantInfo.IN : float = plantInfo.PN/math.sqrt(3)/plantInfo.VN
  plantInfo.SCR : float = pdInput['SCR'][0] # Minimum SCR at PoC
  plantInfo.XRRATIO : float = pdInput['X/R'][0] # XR ratio               
  plantInfo.PROJECTNAME : str = pdInput['ProjectName'][0] # Project name
  plantInfo.Qmode : int = pdInput['Qmode_Q'][0]
  plantInfo.QUmode : int = pdInput['Qmode_V'][0]
  plantInfo.QPFmode : int = pdInput['Qmode_PF'][0]

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
  case.Included = bool(pdCases['Included'][caseIndex])
  case.ION : str = str(pdCases['ION'][caseIndex])
  case.Rank : int = int(pdCases['Rank'][caseIndex])
  case.TestType : str = str(pdCases['TestType'][caseIndex])
  case.P0 : float = float(pdCases['P0'][caseIndex])
  case.InitValue : float = float(pdCases['InitValue'][caseIndex])
  case.Qmode : int = int(pdCases['Qmode'][caseIndex]) 
  
  if case.Qmode == plantInfo.Qmode:
    case.internalQmode = 0
  elif case.Qmode == plantInfo.QUmode:
    case.internalQmode = 1
  elif case.Qmode == plantInfo.QPFmode:
    case.internalQmode = 2
  else:
    app.PrintWarn('Invalid Qmode: {}. Assuming Q control.'.format(case.Qmode))
    case.internalQmode = 0
  try:
    case.U0 : float = float(pdCases['U0'][caseIndex])
  except:
    case.U0 : float = 1.0
  case.GridImped : float = float(pdCases['GridImped'][caseIndex])
  case.Tstop : float = float(pdCases['Tstop (s)'][caseIndex])
  case.PrefCtrl : bool = bool(pdCases['PrefCtrl'][caseIndex])
  case.QrefCtrl : bool = bool(pdCases['QrefCtrl'][caseIndex])
  case.VSPhaseCtrl : bool = bool(pdCases['VSPhaseCtrl'][caseIndex])
  case.VSVoltCtrl : bool = bool(pdCases['VSVoltCtrl'][caseIndex])
  case.VSFreqCtrl : bool = bool(pdCases['VSFreqCtrl'][caseIndex])
  case.FaultType : int = int(pdCases['FaultType'][caseIndex])
  case.FaultPeriod : float = float(pdCases['FaultPeriod'][caseIndex])
  case.FaultDepth : float = float(pdCases['FaultDepth'][caseIndex]) #Misleading name. Residual voltage.
  case.events = list()
  evLastEvent = -math.inf
  eIndex = 1
  while True:
    startLabel = 'C{}start'.format(eIndex)
    spLabel = 'C{}setpoint'.format(eIndex)
    rpLabel = 'C{}ramp'.format(eIndex)
    if {startLabel, spLabel, rpLabel}.issubset(pdCases.columns):
      evStart = float(pdCases[startLabel][caseIndex])
      evSp = float(pdCases[spLabel][caseIndex])
      evRp = float(pdCases[rpLabel][caseIndex])
    if math.isnan(evStart) or math.isnan(evSp) or math.isnan(evRp) or evStart >= case.Tstop or evStart <= evLastEvent:
      break
    evLastEvent = evStart
    case.events.append([evStart,evSp,evRp])
    eIndex += 1
  return case

def loadCases(app, options) -> pd.DataFrame:
  app.PrintPlain('Read Test Matrix: {}'.format(options.file)) 
  return pd.read_excel(open(options.file,'rb'), sheet_name='Cases', header = 0, index_col=None, usecols=lambda x: 'Unnamed' not in x)

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

  for qvar in (options.QmodeVar, options.QUmodeVar, options.QPFmodeVar):
    if qvar is not None and qvar.GetAttribute('e:tToAc') >= studyTime:
      app.PrintWarn('The Q control mode variation {} is active after or at the same time as the base case.'.format(qvar.GetFullName(0)))
    if not options.consolidate and qvar is not None:
      for qstage in qvar.GetContents('*.IntSstage'):
        if qstage.GetAttribute('e:tAcTime') == studyTime:
          exit('The Q control mode variation stage {} is active at the same time as the the base case.'.format(qstage.GetFullName(0)))

  recordingStage = app.GetRecordingStage()

  if not options.consolidate and recordingStage is not None and recordingStage.GetAttribute('e:tAcTime') == studyTime:
    exit('Expansionstage {} conflicts with PP-MTB setup.'.format(recordingStage.GetFullName(0))) 

  # Create version
  if options.backup:
    project.CreateVersion('MTB_{}'.format(datetime.now().strftime(r'%d%m%Y%H%M%S')))

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
  maxRank = max(cases['Rank'])
  
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
  ctrlMode : int = int(subScripts.setupPlots.GetInputParameterInt('ctrlMode')[1])
  Qmode : int = int(subScripts.setupPlots.GetInputParameterInt('Qmode')[1])
  symSim : bool = bool(subScripts.setupPlots.GetInputParameterInt('symSim')[1])
  inputSignal : str = str(subScripts.setupPlots.GetInputParameterString('inputSignal')[1])
  inputScaling : float = float(subScripts.setupPlots.GetInputParameterDouble('inputScaling')[1])
  inputBlock = subScripts.setupPlots.GetExternalObject('inputBlock')[1]

  setupPlots(app, eventPlot, faultPlot, phasePlot, uLim, ctrlMode, Qmode, symSim, inputBlock, inputSignal, inputScaling)

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

if __name__ == "__main__":
  main()