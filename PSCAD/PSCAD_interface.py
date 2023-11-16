#author: mkt
from jinja2 import Environment, BaseLoader
from typing import Union
from warnings import warn
from math import isnan
from copy import copy

class waveform:
    def __init__(self, s0 : float, r0 : float = 0.0):
        s_ = float(s0)
        r_ = float(r0)
        self.__t__ = [0.0]
        self.__s__ = [s_]
        self.__r__ = [r_]
        self.type = 0

    def __getitem__(self, key):
        return (self.__t__[key], self.__s__[key], self.__r__[key])

    def __delitem__(self, key):
        del self.__t__[key]
        del self.__s__[key]
        del self.__r__[key]

    def add(self, t : float, s : float, r : float = 0.0):
        t_ = float(t)
        s_ = float(s)
        r_ = float(r)
        for i in range(len(self.__s__) - 1, -2, -1):
            if i < 0 or self.__t__[i] <= t_:
                newIndex = i + 1
                leftIndex = max(0, i)

                if s_ is None or isnan(s_):  
                    self.__s__.insert(newIndex, self.__s__[leftIndex])
                else:
                    self.__s__.insert(newIndex, s_)

                if r_ is None or isnan(r_):
                    self.__r__.insert(newIndex, self.__r__[leftIndex])
                else:
                    self.__r__.insert(newIndex, r_)

                self.__t__.insert(newIndex, t_)
                return self.__t__

    def t(self, minLength = None):
        if minLength is None:
            return self.__t__
        elif len(self.__t__) >= minLength:
            return self.__t__
        else:
            return self.__t__ + (minLength - len(self.__t__)) * [0.0]

    def s(self, minLength = None):
        if minLength is None:
            return self.__s__
        elif len(self.__s__) >= minLength:
            return self.__s__
        else:
            return self.__s__ + (minLength - len(self.__s__)) * [0.0]

    def r(self, minLength = None):
        if minLength is None:
            return self.__r__
        elif len(self.__r__) >= minLength:
            return self.__r__
        else:
            return self.__r__ + (minLength - len(self.__r__)) * [0.0]

    def len(self):
        return len(self.__t__)

    def __eq__(self, other):
        if type(other) == waveform:
            return self.__t__ == other.__t__ and self.__s__ == other.__s__ and self.__r__ == other.__r__
        else:
            return False

class recordedWaveform:    
    def __init__(self, path : str, length : float = None):
        self.path = path
        self.type = 1
        self.len = length

    def updateLen(self):
        exit('updateLen not implemented')

    def __eq__(self, other):
        if type(other) == recordedWaveform:
            return self.path == other.path and self.len == other.len
        else:
            return False

class constant:
    def __init__(self, name, value):
        self.__name__ = name
        self.__value__ = value
        self.__signalTemplate__ = \
f"""subroutine {name}_const(y)
    implicit none
    real, intent(out) :: y

    y = {value}

end subroutine {name}_const"""
    def value(self):
        return self.__value__

    def render(self):
        return self.__signalTemplate__

class signal:
    def __init__(self, name : str):
        self.__name__ = name
        self.__waves__ = dict()
        self.__hasWaveforms__ = False
        self.__hasRecWaveforms__ = False
        self.__default__ = None
        self.__hasDefault__ = False
        self.__signalTemplate__ = \
"""subroutine {{ signal.name() }}_signal(rank, y)
    include 'emtstor.h'
    include 's1.h'  
    implicit none

    integer, intent(in) :: rank
    real, intent(out) ::  y
    integer :: inrank{% if signal.hasWaveforms() %}, events, index{% endif %}

    {% if signal.hasWaveforms() %}
    real  :: tx({{ signal.arraySize() }})
    real :: sy({{ signal.arraySize() }})
    real :: ry({{ signal.arraySize() }})
    {% endif %}

    {% if signal.hasRecWaveforms() %}
    real :: frout(11) 
    {% endif %}

    if(TIMEZERO) then
        {% if signal.hasWaveforms() %}
        index = 1
        {% endif %}
        inrank = rank
        STORI(NSTORI + 1) = inrank
    else    
        {% if signal.hasWaveforms() %}
        index = STORI(NSTORI)
        {% endif %}
        inrank = STORI(NSTORI + 1)
    endif
    
    ranks: select case(inrank)
    {% for rank in signal.ranks() %}
    {% if not signal[rank] == signal.default() %}
    {% if signal[rank].type == 0 %}
    case({{ rank }})
        tx = {{ signal[rank].t(signal.arraySize()) }}
        sy = {{ signal[rank].s(signal.arraySize()) }}
        ry = {{ signal[rank].r(signal.arraySize()) }}
        events = {{ signal[rank].len() }}
    {% else %}
    case({{ rank }})
        NSTORI = NSTORI + 2
        call FILEREAD2("{{ signal[rank].path }}", 0, 2, 0, 0, 0.0, 1.0, 0.0, frout)
         y = frout(2)
        {% if signal.hasWaveforms() %}
        events = -1
        {% endif %}
    {% endif %}
    {% endif %}
    {% endfor %}
    {% if signal.hasDefault() %}
    case default 
    {% if signal.default().type == 0 %}
    tx = {{ signal.default().t(signal.arraySize()) }}
    sy = {{ signal.default().s(signal.arraySize()) }}
    ry = {{ signal.default().r(signal.arraySize()) }}
    events = {{ signal.default().len() }}    
    {% else %}
    NSTORI = NSTORI + 2
    call FILEREAD2("{{ signal.default().path }}", 0, 2, 0, 0, 0.0, 1.0, 0.0, frout)
     y = frout(2)
    {% if signal.hasWaveforms() %}
    events = -1
    {% endif %}    
    {% endif %}
    {% endif %}
    end select ranks

    {% if signal.hasWaveforms() %}
    if( events > -1) then
        if( ( index < events ) .and. (TIME >= tx(index + 1)) ) then
            index = index + 1
        endif
         y =  sy(index) + (TIME - tx(index)) * ry(index)
        STORI(NSTORI) = index
        NSTORI = NSTORI + 5
        NSTORF = NSTORF + 35
    endif  
    {% elif not signal.hasRecWaveforms() %}
    NSTORI = NSTORI + 5
    NSTORF = NSTORF + 35
    {% endif %}
end subroutine {{ signal.name() }}_signal"""

    def name(self):
        return self.__name__

    def __setitem__(self, rank, wave ):
        self.addRank(rank, wave)

    def __getitem__(self, rank):
        return self.__waves__[rank]

    def __delitem__(self, rank):
        del self.__waves__[rank]

    def addRank(self, rank : int, wave : Union[ waveform, recordedWaveform ]):
        if type(rank) != int:
            exit(f'Non integer rank added to signal {self.name()}')
        
        self.__waves__[rank] = wave
        if wave.type == 0:
            self.__hasWaveforms__ = True
        else:
            self.__hasRecWaveforms__ = True

    def addDefault(self, wave : Union[ waveform, recordedWaveform ]):
        self.__default__ = wave
        self.__hasDefault__ = True
        if type(wave) == waveform:
            self.__hasWaveforms__ = True
        elif type(wave) == recordedWaveform:
            self.__hasRecWaveforms__ = True

    def default(self):
        return self.__default__

    def hasDefault(self):
        return self.__hasDefault__
    
    def hasRecWaveforms(self):
        return self.__hasRecWaveforms__
    
    def hasWaveforms(self):
        return self.__hasWaveforms__

    def arraySize(self):
        maxLength = -1
        if self.__hasWaveforms__:
            for k in self.__waves__.keys():
                if self.__waves__[k].type == 0:
                    maxLength = max(self.__waves__[k].len(), maxLength)
            if self.__hasDefault__ and type(self.__default__) == waveform:
                maxLength = max(self.__default__.len(), maxLength)
        return maxLength

    def ranks(self):
        return self.__waves__.keys()

    def __groupRanks__(self):
        groups = []
        for rank in self.ranks():
            foundGroup = False
            for group in groups:
                if self[rank] == group[0][1]:
                    group.append((rank, self[rank]))
                    foundGroup = True
                    continue

            if not foundGroup:
                groups.append([(rank, self[rank])])
        

        groupedSignal = copy(self)
        groupedSignal.__waves__ = dict()

        for group in groups:
            group = sorted(group)
            
            groupIndex = lastPlaced = group[0][0]
            ranks = f'{group[0][0]}'    

            for wave in group[1:]:
                if wave[0] == groupIndex + 1:
                    groupIndex = wave[0]
                else:
                    if groupIndex == lastPlaced:
                        ranks += f', {wave[0]}'
                    elif groupIndex == lastPlaced + 1:
                        ranks += f',{groupIndex}, {wave[0]}'
                    else:
                        ranks += f':{groupIndex}, {wave[0]}'

                    groupIndex = lastPlaced = wave[0]
                    
            if groupIndex == lastPlaced + 1:
                ranks += f',{groupIndex}'
            elif groupIndex > lastPlaced:
                ranks += f':{groupIndex}'

            groupedSignal.__waves__[ranks] = group[0][1]

        return groupedSignal
        

    def render(self):
        template = Environment(loader=BaseLoader,trim_blocks=True,lstrip_blocks=True).from_string(self.__signalTemplate__)
        return template.render(signal = self.__groupRanks__())

class timeInvariant:
    def __init__(self, name):
        self.__name__ = name
        self.__values__ = dict()
        self.__default__ = None
        self.__hasDefault__ = False
        self.__signalTemplate__ = \
"""subroutine {{ timeInv.__name__ }}_tinv(rank, y)
    implicit none
    integer, intent(in) :: rank
    real, intent(out) ::  y

    ranks: select case(rank)
    {% for rank in timeInv.ranks() %}
    {% if not timeInv[rank] == timeInv.default() %}
    case({{ rank }})
         y = {{ timeInv[rank] }}
    {% endif %}
    {% endfor %}
    {% if timeInv.hasDefault() %}
    case default 
         y = {{ timeInv.default() }}
    {% endif %}
    end select ranks
end subroutine {{ timeInv.name() }}_tinv"""            

    def name(self):
        return self.__name__

    def __getitem__(self, rank):
        return self.__values__[rank]

    def __setitem__(self, rank, value):
        self.__values__[rank] = value

    def __delitem__(self, rank):
        del self.__values__[rank]

    def addRank(self, rank : int, value : float):
        if type(rank) != int:
            exit(f'Non integer rank added to timeinvariant {self.name()}')
        self.__values__[rank] = value

    def addDefault(self, value : float):
        self.__default__ = value
        self.__hasDefault__ = True

    def default(self):
        return self.__default__

    def hasDefault(self):
        return self.__hasDefault__

    def ranks(self):
        return self.__values__.keys()

    def __groupRanks__(self):
        groups = []
        for rank in self.ranks():
            foundGroup = False
            for group in groups:
                if self[rank] == group[0][1]:
                    group.append((rank, self[rank]))
                    foundGroup = True
                    continue

            if not foundGroup:
                groups.append([(rank, self[rank])])
        

        groupedTimeinvariant = copy(self)
        groupedTimeinvariant.__values__ = dict()

        for group in groups:
            group = sorted(group)
            
            groupIndex = lastPlaced = group[0][0]
            ranks = f'{group[0][0]}'    

            for wave in group[1:]:
                if wave[0] == groupIndex + 1:
                    groupIndex = wave[0]
                else:
                    if groupIndex == lastPlaced:
                        ranks += f', {wave[0]}'
                    elif groupIndex == lastPlaced + 1:
                        ranks += f',{groupIndex}, {wave[0]}'
                    else:
                        ranks += f':{groupIndex}, {wave[0]}'

                    groupIndex = lastPlaced = wave[0]
                    
            if groupIndex == lastPlaced + 1:
                ranks += f',{groupIndex}'
            elif groupIndex > lastPlaced:
                ranks += f':{groupIndex}'

            groupedTimeinvariant.__values__[ranks] = group[0][1]

        return groupedTimeinvariant

    def render(self):
        template = Environment(loader=BaseLoader,trim_blocks=True,lstrip_blocks=True).from_string(self.__signalTemplate__)
        return template.render(timeInv = self.__groupRanks__())   

class fortranFile:
    def __init__(self, path):
        self.__path__ = path
        self.__symbols__ = dict()

    def __getitem__(self, name):
        return self.__symbols__[name]

    def __delitem__(self, name):
        del self.__symbols__[name]

    def __add__(self, symbol : Union[signal, timeInvariant, constant]):
        existingSymbols = self.symbols()
        if symbol.__name__ in existingSymbols:
            warn(f'"{symbol.__name__}" already exists in {self.__path__}. Existing symbol overwritten.')
        self.__symbols__[symbol.__name__] = symbol
        return symbol

    def newSignal(self, name : str, default : Union[waveform, recordedWaveform] = None) -> signal:
        newestSignal = signal(name)
        if type(default) == waveform or type(default) == recordedWaveform: 
            newestSignal.addDefault(default)
        return self.__add__(newestSignal)

    def newConstant(self, name : str, value : float) -> constant:
        return self.__add__(constant(name, value))

    def newTimeInv(self, name : str, default : float = None) -> timeInvariant:
        newestTi = timeInvariant(name)
        if default != None:
            newestTi.addDefault(default)
        return self.__add__(newestTi)
            
    def symbols(self):
        return self.__symbols__.keys()

    def remove(self, *names : str):
        for name in names:
            del self.__symbols__[name]

    def render(self):
        result = ''

        for name in self.symbols():
            result += self.__symbols__[name].render() + '\n\n'

        with open(self.__path__, mode='w') as f:
            f.write(result)
            f.close()