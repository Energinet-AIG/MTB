# MTB
Connecting new electricity generation & demand facilities to the public transmission and distribution systems in Denmark requires grid compliance studies with both RMS/PDT and EMT plant level models. The danish TSO Energinet requires RMS/PDT models in [DIgSILENT Powerfactory](https://www.digsilent.de/en/powerfactory.html) and EMT models in [PSCAD](https://www.pscad.com/).

  Energinets MTB (**M**odel **T**est **B**ench) is a tool for automation of studycase setup and simulation in both PowerFactory and PSCAD with external visualizing of results. The MTB is meant as a tool to help guide in checking simulation and plant performance of RMS/PDT- and EMT-models in regards to the danish grid code and the requirements for simulation models. A set of predefined cases are available with the option to add custom cases or remove exisiting ones.
  The MTB, originally an internal Energinet tool, has been open-sourced as an strategic initiative to support the grid connecting parties. 

  Latest release notes can be found under [Releases](https://github.com/Energinet-AIG/MTB/releases).
  
  Read more about the regulations for grid connection of new facilities here: [danish](https://energinet.dk/regler/el/nettilslutning) or [english](https://en.energinet.dk/electricity/rules-and-regulations/regulations-for-new-facilities).

## Getting Started
  To get started, follow the Quickstart Guides on the MTB wiki [Home](https://github.com/Energinet-IG/MTB/wiki) page of the [MTB GitHub](https://github.com/Energinet-AIG/MTB). Here you will find guides for the Casesheet, PowerFactory, PSCAD and the plotter.

## Requirements
  Dependencies are installed by running `pip install -r requirements.txt`. 

### Tested environments



The Powerfactory tool has been tested in the following environments and dependency versions as listed in requirements.txt:
* 2024 SP4 with Python version >= 3.8.8 

### Tested PSCAD environments
The PSCAD tool has been tested with in following environments and dependency versions as listed in requirements.txt:
* 5.0.2.0 with Python 3.7.2 (embedded python)

#### Tested Fortran Compilers
Intel(R) Visual Fortran Compiler XE:
* 12.1.7.371
* 15.0.1.148
* 15.0.1.148 (64-bit)
* 15.0.5.280
* 15.0.5.280 (64-bit)
  
## Contribution
  If you are interested in contributing, please feel free to file an issue. This is done by using the MTB [Issues](https://github.com/Energinet-AIG/MTB/issues) tab. Here you can report bugs, feature requests or improvements, but please check for known issues beforehand. 

  When you file an issue, please try making it as specific and independent of other issues as possible. Make use of the Labels to hightlight what problem or tool the issue revolves arround. We encourage you to contribute with any bug, improvement or idea you might come across to help make this tool as useful and user-friendly as possible.
  
## Contact
  Inquiries can be directed at 
  * Senior engineer Mathias Kristensen at mkt@energinet.dk 
  * Engineer Casper Lindgaard at cvl@energinet.dk
