# MTB

Connecting new electricity generation and demand facilities to Denmark's public transmission and distribution systems requires comprehensive grid compliance studies using both RMS/PDT and EMT plant-level models. The Danish TSO, Energinet, mandates that RMS/PDT models be developed in [DIgSILENT PowerFactory](https://www.digsilent.de/en/powerfactory.html) and EMT models in [PSCAD](https://www.pscad.com/).

The MTB (**M**odel **T**est **B**ench) simplifies and automates the process of conducting grid connection studies across both PowerFactory and PSCAD environments. The workflow is straightforward:

1. **Define Studies**: Use an Excel sheet to define the required studies and any custom cases. The MTB comes with predefined studies necessary for both TSO and DSO-level grid connection studies in Denmark.
2. **Integrate PSCAD Component**: Integrate the MTB PSCAD component into the plant's PSCAD model.
3. **Integrate PowerFactory Component**: Integrate the MTB PowerFactory component into the plant's PowerFactory model.
4. **Execute Simulations**: Automatically run the required simulations in both PowerFactory and PSCAD environments.
5. **Visualize Results**: Use the included plotter tool to visualize and analyze the results.

For the latest release notes, please visit [Releases](https://github.com/Energinet-AIG/MTB/releases). Learn more about the regulations for grid connection of new facilities in Denmark: [Danish](https://energinet.dk/regler/el/nettilslutning) or [English](https://en.energinet.dk/electricity/rules-and-regulations/regulations-for-new-facilities).

## Getting Started

To begin using the MTB, please refer to the Quickstart Guides available on the [MTB wiki Home page](https://github.com/Energinet-IG/MTB/wiki) on GitHub. These guides provide instructions on using the Casesheet, PowerFactory, PSCAD, and the plotter tool.

## Requirements

Install all necessary dependencies by running:

pip install -r requirements.txt


### Tested Environments

- **PowerFactory**: Tested on version 2024 SP4 with Python versions >= 3.8.8.
- **PSCAD**: Tested on version 5.0.1.0 with Python 3.7.2 (embedded Python). Compatibility is guaranteed only with Intel Fortran Compilers.

## Contribution

We welcome contributions! To contribute, please file an issue via the MTB [Issues](https://github.com/Energinet-AIG/MTB/issues) tab. You can report bugs, request features, or suggest improvements. Before submitting, please check for any known issues.

## Contact

For inquiries, please contact:
- **Senior Engineer Mathias Kristensen**: mkt@energinet.dk
- **Engineer Casper Lindgaard**: cvl@energinet.dk
