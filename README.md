
# MTB (Model Test Bench)

Connecting new electricity generation and demand facilities to Denmark's public transmission and distribution systems requires thorough grid compliance studies using both RMS/PDT and EMT plant-level models. The Danish TSO, Energinet, mandates that RMS/PDT models be created in [DIgSILENT PowerFactory](https://www.digsilent.de/en/powerfactory.html)  and EMT models in [PSCAD](https://www.pscad.com/). Before any facility can begin operation, all electrically significant plants must have their RMS and EMT models reviewed and approved by Energinet to ensure both grid compliance and model quality. Conducting the necessary studies to demonstrate compliance and validate model quality through comparisons of RMS and EMT models can be both time-consuming and prone to error.

The MTB (Model Test Bench) simplifies and automates this process by enabling seamless grid connection studies across PowerFactory and PSCAD environments. Energinet relies on the MTB for all grid connection studies and strongly recommends its use to all connecting parties. By using the MTB, developers can conduct studies under the exact same conditions as Energinet, ensuring they achieve the same results that Energinet will evaluate.

The workflow is simple:

1. **Define the Required Studies** in the provided Excel sheet. The MTB is preconfigured for the studies required in most grid connection cases in Denmark but is also adaptable to all regions following the EU RfG. Modifying or extending the study case set is straightforward.
2. **Integrate the PSCAD MTB Component** into the plant's PSCAD model.
3. **Integrate the PowerFactory MTB Component** into the plant's PowerFactory model.
4. **Execute Simulations** using the MTB Python scripts.
5. **Visualize the Results** with the included plotter tool. The example below shows a 

For the latest release notes, please visit the [Releases page](https://github.com/Energinet-AIG/MTB/releases). Learn more about the regulations for grid connection of new facilities in Denmark: [Danish](https://energinet.dk/regler/el/nettilslutning) or [English](https://en.energinet.dk/electricity/rules-and-regulations/regulations-for-new-facilities).

![96](https://github.com/user-attachments/assets/6ce6746c-83b6-4d3f-a433-71c7ce5409de)
*Example comparative study between RMS (red) and EMT (blue) models.*
## Getting Started

To start using the MTB, refer to the Quickstart Guides available on the [MTB wiki Home page](https://github.com/Energinet-IG/MTB/wiki) on GitHub. These guides provide instructions on using the Casesheet, PowerFactory, PSCAD, and the plotter tool.

## Requirements

To install all necessary dependencies, run:

```bash
pip install -r requirements.txt
```

### Tested Environments

- **PowerFactory**: Tested on version 2024 SP4 with Python versions >= 3.8.8.
- **PSCAD**: Tested on version 5.0.1.0 with Python 3.7.2 (embedded Python). Compatibility is guaranteed only with Intel Fortran Compilers.

## Contributing

We welcome contributions! To contribute, please file an issue via the MTB [Issues tab](https://github.com/Energinet-AIG/MTB/issues). You can report bugs, request features, or suggest improvements. Before submitting, please check for any known issues.

## Contact

For inquiries, please contact the Energinet simulation model team: simuleringsmodeller@energinet.dk
