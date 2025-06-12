# USDM to SDTM Trial Summary Tool

## Input
The input for the tool is based on the [SDTM mappings available in the CDISC DDF Github] (https://github.com/cdisc-org/DDF-RA/blob/main/Documents/Mappings/sdtm_mapping.xlsx)
The mappings in this file are converted to Jsonata and added to the input Excel [file sdtm_mapping_paths.xlsx] (https://github.com/ClinLine/SDTM_mapper/blob/main/Maps/sdtm_mapping_paths.xlsx)
These mappings will be used by the mapping tool and include:
- Actual result mapping
- Null flavour indication
- Code mapping, if applicable
- Code System mapping, if applicable
- Code System version mapping, if applicable
More jsonata mapping will be added until complete.

For running the python code install the following packages:

Jsonata-Python:
pip install jsonata-python

Openpyxl:
$pip install Openpyxl
