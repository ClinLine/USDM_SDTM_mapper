# USDM to SDTM Trial Design Domain Tool
The tool shows how the mappings, transformed to jsonata requests can inform the actual creation of the SDTM trial summary datasets TA, TE, TV, TI, and TS.
The tool is based on json, jsonata and opepyxl packages to enable the following steps:
- Read the defined jsonata requests from the input Excel file
- Open the USDM API file in json format
- Run the jsonata requests on the USDM API file
- Process the request results to create valid output including appropriate sequence numbering and grouping
- Output the results to a result output Excel file
  
The current version includes the creation of all Trial design domains. Some additional functionality like HTML parsing and reference resolution will be added soon.

We presented this tool at the CDISC COSA webinar on 23 September 2025 which explains the tool in more detail. See [Youtube video](https://youtu.be/j0myfrOjCcs)

## Input
The input for the tool is based on the [SDTM mappings available in the CDISC DDF Github](https://github.com/cdisc-org/DDF-RA/blob/main/Documents/Mappings/sdtm_mapping.xlsx)
The mappings in this file are converted to Jsonata and added to the input Excel [file sdtm_mapping_paths.xlsx](https://github.com/ClinLine/SDTM_mapper/blob/main/Maps/sdtm_mapping_paths.xlsx)
These mappings will be used by the mapping tool and include:
- Actual result mapping
- Null flavour indication
- Code mapping, if applicable
- Code System mapping, if applicable
- Code System version mapping, if applicable

More jsonata mapping will be added until complete.

## Tool Functionality
For running the python code install the following packages:
 - Jsonata-Python:  pip install jsonata-python
 - Openpyxl: pip install Openpyxl

## Output
The output SDTM datasets TA, TE, TV, TI, and TS is added in the same format as the original input Excel file. The file is stored as [sdtm_mapping_results.xlsx](https://github.com/ClinLine/SDTM_mapper/blob/main/Output/sdtm_mapping_results.xlsx).
The results for the TS Summary parameters, if not empty, are added to the TS sheet including the corresponding StudyId.

## Acknowledgements
This ClinLine open source tool is created by Noah Brezet and Berber Snoeijer
Please contact us via info@clinline.eu if you like to learn more and/or support for integration of its features.
