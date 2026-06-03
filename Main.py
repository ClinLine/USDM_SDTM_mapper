import argparse
import os
import json
import openpyxl
import TS
import TI
import TE
import TA
import TV
from create_define import Create_Define_XML, Create_Define_json


def run_conversion(json_input, map_input, output, usdm_version):
    code_lists_map = {}
    all_vars= {}
    wb = openpyxl.load_workbook(map_input)
    ts_var = TS.Create_TS(wb, json_input)
    all_vars["TS"] = ts_var
    ti_var, ti_codes = TI.Create_TI(wb, json_input)
    all_vars["TI"] = ti_var
    TE.Create_TE(wb, json_input)
    ta_var, ta_codes = TA.Create_TA(wb, json_input)
    all_vars["TA"] = ta_var
    TV.Create_TV(wb, json_input)
    code_lists_map.update(ta_codes)
    code_lists_map.update(ti_codes)
    Create_Define_XML(wb, all_vars,code_lists_map)
    
    # Generate define.json in the output directory - temporary test output for now, will be removed in the future when the define.json is fully implemented and tested
    define_json = Create_Define_json(wb, code_lists_map)
    output_dir = os.path.dirname(output)
    define_json_path = os.path.join(output_dir, "define.json")
    with open(define_json_path, "w", encoding="utf-8") as f:
        json.dump(define_json, f, indent=2)
    
    wb.save(output)
    wb.close()
    print(f"Output saved to {output}")
    print(f"Define JSON saved to {define_json_path}")


def main():
    parser = argparse.ArgumentParser(prog="main", description="USDM to SDTM TDM Generator")
    subparsers = parser.add_subparsers(dest="command", required=True)

    # 'generate' command
    gen_parser = subparsers.add_parser("generate", help="Generate SDTM TDM from USDM JSON input")
    gen_parser.add_argument("-v", "--usdm-version", required=True, help="USDM version (e.g., 4-0)")
    gen_parser.add_argument("-i", "--input", required=True, help="Input JSON file")
    gen_parser.add_argument("-m", "--map", default="Maps/sdtm_mapping_paths.xlsx", help="Mapping file")
    gen_parser.add_argument("-o", "--output", default="Output/output.xlsx", help="Output file")

    args = parser.parse_args()

    if args.command == "generate":
        print(f"Converting USDM v{args.usdm_version} to SDTM TDM")
        run_conversion(args.input, args.map, args.output, args.usdm_version)


if __name__ == "__main__":
    main()
