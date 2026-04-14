import argparse
import os
import logging
import pandas as pd
from datetime import datetime



def build_csv(srcfile,outdir,sheet_name,column_name,column_separator):
    """
    Build a CSV based on values in an excel file column
    """
    print(f'...Building CSV file')

    #  build the csv_file name based on the srcfile basename
    os.makedirs(outdir, exist_ok=True)
    csv_file = os.path.join(outdir,f'{os.path.splitext(os.path.basename(srcfile))[0]}.txt')
    df1=pd.read_excel(srcfile,sheet_name=sheet_name, usecols=[column_name], dtype={column_name:str})
    
    # Split data in column_name on a semi colon and add each list entry into a new single dataframe
    # "Serum;Plasma" is treated as a single value (either/or) — protect it before splitting
    _PLACEHOLDER = '\x00SERUM_PLASMA\x00'
    code_series = (
        df1[column_name]
        .dropna()
        .astype(str)
        .str.replace('Serum;Plasma', _PLACEHOLDER, regex=False)
        .str.replace('Plasma;Serum', _PLACEHOLDER, regex=False)
        .str.split(';')
        .explode()
        .str.strip()
        .str.replace(_PLACEHOLDER, 'Serum;Plasma', regex=False)
    )
    code_series = code_series[code_series.ne('')]
    cs_df = pd.DataFrame({'display': code_series})

    # Remove duplicates
    unique_df = cs_df.drop_duplicates(subset=['display'], keep='first')

    # Sort the DataFrame
    cs_sorted_df = unique_df.sort_values(by='display', kind='stable').reset_index(drop=True)
    cs_sorted_df.insert(0, 'code', cs_sorted_df.index + 1)

    # Output as csv_file with column_separator as the csv separator
    cs_sorted_df.to_csv(csv_file, sep=column_separator, index=False)

    return csv_file

def main():
    """
    Create a CSV for SNAP2SNOMED import from an RCPA XLS File
    """
    
    homedir=os.environ['HOME']
    parser = argparse.ArgumentParser()
    def_indir=os.path.join(homedir,"data","rcpa","in")
    def_outdir=os.path.join(homedir,"data","rcpa","out")
    def_sheet_name = "RCPA SPIA Requesting_Mar 2026"
    def_col_name = "Specimen"
   

    logger = logging.getLogger(__name__)
    parser.add_argument("-i", "--indir", help="infile path for excel file", default=def_indir)
    parser.add_argument("-o", "--outdir", help="dir name for file output", default=def_outdir)
    parser.add_argument("-s", "--sheet_name", help="sheet name for excel data", default=def_sheet_name)
    parser.add_argument("-c", "--column_name", help="excel column name to be extracted", default=def_col_name)
    args = parser.parse_args()
    now = datetime.now() # current date and time
    ts = now.strftime("%Y%m%d-%H%M%S")
    logsdir=os.path.join(homedir,"data","ucum","logs")
    FORMAT='%(asctime)s %(lineno)d : %(message)s'
    os.makedirs(logsdir, exist_ok=True)
    logging.basicConfig(format=FORMAT, filename=os.path.join(logsdir,f'xls2cs-{ts}.log'),level=logging.INFO)
    logger.info('Started')

    ## Do a readdir on args.indir and process each file, passing the file name into build_csv
    for entry in sorted(os.listdir(args.indir)):
        infile = os.path.join(args.indir, entry)
        if not os.path.isfile(infile) or entry.startswith('~$'):
            continue
        if not entry.lower().endswith(('.xls', '.xlsx', '.xlsm', '.xlsb')):
            continue

        logger.info(f'Processing {infile}...')
        build_csv(infile,args.outdir,args.sheet_name,args.column_name,"\t")
    logger.info("Finished")

if __name__ == '__main__':
    main()