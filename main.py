"""
Created on 19 Jun 2022

@author: turch
"""

import argparse
from set_parameters import Begin
import pathlib
from pathlib import Path


def init():
    parser = argparse.ArgumentParser(
        description="Automatic translation module with glossary support"
    )
    current_file_location = pathlib.Path(__file__).parent.resolve()
    gloss_path = pathlib.Path(current_file_location / "glossary.xlsx")

    parser.add_argument(
        "-d",
        "--data_path",
        type=str,
        default=r"C:\Users\turch\OneDrive\Documents\0CAAL\KG",
        help="Path to folder/file with data to translate",
    )
    parser.add_argument(
        "-il",
        "--input_language",
        type=str,
        choices=["ru", "en"],
        default="ru",
        help="Language of the input data (options: 'ru', 'en'), default is Russian (ru)",
    )
    parser.add_argument(
        "-ol",
        "--output_language",
        type=str,
        default="en",
        help="Language to translate the data into, default is English (en)",
    )
    parser.add_argument(
        "-g",
        "--glossfile",
        type=str,
        default="glossary.xlsx",
        help="Path to the glossary file (e.g., glossary.xlsx)",
    )

    parser.add_argument(
        "-s",
        "--sheet",
        type=str,
        choices=["Monument", "Archive", "other excel", "other filetype"],
        default="Archive",
        help="Sheet name to translate, default is Archives",
    )
    parser.add_argument(
        "-ic",
        "--input_columns",
        nargs="+",
        type=str,
        default="",
        help="The columns to be translated, default for 'Archives' is C, E, G, H, J; for 'Monuments' is B, C, D",
    )
    '''parser.add_argument(
        "-oc",
        "--output_columns",
        nargs="+",
        type=str,
        required=True,
        help="The columns to insert translated data into. Please type in corresponding order to input columns", 
    )''' # Removed as not needed for CSV output
    parser.add_argument(
        "-r",
        "--start_row",
        type=int,
        default=5,
        help="Row to start translation from (excluding column names, default: 5)",
    )

    parser.add_argument(
        "-od", 
        "--output_directory",
        type=Path,
        default=Path.cwd(),
        help="Path to CSV output directory",
    )

    args = parser.parse_args()
    if args.sheet == "Monuments":
        sheet = "Data Sheet"
        if not args.input_columns:
            columns = ['B', 'C', 'D',]
        else: 
            columns = args.input_columns
    elif args.sheet == "Archive":
        sheet = "2.Описание"
        if not args.input_columns:
            columns = ['C', 'E', 'G', 'H', 'J']
        else:
            columns = args.input_columns
    else:
        sheet = input("Please enter name of the sheet to translate: ")


    if not gloss_path.exists():
        raise FileNotFoundError(
            f"Glossary file not found at {gloss_path}. Please provide a valid path."
        )
    filetype = str(args.sheet)  # fixed as per original function
    print(gloss_path)
    print(columns)
    #
    Begin.modrun(
        self=None,
        data_path=args.data_path,
        ilanguage=args.input_language,
        olanguage=args.output_language,
        filetype=filetype,
        glossfile=args.glossfile,
        sheet=sheet,
        input_column=columns,
        #output_column=args.output_columns,
        start_row=args.start_row,
        output_dir=args.output_directory
    )


init()
