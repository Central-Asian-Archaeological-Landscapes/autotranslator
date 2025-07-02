"""
Created on 19 Jun 2022

@author: turch
"""

import argparse
from set_parameters import Begin
import pathlib


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
        required=True,
        help="Path to folder/file with data to translate",
    )
    parser.add_argument(
        "-il",
        "--input_language",
        type=str,
        choices=["ru", "en"],
        default="ru",
        help="Language of the input data (options: 'ru', 'en')",
    )
    parser.add_argument(
        "-ol",
        "--output_language",
        type=str,
        default="en",
        help="Language to translate the data into",
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
        choices=["Monuments", "Archive", "Other"],
        default="Archive",
        required=True,
        help="Sheet name to translate",
    )
    parser.add_argument(
        "-ic",
        "--input_columns",
        nargs="+",
        type=str,
        required=True,
        help="The columns to be translated",
    )
    parser.add_argument(
        "-oc",
        "--output_columns",
        nargs="+",
        type=str,
        required=True,
        help="The columns to insert translated data into. Please type in corresponding order to input columns",
    )
    parser.add_argument(
        "-r",
        "--start_row",
        type=int,
        default=5,
        help="Row to start translation from (excluding column names, default: 5)",
    )
    args = parser.parse_args()

    if args.sheet == "Monuments":
        sheet = "Data Sheet"
    elif args.sheet == "Archive":
        sheet = "2.Описание"
    else:
        sheet = input("Please enter name of the sheet to translate: ")

    if not gloss_path.exists():
        raise FileNotFoundError(
            f"Glossary file not found at {gloss_path}. Please provide a valid path."
        )
    filetype = "Excel"  # fixed as per original function
    print(gloss_path)
    #
    Begin.modrun(
        self=None,
        data_path=args.data_path,
        ilanguage=args.ilanguage,
        olanguage=args.olanguage,
        filetype=filetype,
        glossfile=args.glossfile,
        sheet=sheet,
        input_column=args.input,
        output_column=args.output,
        start_row=args.start_row,
    )


init()
