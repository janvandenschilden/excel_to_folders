import os
import sys
import pandas as pd
import argparse
import shutil

def main():
    """
    Main function

    :return: None
    """
    # Parse arguments and assign to variables
    excel_file, sheet_name, output_folder,document_folder = parse_args()
    
    # Remove output folder if it exists and make new one
    ## it is easier to remove and recreate the folder than to check if every document is in the right place or not
    remove_folder(output_folder)
    create_folder(output_folder)

    # Read excel file
    df = read_excel(excel_file, sheet_name)

    # Check whether documents exists in the document folder
    check_whether_documents_exists(document_folder, df)

    # Create folders from excel file
    create_folders_from_excel(df, output_folder)

    # Copy documents to folders
    copy_documents_to_folders(df, document_folder, output_folder)

def parse_args():
    """
    Parse arguments

    :return: args
    """
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--excel_file", 
        help="excel file to read",
        default="example.xlsx",
        type=str,
    )
    parser.add_argument(
        "--sheet_name",
        help="sheet name to read",
        default="Sheet1",
        type=str,
    )
    parser.add_argument(
        "--output_folder",
        help="output folder to write",
        default="output",
        type=str,
    )
    parser.add_argument(
        "--document_folder",
        help="folder with documents",
        default="documents",
        type=str,
    )
    args = parser.parse_args()
    return args.excel_file, args.sheet_name, args.output_folder, args.document_folder
    args = parser.parse_args()
    return args.excel_file, args.sheet_name, args.output_folder, args.document_folder

def remove_folder(folder):
    """
    Remove folder if it exists

    :param folder: folder to remove
    :return: None
    """
    if os.path.exists(folder):
        shutil.rmtree(folder)

def create_folder(folder):
    """
    Create folder

    :param folder: folder to create
    :return: None
    """
    os.makedirs(folder, exist_ok=True)

def read_excel(excel_file, sheet_name):
    """
    Read excel file

    :param excel_file: excel file to read
    :param sheet_name: sheet name to read
    :return: df
    """
    # df is a pandas dataframe, the name is by convention
    df = pd.read_excel(
        excel_file, 
        sheet_name=sheet_name,  # allows the excel file to have multiple sheets
        index_col=0             # index_col is the column to use as the index, makes it easier to iterate over the rows of pandas dataframe
    ).fillna(False)             # fillna(False) fills all empty cells with False
    return df

def check_whether_document_exists(document_folder, document_name):
    """
    Check whether document exists in the document folder

    :param document_folder: folder with documents
    :param document_name: document name
    :return: None
    """
    document_path = os.path.join(document_folder, document_name)
    if not os.path.exists(document_path):
        assert False, f"Document {document_name} does not exist in the {document_folder}/ folder"
        sys.exit(1)

def check_whether_documents_exists(document_folder, df):
    """
    Check whether documents exists in the document folder

    :param document_folder: folder with documents
    :param df: pandas dataframe
    :return: None
    """
    for document_name in df.columns:
        check_whether_document_exists(document_folder, document_name)

def create_folders_from_excel(df, output_folder):
    """
    Create folders inside the output folder

    :param df: pandas dataframe
    :param output_folder: output folder to write subfolders in
    :return: None
    """
    for index, row in df.iterrows():
        folder_path = os.path.join(output_folder, index)
        create_folder(folder_path)

def copy_document_to_folder(document_folder, document_name, output_folder, folder_name):
    """
    Copy document to folder

    :param document_folder: folder with documents
    :param document_name: document name
    :param output_folder: output folder to write subfolders in
    :param folder_name: folder name
    :return: None
    """
    document_path = os.path.join(document_folder, document_name)
    folder_path = os.path.join(output_folder, folder_name)
    shutil.copy(document_path, folder_path)

def copy_documents_to_folders(df, document_folder, output_folder):
    """
    Copy documents to folders

    :param df: pandas dataframe
    :param document_folder: folder with documents
    :param output_folder: output folder to write subfolders in
    :return: None
    """
    for index, row in df.iterrows():
        for document_name in df.columns:
            if row[document_name] != False: # only empty cells are False, so this checks if the cell is empty
                copy_document_to_folder(document_folder, document_name, output_folder, index)


if __name__ == "__main__":
    main()