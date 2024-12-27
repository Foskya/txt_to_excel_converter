import pandas as pd
from pandas import ExcelWriter
import time
import os
import sys
import random


MAXIMUM_WORDS_PER_CELL = 20 # edit this to set the amout of words you want contained in a cell

def input_file_selection():
    def txt_files_in_directory_list() :
        txt_files_list = []
        current_directory = os.getcwd()
        with os.scandir(current_directory) as directory_content_list :
            for element in directory_content_list :
                if element.name[-4:] == ".txt" or element.name[-4:] == ".TXT" :
                    txt_files_list.append(element.name)
        return txt_files_list

    def file_selection(txt_files_list) :
        while len(txt_files_list) > 1 :
            print("In the present directory there are multiple .txt files: \n")
            for txt_nr, txt_name in enumerate(txt_files_list) :
                print(f"n: {txt_nr}\t{txt_name}")
            try:
                chosen_txt = input("\nenter the number related to the file you want to select: ")
                txt_files_list = [txt_files_list[int(chosen_txt)]]
            except Exception as _exeption:
                print(f"\n{chosen_txt} is not a valid result, retry\n")
                time.sleep(1)
        return txt_files_list[0]
    
    txt_files_list = txt_files_in_directory_list()
    return file_selection(txt_files_list)

def file_creation(input_file):
    def split_line_into_batches(book_line):
        book_line_list = book_line.split()
        for i in range(0, len(book_line_list), MAXIMUM_WORDS_PER_CELL):
            batch_of_words_per_cell = ' '.join(book_line_list[i:i+MAXIMUM_WORDS_PER_CELL])
            yield batch_of_words_per_cell


    def dataframe_book_column_writer(input_file):
        rows = []
        cell_written_count = 0
        with open(input_file, 'r', encoding='utf-8') as book_file:
            for line in book_file:
                for cell_content in split_line_into_batches(line):
                    rows.append({"BOOK_COLUMN": cell_content})
                    cell_written_count += 1
                    sys.stdout.flush()
                    sys.stdout.write(f"\rcell number: {cell_written_count}")
        return pd.DataFrame(rows)

    def dataframe_filler_columns_writer(df):
        random_result_touple = ("a", "b", "c")
        df['FILLER_1'] = df.apply(lambda x: random.choices(random_result_touple, weights=(5, 65, 30), k=1)[0], axis=1)
        df['FILLER_2'] = df['FILLER_1'].apply(lambda x: str(random.randint(1, 10)) if x == 'a' else "not_a")
        #df['FILLER3'] = df['FILLER_2'] * 2

    def file_excel_writer(excel_name, df):
        with ExcelWriter(f"{excel_name}.xlsx", engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='book_sheet', index=False)

    df = dataframe_book_column_writer(input_file)
    dataframe_filler_columns_writer(df)
    file_excel_writer(input_file[:-4], df)

def main():
    input_file = input_file_selection()
    os.system('cls')
    print(input_file[:-4])
    file_creation(input_file)

if __name__ == "__main__":
  main()
