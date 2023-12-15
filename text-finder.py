import os
import pandas as pd

def find_text_pandas(folder_path, text):
    found_files = []
    counter = 0

    # Cleaning the search text
    cleaned_text = text.lower()

    for filename in os.listdir(folder_path):
        if filename.endswith(".xlsx"):
            file_path = os.path.join(folder_path, filename)

            try:
                df = pd.read_excel(file_path, sheet_name=None)
                for sheet_name, sheet in df.items():
                  for col in sheet.columns:
                      for index, cell in enumerate(sheet[col]):
                          cell_value = str(cell).lower().strip()  # Add .strip() to remove leading/trailing whitespaces
                          if cell_value == cleaned_text:
                              found_files.append(file_path)
                              print(f"Text ditemukan di file {file_path}, sheet {sheet_name}")
                              return
            except Exception as e:
                print(f"Error opening file {file_path}: {e}")

            counter += 1

    print("Total file yang diproses:", counter)
    if found_files:
        print("File yang ditemukan:", found_files)
    else:
        print("Text tidak ditemukan")

if __name__ == "__main__":
    folder_path = "/content/drive/My Drive/your_drive/"
    text = "NRR" // text you looking for
    find_text_pandas(folder_path, text)
