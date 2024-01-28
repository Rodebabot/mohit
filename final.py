import fitz  # pip install PyMuPDF
import os
import pandas as pd
from prettytable import PrettyTable

def find_words_in_pdf(pdf_path, target_words):
    try:
        # Open the PDF file
        pdf_document = fitz.open(pdf_path)

        # Initialize list to store results for each word
        results = []

        # Iterate through each word
        for target_word in target_words:
            # Initialize list to store page numbers and frequency for the current word
            found_pages = []
            total_frequency = 0

            # Iterate through each page
            for page_number in range(pdf_document.page_count):
                page = pdf_document[page_number]

                # Get the page text
                page_text = page.get_text()

                # Check if the word is present on the page
                occurrences = page_text.lower().count(target_word.lower())
                if occurrences > 0:
                    total_frequency += occurrences

                    # Record the next 20 characters after the first instance
                    index = page_text.lower().find(target_word.lower())
                    following_text = page_text[index + len(target_word):index + len(target_word) + 20]

                    found_pages.append({
                        "Page Number": page_number + 1,
                        "Frequency on Page": occurrences,
                        "Following Text of First Instance": following_text
                    })

            # Append results for the current word to the overall results list
            results.append({"Word": target_word, "Total Frequency": total_frequency, "Found Pages": found_pages})

        # Close the PDF file
        pdf_document.close()

        return results
    except Exception as e:
        print(f"An error occurred: {e}")
        return None

def extract_images_from_pdf(pdf_path, output_folder):
    try:
        # Open the PDF file
        pdf_document = fitz.open(pdf_path)

        # Iterate through each page
        for page_number in range(pdf_document.page_count):
            page = pdf_document[page_number]

            # Get the images on the page
            images = page.get_images(full=True)

            # Check if there are images on the page
            if images:
                # Create a subfolder for each page
                page_folder = os.path.join(output_folder, f"page_{page_number + 1}")
                os.makedirs(page_folder, exist_ok=True)

                # Iterate through each image on the page
                for img_index, img_info in enumerate(images):
                    image_index = img_info[0]
                    base_image = pdf_document.extract_image(image_index)
                    image_bytes = base_image["image"]

                    # Determine the image file format
                    image_format = base_image["ext"]

                    # Save the image
                    image_filename = f"{page_folder}/page_{page_number + 1}_image_{img_index + 1}.{image_format}"
                    with open(image_filename, "wb") as image_file:
                        image_file.write(image_bytes)

                print(f"Images found on page {page_number + 1}. Image saved.")

        # Close the PDF file
        pdf_document.close()
    except Exception as e:
        print(f"An error occurred: {e}")

def print_results(results):
    for result in results:
        if result["Found Pages"]:
            table = PrettyTable()
            table.field_names = ["Page Number", "Frequency on Page", "Following Text of First Instance"]

            for page_info in result["Found Pages"]:
                table.add_row([page_info['Page Number'], page_info['Frequency on Page'], page_info['Following Text of First Instance']])

            print(f"\nThe word '{result['Word']}' was found with a total frequency of {result['Total Frequency']} on the following page(s):")
            print(table)
        else:
            print(f"\nThe word '{result['Word']}' was not found in the PDF.")

def main():
    try:
        pdf_path = input('Enter PDF path: ')
        output_folder = input('Enter Destination folder path: ')

        # Validate PDF file path
        if not os.path.isfile(pdf_path) or not pdf_path.lower().endswith('.pdf'):
            raise ValueError("Invalid PDF file path")

        # Validate output folder path
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        elif not os.path.isdir(output_folder):
            raise ValueError("Invalid output folder path")

        # Prompt user to enter the path to a CSV or Excel file
        file_path = input('Enter CSV/Excel file path containing names in the first column: ')

        # Validate CSV/Excel file path
        if not os.path.isfile(file_path) or not (file_path.lower().endswith(('.csv', '.xlsx', '.xls'))):
            raise ValueError("Invalid CSV/Excel file path")

        # Load data from CSV/Excel using pandas
        if file_path.lower().endswith('.csv'):
            df = pd.read_csv(file_path)
        elif file_path.lower().endswith(('.xlsx', '.xls')):
            df = pd.read_excel(file_path)

        # Extract names from the first column
        target_words = df.iloc[:, 0].astype(str).tolist()

        # Search for each word in the PDF
        results = find_words_in_pdf(pdf_path, target_words)

        if results is not None:
            print_results(results)

        # Extract images from the PDF
        print('\n\nIMAGES - ')
        extract_images_from_pdf(pdf_path, output_folder)

    except ValueError as ve:
        print(f"Validation error: {ve}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    main()
