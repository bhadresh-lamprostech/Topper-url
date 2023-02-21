from selenium import webdriver
from bs4 import BeautifulSoup
from openpyxl import load_workbook, Workbook

# Load the input workbook and worksheet
input_wb = load_workbook("input.xlsx")
input_ws = input_wb.active

# Create a new workbook and worksheet
output_wb = Workbook()
output_ws = output_wb.active

# Write column headings to output worksheet
output_ws.append(["Question", "Choice 1", "Choice 2", "Choice 3", "Choice 4"])

for row in input_ws.iter_rows(min_row=2, values_only=True):
    url = row[0]

    driver = webdriver.Chrome()
    driver.get(url)

    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')

    # Find the question body and extract the text content
    question = soup.find("div", class_="TopprQuestion_questionBody__OBlTc").get_text()

    # Find all four choices and extract the text content
    choices = soup.find_all("div", class_="Option_content__2ZU_b")
    choice_texts = [choice.get_text() for choice in choices]

    # Rearrange choices based on their order on the page
    ordered_choices = [None] * 4
    for i, choice in enumerate(choices):
        choice_text = choice.get_text()
        index = choice_texts.index(choice_text)
        ordered_choices[index] = choice_text

    # Write the question and choices to the output worksheet
    output_ws.append([question] + choice_texts)


    driver.quit()

# Save the output workbook to a file
output_wb.save("output.xlsx")
