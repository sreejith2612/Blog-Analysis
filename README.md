# Blog Analysis
# Web Scraping and Text Analysis

## Instructions:

1. **Extract the zip file**: 
   - Locate the zip file containing all the required files (e.g., `program_files.zip`).
   - Right-click on the zip file.
   - Choose "Extract All" or similar option based on your operating system.
   - Select a destination folder where you want to extract the files and click "Extract".

2. **Navigate to the extracted folder**: 
   - Open the folder where you extracted the files. You should see all the necessary files including `main.py`, `OutputDataStructure.xlsx`, and two text files and requirements.txt.

3. **Check Python installation**: 
   - Ensure that Python is installed on your system. You can do this by opening a command prompt (Windows) or terminal (macOS/Linux) and typing:
     ```
     python --version
     ```
   - If Python is not installed, download and install it from the official Python website: [Python Downloads](https://www.python.org/downloads/)

4. **Install required libraries**:
   - Open a command prompt or terminal.
   - Navigate to the directory where you extracted the files using the `cd` command.
   - Install the required libraries using pip and the `requirements.txt` file. Run the following command:
     ```
     pip install -r requirements.txt
     ```
   
5. **Run the Python program**:
   - Open a command prompt (Windows) or terminal (macOS/Linux).
   - Navigate to the directory where you extracted the files using the `cd` command.
   - Run the Python program by typing:
     ```
     python main.py
     ```
   - Press Enter to execute the command.
   - The program should start running.

6. **Review the output**:
   - Once the program finishes execution, the output is updated in the OutputDataStructure.xlsx in the local directory

------------------------------------------------------------------------------------------------------------------------

## Improved Approach to the Solution:

### Introduction:
The task of web scraping the Blackcoffer website involved getting the text data from the url using requests library.
This was done by examining the html class name for the website, it was as straightforward task as it was consistent across most urls.

### Text Cleaning:
Once the articles were retrieved, next step was to tidy up the text. 
NLTK library was used to remove all the stopwords and punctuations as mentioned in the objective.

### Analysis Methods:
To facilitate analysis, several methods were implemented:

1. `calculate_positive_score`
2. `calculate_negative_score`
3. `calculate_polarity_score`
4. `calculate_subjectivity_score`
5. `calculate_complex_words`
6. `calculate_average_sentence_length`
7. `calculate_avg_word_per_sentence`
8. `calculate_fog_index`
9. `count_syllables`
10. `calculate_average_syllable_count`
11. `count_personal_pronouns`
12. `calculate_average_word_length`

These methods analyze the text and provide the necessary variables. The resulting variables are automatically stored within the workbook.

### Challenges Faced:

1. **Error Handling**: Some URLs throw a 404 error. In such cases, all variables are set to 0.
2. **Different Templates**: Certain URLs utilize distinct HTML templates for articles, necessitating generalized code adaptable to all templates.
3. **Workbook Automation**: Streamlining the process of updating variables within the workbook.
4. **Modular Approach**: Adhering to a modular approach to prevent code clutter and enhance readability.

------------------------------------------------------------------------------------------------------------------------

