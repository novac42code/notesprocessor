import openai
import pandas as pd
from openai.api_resources.completion import Completion

# Load your API key from an environment variable or secret management service
openai.api_key = 'TO BE Replaced'

# Define the file paths
input_file_path = r'C:\Users\filepath'
output_file_path = r'C:\Users\filepath'
category_file_path = r'C:\Users\filepath'

# Load the excel file
df = pd.read_excel(input_file_path)
df_category = pd.read_excel(category_file_path)

# Define the instruction
instruction = "Now you act as . Please help categorize the given text into the most probable topical area based on the given categories, giving no more than top3 categories. Please just give a list of categories without any other introductory words. If none of the given categories apply, please create one you consider as most appropriate."

# Now we can create a dictionary where each category maps to a list of its subcategories
categories = {category: [] for category in df_category['Category'].unique()}

# Define a function to get category using OpenAI API
def get_category(title, description):
    # Get the first level category
    prompt = instruction + "this issue's title is:" + title + "this issue's description is:" + description + ". What category does this issue belong to: " + ', '.join(categories.keys()) + "?"
    response = Completion.create(engine="text-davinci-003", prompt=prompt, max_tokens=1000)
    category = response['choices'][0]['text'].strip()

    return category

# Apply the function to the combination of 'Title' and 'Description' columns
df['Category'] = df.apply(lambda row: get_category(row['Title'], row['Description']), axis=1)

# Write the dataframe to a new excel file
df.to_excel(output_file_path, index=False)
