#------------------------------------------------------------------------------------------#
#importing the dependencies
from serpapi import GoogleSearch
import pandas as pd
import time
#from pandas import json_normalize
#import json

#------------------------------------------------------------------------------------------#
#Declaring the variables
api_key = "" #enter api key here

#------------------------------------------------------------------------------------------#
#Creating a Excel to store the result

# Create empty dataframes
df_citation_count = pd.DataFrame()
df_complete_data = pd.DataFrame()

# Create an Excel writer object and write the empty dataframes to it

# Define columns for each sheet
citation_columns = [
    "Author Name", "citations_all", "citations_since_2018", "h_index_all",
    "h_index_since_2018", "i10_index_all", "i10_index_since_2018"
]

complete_data_columns = [
    "Author Name", "Affiliations", "Email", "Article Title", "Article Link",
    "Citation_ID", "Article Authors", "Publication", "Year", "Cited By", "Cites ID"
]

# Create empty dataframes with the specified columns
df_citation_count = pd.DataFrame(columns=citation_columns)
df_complete_data = pd.DataFrame(columns=complete_data_columns)

# Create an Excel writer object and write the empty dataframes to it
with pd.ExcelWriter('Faculty_Data_Demo.xlsx') as writer:
    df_citation_count.to_excel(writer, sheet_name='Citation_Data', index=False)
    df_complete_data.to_excel(writer, sheet_name='Complete_Data', index=False)

print("Empty Excel file created successfully!")
 
#----------------------------------------------------------------------------------------------
#Reading the data from the excel
df = pd.read_excel('Faculty_Google_Scholar_Mapping_Demo.xlsx', sheet_name='Name_ID_Mapping')
#----------------------------------------------------------------------------------------------



file_path = "Faculty_Data_Demo.xlsx"

#----------------------Duplication deletion function----------------------------------------------



def remove_duplicates_from_excel(file_path):

    df_duplication_removal_complete_data = pd.read_excel(file_path, sheet_name="Complete_Data")
    df_duplication_removal_citation_data = pd.read_excel(file_path, sheet_name="Citation_Data")

    df_duplication_removal_complete_data = df_duplication_removal_complete_data.drop_duplicates()
    df_duplication_removal_citation_data = df_duplication_removal_citation_data.drop_duplicates()

    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
        # Check if 'Complete_Data' sheet exists and remove it
        if 'Complete_Data' in writer.book.sheetnames:
            writer.book.remove(writer.book['Complete_Data'])
        df_duplication_removal_complete_data.to_excel(writer, sheet_name="Complete_Data", index=False)

    # Write the deduplicated data back to the same sheet
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
        # Check if 'Complete_Data' sheet exists and remove it
        if 'Complete_Data' in writer.book.sheetnames:
            writer.book.remove(writer.book['Citation_Data'])
        df_duplication_removal_citation_data.to_excel(writer, sheet_name="Citation_Data", index=False)


#-----------------------------------------------------------------------------------------------------
# Iterate through each row in the DataFrame
for index, row in df.iterrows():
    faculty_name = row['Name']
    
    scholar_id = row['ID']
    
    has_over_100_articles = row['No of Articles more than 100']
    
    
    # Logic for blank Google Scholar ID
    if not scholar_id or pd.isnull(scholar_id):
        # Create a new row with only the Author Name and append to the dataframes
        new_row_citation = {"Author Name": [faculty_name]}  # Wrapped faculty_name in a list
        df_citation_count = pd.concat([df_citation_count, pd.DataFrame(new_row_citation)], ignore_index=True)  # Using concat instead of append
    
        new_row_complete = {"Author Name": [faculty_name]}  # Wrapped faculty_name in a list
        df_complete_data = pd.concat([df_complete_data, pd.DataFrame(new_row_complete)], ignore_index=True)  # Using concat instead of append
        
        # After appending the new data, read the existing data from the Excel file.
        existing_data_citation = pd.read_excel('Faculty_Data_Demo.xlsx', sheet_name='Citation_Data')
        existing_data_complete = pd.read_excel('Faculty_Data_Demo.xlsx', sheet_name='Complete_Data')
        
        # Update the dataframes by appending new data.
        df_citation_count = pd.concat([existing_data_citation, df_citation_count], ignore_index=True)
        df_complete_data = pd.concat([existing_data_complete, df_complete_data], ignore_index=True)
        
        # Now write the updated dataframes back to the Excel, overwriting the whole file.
        with pd.ExcelWriter('Faculty_Data_Demo.xlsx', engine='openpyxl') as writer:
            df_citation_count.to_excel(writer, sheet_name='Citation_Data', index=False)
            df_complete_data.to_excel(writer, sheet_name='Complete_Data', index=False)
        time.sleep(10)
        remove_duplicates_from_excel(file_path)
        time.sleep(10)
        continue

        
    print("Starting the api hitting logic")
    # API logic -> Execute this when ID is not empty
    params = {
      "engine": "google_scholar_author",
      "hl": "en",
      "author_id": scholar_id,
      "api_key": api_key,
      #"start": "100", #keep this line only for article count more than 100 because this is the limit of the api
      "num": "200"
    }
    print("API logic complete")

    #searching for the params entered
    search = GoogleSearch(params)
    results = search.get_dict()
    #Result recieved sucessfully and the result is validated using the print 


    # Extracting the author details
    author = results['author']
    name = author['name']
    affiliations = author['affiliations']
    email = author['email']

    # Extract article details
    articles_list = results['articles']
    article_data = []

    # Parse each article
    for article in articles_list:
        title = article['title']
        link = article['link']
        citation_id = article['citation_id']
        authors = article['authors']
        publication = article['publication'] if 'publication' in article else None
        year = article['year']
        cited_by_value = article['cited_by']['value'] if 'cited_by' in article else None
        cited_by_cites = article['cited_by']['cites_id'] if ('cited_by' in article and 'cites_id' in article['cited_by']) else None

        #Appending the data
        article_data.append([name, affiliations, email, title, link, citation_id,  authors, publication, year, cited_by_value, cited_by_cites])

    #Adding this data to the excel file
    df_new_articles = pd.DataFrame(article_data, columns=complete_data_columns)
    # Concatenate the new article data with the existing df_complete_data
    df_complete_data = pd.concat([df_complete_data, df_new_articles], ignore_index=True)
    # First, read the existing data from the Excel file.
    existing_data_citation = pd.read_excel('Faculty_Data_Demo.xlsx', sheet_name='Citation_Data')
    existing_data_complete = pd.read_excel('Faculty_Data_Demo.xlsx', sheet_name='Complete_Data')

    # Update the dataframes by appending new data.
    df_citation_count = pd.concat([existing_data_citation, df_citation_count], ignore_index=True)
    df_complete_data = pd.concat([existing_data_complete, df_complete_data], ignore_index=True)

    # Now write the updated dataframes back to the Excel, overwriting the whole file.
    with pd.ExcelWriter('Faculty_Data_Demo.xlsx', engine='openpyxl') as writer:  # Notice, no mode='a' here.
        df_citation_count.to_excel(writer, sheet_name='Citation_Data', index=False)
        df_complete_data.to_excel(writer, sheet_name='Complete_Data', index=False)

    # Extract article details
    cited_by_cleanup_list = results['cited_by']['table']

    # Create empty dictionary to store the data
    data = {
        'citations_all': [],
        'citations_since_2018': [],
        'h_index_all': [],
        'h_index_since_2018': [],
        'i10_index_all': [],
        'i10_index_since_2018': []
    }

    # Extract data from the list
    for item in cited_by_cleanup_list:
        if 'citations' in item:
            data['citations_all'].append(item['citations']['all'])
            data['citations_since_2018'].append(item['citations']['since_2018'])
        else:
            data['citations_all'].append(None)
            data['citations_since_2018'].append(None)

        if 'h_index' in item:
            data['h_index_all'].append(item['h_index']['all'])
            data['h_index_since_2018'].append(item['h_index']['since_2018'])
        else:
            data['h_index_all'].append(None)
            data['h_index_since_2018'].append(None)

        if 'i10_index' in item:
            data['i10_index_all'].append(item['i10_index']['all'])
            data['i10_index_since_2018'].append(item['i10_index']['since_2018'])
        else:
            data['i10_index_all'].append(None)
            data['i10_index_since_2018'].append(None)
    
    #handling the none values in the data
    for key, values in data.items():
        data[key] = [value for value in values if value is not None]
    print("Data")
    print(data)

    # Extract graph citation
    graph = results['cited_by']['graph']
    graph_data = []
    # Convert graph into desired format
    graph_format = {str(item['year']): [item['citations']] for item in graph}
    print("Graph format")
    print(graph_format)
    
    
    # Extract values from data and graph_format
    data_values = [item for sublist in data.values() for item in sublist]
    graph_values = [item for sublist in graph_format.values() for item in sublist]
    print("Data values")
    print(data_values)
    print("Graph values")
    print(graph_values)

    # Concatenate the values
    merged_values = [name] + data_values
    print("Printing Merged values")
    print(merged_values)

    # Convert to a pandas DataFrame
    df_citation_data = pd.DataFrame(merged_values, columns=['Value'])
    #print(df_citation_data)

    # Convert df_citation_data from wide format to long format
    df_citation_data_transposed = df_citation_data.T
    #print(df_citation_data_transposed)

    #print(df_citation_data.head(5))

    #Code to add data to the sheet
    # Code to add data to the sheet
    # First, read the existing data from both sheets
    # Read the existing data from the Excel file.
    

    # Define column names for the transposed data to match with the existing Excel data.
    citation_columns = [
    "Author Name", "citations_all", "citations_since_2018", "h_index_all",
    "h_index_since_2018", "i10_index_all", "i10_index_since_2018"
    ]

    df_citation_data_transposed.columns = citation_columns  # Assigning the column names

    # First, read the existing data from the sheet 'Citation_Data'
    existing_citation_data = pd.read_excel('Faculty_Data_Demo.xlsx', sheet_name='Citation_Data')
    existing_complete_data = pd.read_excel('Faculty_Data_Demo.xlsx', sheet_name='Complete_Data')

    # Append the transposed df_citation_data to the existing data
    updated_citation_data = pd.concat([existing_citation_data, df_citation_data_transposed], ignore_index=True)

    # Write both sheets back to the Excel file
    with pd.ExcelWriter('Faculty_Data_Demo.xlsx', engine='openpyxl') as writer:
        existing_complete_data.to_excel(writer, sheet_name='Complete_Data', index=False)
        updated_citation_data.to_excel(writer, sheet_name='Citation_Data', index=False)
    time.sleep(10)
    remove_duplicates_from_excel(file_path)
    time.sleep(10)

    # Logic for "No of Articles more than 100" is 1
    if has_over_100_articles >= 1:
        print("Starting the over 100 logic")
        # Replace "pass" with your actual logic
        params = {
            "engine": "google_scholar_author",
            "hl": "en",
            "author_id": scholar_id,
            "api_key": api_key,
            "start": "100", #keep this line only for article count more than 100 because this is the limit of the api
            "num": "200"
        }
        print("API hit completed")

        search = GoogleSearch(params)
        results = search.get_dict()

        author = results['author']
        name = author['name']
        affiliations = author['affiliations']
        email = author['email']

        articles_list = results['articles']
        article_data = []

        for article in articles_list:
            title = article['title']
            link = article['link']
            citation_id = article['citation_id']
            authors = article['authors']
            publication = article['publication'] if 'publication' in article else None
            year = article['year']
            cited_by_value = article['cited_by']['value'] if 'cited_by' in article else None
            cited_by_cites = article['cited_by']['cites_id'] if ('cited_by' in article and 'cites_id' in article['cited_by']) else None

            #Appending the data
            article_data.append([name, affiliations, email, title, link, citation_id,  authors, publication, year, cited_by_value, cited_by_cites])

        #Adding this data to the excel file
        df_new_articles = pd.DataFrame(article_data, columns=complete_data_columns)
        # Concatenate the new article data with the existing df_complete_data
        df_complete_data = pd.concat([df_complete_data, df_new_articles], ignore_index=True)
        # First, read the existing data from the Excel file.
        existing_data_citation = pd.read_excel('Faculty_Data_Demo.xlsx', sheet_name='Citation_Data')
        existing_data_complete = pd.read_excel('Faculty_Data_Demo.xlsx', sheet_name='Complete_Data')

        # Update the dataframes by appending new data.
        df_citation_count = pd.concat([existing_data_citation, df_citation_count], ignore_index=True)
        df_complete_data = pd.concat([existing_data_complete, df_complete_data], ignore_index=True)

        # Now write the updated dataframes back to the Excel, overwriting the whole file.
        with pd.ExcelWriter('Faculty_Data_Demo.xlsx', engine='openpyxl') as writer:  # Notice, no mode='a' here.
            df_citation_count.to_excel(writer, sheet_name='Citation_Data', index=False)
            df_complete_data.to_excel(writer, sheet_name='Complete_Data', index=False)

        

    # Logic for "No of Articles more than 100" is 1
    if has_over_100_articles >= 2:
        print("Starting the over 200 logic")
        # Replace "pass" with your actual logic
        params = {
            "engine": "google_scholar_author",
            "hl": "en",
            "author_id": scholar_id,
            "api_key": api_key,
            "start": "200", #keep this line only for article count more than 100 because this is the limit of the api
            "num": "200"
        }
        print("API hit completed")

        search = GoogleSearch(params)
        results = search.get_dict()

        author = results['author']
        name = author['name']
        affiliations = author['affiliations']
        email = author['email']

        articles_list = results['articles']
        article_data = []

        for article in articles_list:
            title = article['title']
            link = article['link']
            citation_id = article['citation_id']
            authors = article['authors']
            publication = article['publication'] if 'publication' in article else None
            year = article['year']
            cited_by_value = article['cited_by']['value'] if 'cited_by' in article else None
            cited_by_cites = article['cited_by']['cites_id'] if ('cited_by' in article and 'cites_id' in article['cited_by']) else None

            #Appending the data
            article_data.append([name, affiliations, email, title, link, citation_id,  authors, publication, year, cited_by_value, cited_by_cites])

        #Adding this data to the excel file
        df_new_articles = pd.DataFrame(article_data, columns=complete_data_columns)
        # Concatenate the new article data with the existing df_complete_data
        df_complete_data = pd.concat([df_complete_data, df_new_articles], ignore_index=True)
        # First, read the existing data from the Excel file.
        existing_data_citation = pd.read_excel('Faculty_Data_Demo.xlsx', sheet_name='Citation_Data')
        existing_data_complete = pd.read_excel('Faculty_Data_Demo.xlsx', sheet_name='Complete_Data')

        # Update the dataframes by appending new data.
        df_citation_count = pd.concat([existing_data_citation, df_citation_count], ignore_index=True)
        df_complete_data = pd.concat([existing_data_complete, df_complete_data], ignore_index=True)

        # Now write the updated dataframes back to the Excel, overwriting the whole file.
        with pd.ExcelWriter('Faculty_Data_Demo.xlsx', engine='openpyxl') as writer:  # Notice, no mode='a' here.
            df_citation_count.to_excel(writer, sheet_name='Citation_Data', index=False)
            df_complete_data.to_excel(writer, sheet_name='Complete_Data', index=False)
        
    # Logic for "No of Articles more than 100" is 1
    if has_over_100_articles >= 3:
        print("Starting the over 300 logic")
        # Replace "pass" with your actual logic
        params = {
            "engine": "google_scholar_author",
            "hl": "en",
            "author_id": scholar_id,
            "api_key": api_key,
            "start": "300", #keep this line only for article count more than 100 because this is the limit of the api
            "num": "200"
        }
        print("API hit completed")

        search = GoogleSearch(params)
        results = search.get_dict()

        author = results['author']
        name = author['name']
        affiliations = author['affiliations']
        email = author['email']

        articles_list = results['articles']
        article_data = []

        for article in articles_list:
            title = article['title']
            link = article['link']
            citation_id = article['citation_id']
            authors = article['authors']
            publication = article['publication'] if 'publication' in article else None
            year = article['year']
            cited_by_value = article['cited_by']['value'] if 'cited_by' in article else None
            cited_by_cites = article['cited_by']['cites_id'] if ('cited_by' in article and 'cites_id' in article['cited_by']) else None

            #Appending the data
            article_data.append([name, affiliations, email, title, link, citation_id,  authors, publication, year, cited_by_value, cited_by_cites])

        #Adding this data to the excel file
        df_new_articles = pd.DataFrame(article_data, columns=complete_data_columns)
        # Concatenate the new article data with the existing df_complete_data
        df_complete_data = pd.concat([df_complete_data, df_new_articles], ignore_index=True)
        # First, read the existing data from the Excel file.
        existing_data_citation = pd.read_excel('Faculty_Data_Demo.xlsx', sheet_name='Citation_Data')
        existing_data_complete = pd.read_excel('Faculty_Data_Demo.xlsx', sheet_name='Complete_Data')

        # Update the dataframes by appending new data.
        df_citation_count = pd.concat([existing_data_citation, df_citation_count], ignore_index=True)
        df_complete_data = pd.concat([existing_data_complete, df_complete_data], ignore_index=True)

        # Now write the updated dataframes back to the Excel, overwriting the whole file.
        with pd.ExcelWriter('Faculty_Data_Demo.xlsx', engine='openpyxl') as writer:  # Notice, no mode='a' here.
            df_citation_count.to_excel(writer, sheet_name='Citation_Data', index=False)
            df_complete_data.to_excel(writer, sheet_name='Complete_Data', index=False)

    # Logic for "No of Articles more than 100" is 1
    if has_over_100_articles >= 4:
        print("Starting the over 400 logic")
        # Replace "pass" with your actual logic
        params = {
            "engine": "google_scholar_author",
            "hl": "en",
            "author_id": scholar_id,
            "api_key": api_key,
            "start": "400", #keep this line only for article count more than 100 because this is the limit of the api
            "num": "200"
        }
        print("API hit completed")

        search = GoogleSearch(params)
        results = search.get_dict()

        author = results['author']
        name = author['name']
        affiliations = author['affiliations']
        email = author['email']

        articles_list = results['articles']
        article_data = []

        for article in articles_list:
            title = article['title']
            link = article['link']
            citation_id = article['citation_id']
            authors = article['authors']
            publication = article['publication'] if 'publication' in article else None
            year = article['year']
            cited_by_value = article['cited_by']['value'] if 'cited_by' in article else None
            cited_by_cites = article['cited_by']['cites_id'] if ('cited_by' in article and 'cites_id' in article['cited_by']) else None

            #Appending the data
            article_data.append([name, affiliations, email, title, link, citation_id,  authors, publication, year, cited_by_value, cited_by_cites])

        #Adding this data to the excel file
        df_new_articles = pd.DataFrame(article_data, columns=complete_data_columns)
        # Concatenate the new article data with the existing df_complete_data
        df_complete_data = pd.concat([df_complete_data, df_new_articles], ignore_index=True)
        # First, read the existing data from the Excel file.
        existing_data_citation = pd.read_excel('Faculty_Data_Demo.xlsx', sheet_name='Citation_Data')
        existing_data_complete = pd.read_excel('Faculty_Data_Demo.xlsx', sheet_name='Complete_Data')

        # Update the dataframes by appending new data.
        df_citation_count = pd.concat([existing_data_citation, df_citation_count], ignore_index=True)
        df_complete_data = pd.concat([existing_data_complete, df_complete_data], ignore_index=True)

        # Now write the updated dataframes back to the Excel, overwriting the whole file.
        with pd.ExcelWriter('Faculty_Data_Demo.xlsx', engine='openpyxl') as writer:  # Notice, no mode='a' here.
            df_citation_count.to_excel(writer, sheet_name='Citation_Data', index=False)
            df_complete_data.to_excel(writer, sheet_name='Complete_Data', index=False)
        
time.sleep(10)
remove_duplicates_from_excel(file_path)

'''
df_duplication_removal_complete_data = pd.read_excel(file_path, sheet_name="Complete_Data")
df_duplication_removal_citation_data = pd.read_excel(file_path, sheet_name="Citation_Data")

# Remove duplicate rows
df_duplication_removal_complete_data = df_duplication_removal_complete_data.drop_duplicates()
df_duplication_removal_citation_data = df_duplication_removal_citation_data.drop_duplicates()

# Write the deduplicated data back to the same sheet
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
    # Check if 'Complete_Data' sheet exists and remove it
    if 'Complete_Data' in writer.book.sheetnames:
        writer.book.remove(writer.book['Complete_Data'])
    df_duplication_removal_complete_data.to_excel(writer, sheet_name="Complete_Data", index=False)

# Write the deduplicated data back to the same sheet
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
    # Check if 'Complete_Data' sheet exists and remove it
    if 'Complete_Data' in writer.book.sheetnames:
        writer.book.remove(writer.book['Citation_Data'])
    df_duplication_removal_citation_data.to_excel(writer, sheet_name="Citation_Data", index=False)
'''
