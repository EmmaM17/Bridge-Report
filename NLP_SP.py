#Import all the modules needed
import pandas as pd

#Text modification/cleaning
from nltk.corpus import stopwords
from nltk.corpus import wordnet as wn
from nltk.stem import WordNetLemmatizer
from nltk.tokenize import word_tokenize
from nltk import pos_tag
import nltk
from collections import defaultdict

#ML/SVM
from sklearn.model_selection import train_test_split
from sklearn.svm import SVC
from sklearn.preprocessing import LabelEncoder
from sklearn.metrics import confusion_matrix, ConfusionMatrixDisplay
from sklearn.metrics import accuracy_score
from sklearn.feature_extraction.text import TfidfVectorizer

#Saving the model
import pickle

#Visualising accuracy (confusion matrix)
import matplotlib.pyplot as plt

#connect to sharepoint and get the excel file
from office365.sharepoint.client_context import ClientContext;
from office365.runtime.auth.client_credential import ClientCredential
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.files.file import File
import io 
import os
# Set up link to sharepoint
site_url = 'https://comapnyName.sharepoint.com/teams/sharepointName/'
read_folder_URL ='/teams/sharepointName/Shared Documents/Structural Report Output/'

client_id = #### enter client id here ####
client_secret = #### enter client secret here ####'
client_credentials = ClientCredential(client_id, client_secret)
ctx = ClientContext(site_url).with_credentials(client_credentials)
read_folder = ctx.web.get_folder_by_server_relative_url(read_folder_URL)

# Retrieve the files within the folder
files = read_folder.files
ctx.load(files)
ctx.execute_query()

for f in files:
    if f.properties["Name"] == 'Bridge_Report_Data.xlsx':
        file = f
        break

response = file.open_binary(ctx, file.serverRelativeUrl)
bytes_file_obj = io.BytesIO()
bytes_file_obj.write(response.content)
bytes_file_obj.seek(0)
page = 0


####Real data we want to categorize####
report_df= pd.read_excel(bytes_file_obj,sheet_name='History')

masterlist_df= pd.read_excel(bytes_file_obj,sheet_name='Masterlist')
information_df= pd.read_excel(bytes_file_obj,sheet_name='Information')
sectiona_df= pd.read_excel(bytes_file_obj,sheet_name='Section A')


####Training Data####
url = 'Training Data 2.xlsx' #Add your own specific file path
training_df = pd.read_excel(url) #Can also filter if data is part of a wider dataset

####Create Machine Learning####
def my_NLP(report_df,training_df):
    #There can't be any blank data, so use .fillna to choose what category that stuff will go into.
    #My preference is usually a 'null' category with 'no comment' as the description.
    data = list(report_df['Description'].fillna('null'))
    test_data= list(training_df['Description'].fillna('null'))
    test_labels = list(training_df['Category'].fillna('no comment'))

    # Make data lowercase and labels capitalized for consistency
    data = [entry.lower() for entry in data]
    test_data=[entry.lower() for entry in test_data]
    test_labels = [entry.capitalize() for entry in test_labels]

    #Tokenize Data
    data = [word_tokenize(entry) for entry in data] #Splits up (tokenizes) the descriptions into individual words
    test_data= [word_tokenize(entry) for entry in test_data] 

    #lemmatize data
    filtered_test_data=lemmatization(test_data)
    filtered_data=lemmatization(data)
    

    #Split into training/test data
    
    X_train, X_test, y_train, y_test = train_test_split(filtered_test_data, test_labels,test_size=0.2) 
    #Note here -- test_size changes the proportion of data that is held back to be used for testing, so in
    #the above case 20% of the data is held behind.

    #Word Vectorisation - we need to convert the words into a form the machine can interpret (numbers essentially). 
    Tfidf_vect = TfidfVectorizer(max_features=6000) #initiates vectorizer - max_features can be tweaked
    Tfidf_test_vect = TfidfVectorizer(max_features=6000)
    Tfidf_vect.fit(filtered_data) #fits vectorizer to dataset
    Tfidf_test_vect.fit(filtered_test_data)

    #Transform text data
    X_train_tfidf = Tfidf_vect.transform(X_train)
    X_test_tfidf = Tfidf_vect.transform(X_test)
    X_test_real_tfidf = Tfidf_vect.transform(filtered_data)

    
    #Set up and train the machine
    SVM = SVC(kernel='poly',degree=4) #several parameters can be changed here depending on needs. Watch video series @ link above or google for better idea
    SVM.fit(X_train_tfidf,y_train) #trains the machine using the training data

    #Generate predictions using the test data
    predictions = SVM.predict(X_test_tfidf)
  
    #Calculate the accuracy for test data
    #print("SVM Accuracy Score for test data -> ",accuracy_score(predictions, y_test)*100) 

    #now we want to predict categories for our real data
    predict_real= SVM.predict(X_test_real_tfidf)
  
    # Add predicted categories to report_df
    report_df['Category'] = predict_real

    # Print report_df with added Category column
    print(report_df)
  

####lemmatized words####
def lemmatization(data):
    #Lemmatize words (get into most basic form so machine can read) and remove stopwords (common words that likely offer no value)
    word_lemmatized = WordNetLemmatizer() #Initiate lemmatizer
    filtered_data = [] #Will be final dataset
    for index,entry in enumerate(data):
        final_words = []
        for word, tag in pos_tag(entry):
            if word not in stopwords.words('english') and word.isalpha(): #.isalpha checks to make sure no characters other than a-z
                word_final = word_lemmatized.lemmatize(word) #makes words concise (makes plurals singular, running->run, better->good)
                final_words.append(word_final)         
        filtered_data.append(' '.join(final_words))

    return filtered_data

####Generate Confusion Matrix####
def confusion_matrix_plot(y_test,predictions,SVM):
    #Generate confusion matrix 
    cm = confusion_matrix(y_test,predictions,labels=SVM.classes_)
    disp = ConfusionMatrixDisplay(confusion_matrix=cm,display_labels=SVM.classes_,)
    disp.plot(xticks_rotation=30, cmap='Oranges')
    plt.show()

my_NLP(report_df,training_df)


####Export category ####

def export(report_df):

    #Export example_df
    buffer = io.BytesIO()      # Create a buffer object

    # Write the dataframe to the buffer
    writer = pd.ExcelWriter(buffer, engine='xlsxwriter')
    masterlist_df.to_excel(writer, sheet_name='Masterlist', index=False)
    information_df.to_excel(writer, sheet_name='Information', index=False)
    sectiona_df.to_excel(writer, sheet_name='Section A', index=False)
    report_df.to_excel(writer, sheet_name='History', index=False)
    # Save and close the Excel writer
    writer._save()

    # Retrieve the file content
    file_content = buffer.getvalue()
    buffer.seek(0)
    file_content = buffer.read()
    #Create output path
    path = "sharepoint name/Outputs/Bridge_Report_Data.xlsx"
    target_folder = read_folder
    name = os.path.basename(path)
    #Here is where we actually upload the file - using the "execute_query()" command again from ctx
    target_file = target_folder.upload_file(name, file_content).execute_query()
    print("File has been uploaded to url: {0}".format(target_file.serverRelativeUrl))

export(report_df)
