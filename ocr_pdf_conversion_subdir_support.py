#!/usr/bin/env python
# coding: utf-8

# In[1]:


#NOTE copy of ocr_pdf_conversion_minimal.py but now supports ocr on each subdir as a municipality


# In[2]:


import os
import fitz
import pandas as pd
import docx2txt
import time
import nltk
from nltk.corpus import brown
from nltk.corpus import stopwords

import collections
from collections import defaultdict
from collections import OrderedDict
import difflib

from pprint import pprint
import io
import re
import sys

start_time = time.perf_counter()

# sys.stdout = open('tail.log', 'w')

if not os.path.exists('C:\inetpub\OurFTPFolder\OCR\start'):
    exit()
else:
    os.remove('C:\inetpub\OurFTPFolder\OCR\start')
    


nltk.download('stopwords')
nltk.download('brown')

# In[3]:


"""

Variables to change: 
path_to_base_dir 
"""

#subdir example
path_to_base_dir = r"C:\inetpub\OurFTPFolder\OCR"
completed_dir = r"C:\inetpub\OurFTPFolder\Completed\\"
completed_meta_dir = r"C:\inetpub\OurFTPFolder\Completed-Meta\\"
archive_dir = r"C:\inetpub\OurFTPFolder\Archive"


# In[4]:``


def use_ocr_text(row):
    
    #if row.page_text == Nan then  row.page_text == row.page_text returns False so not makes it true so the program knows no text in that row.    
    if row.page_text == "nan" or  row.page_text == "" or row.page_text.isspace():
        return True#use OCR text because embedded doesn't have any text or text worth using( whitespace only)
    else: 
        #embedded text can be used
        #NOTE: this is where one could check if more OCR text rather than whatever is in embedded.
        return False

    


# In[5]:


"""
Used to determine if rotating the PDF results in more correctly spelled english words, if so then 
the PDF needs to be rotated before applying OCR because the page format is setup differently.

An attempt to maximize words captured with OCR.
TODO: Potentially remove stop words from brown.words to not weight them heavily
"""
word_list = brown.words()
word_set_original = set(word_list)



stopwords =set(stopwords.words('english'))

"""
Note Brown corpus considers single character english tokens as words which might throw off word count if rotated valid word count is close to not rotated valid word count: B,o,b, S, every single character it seems.
Removing below.

TODO: Get a better word checking model/method. 
TODO: brown.words() still contains 2 char words which are
old words but are more likely to be OCR error than actually occur in the document. 
For example:  en, ye, pa

TODO: Could get all 2 character words from brown and remove the old ones to clean up the accepted OCR words. 
"""

# only want to count 'a' and 'i' as valid single char word. Also removing ] [ _ etc
#single_alphabetical_chars_to_not_count_as_valid_word = {'_','[',']','b', 'c', 'd', 'e', 'f', 'g', 'h', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z'}#Otherwise Brown considers a-z{1} to be a valid "word"

#without 'a' and 'i' because they over power actual words. These two can easily be OCR errors
single_alphabetical_chars_to_not_count_as_valid_word = {'a','i','_','[',']','b', 'c', 'd', 'e', 'f', 'g', 'h', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z'}#Otherwise Brown considers a-z{1} to be a valid "word"


#removing certain words from brown words
word_set=word_set_original - single_alphabetical_chars_to_not_count_as_valid_word
print("word_set length with stopwords:{}".format(len(word_set)))

#removing stopwords from brown words
#word_set_wo_stopwords=word_set_original - stopwords
word_set=(word_set_original - stopwords) - single_alphabetical_chars_to_not_count_as_valid_word
print("word_set length without stopwords:{}".format(len(word_set)))


# In[6]:



"""
TODO: Might need to add below to env variables for setup to work. 
TESSDATA_PREFIX  C:\Program Files\Tesseract-OCR\tessdata 

Note: All variables that need to be changed to point to input & output directories are below in this cell.

Main variables to change: 
#dir containing municipalities as subfolders 
path_to_base_dir
ocr_output_files_dir_name
ocr_output_file_metadata_dir_name

path_to_website_documents
path_to_ocr_output_dir
minimial_file_output

"""



website_files_to_ocr_dir_name = "website_files_to_ocr"
#Base dir to search from for files
#Note folders would have to be made already or add folder creation into code
ocr_output_files_dir_name = "..\..\Completed"
ocr_output_file_metadata_dir_name= "..\..\Completed-Meta"
#NOTE: This would be the dir containing sub-folders/subdirectories each as one municipality containing all the scrapped documents and webpages .
#path_to_base_dir = r"C:\Users\data_bindu\Documents\welcomehomes Automation of municipalities rules"




output_excel_file_name = "ocred_website_files.xlsx"#results of OCR on PDF pages. Not sure if this is needed because more complex page joins( embedded and OCR) are done to maximize text in PDF 
output_excel_file_metadata_name = "ocred_website_files_metadata.xlsx"

output_excel_original_text_file_name = "original_text_from_website_files.xlsx"#text which was embedded in PDF( not a scan)

output_excel_embedded_text_per_page_and_file_name = "embedded_text_per_page_from_website_files.xlsx"#text which was embedded in PDF( not a scan)

#Contains the orientation_metric_ocred_page_text for all pages ( use_ocr_x == True or use_ocr_x ==False) not just the text pages which don't have embedded text i.e. use_ocr_x == True as seen in page_orientation_stats_extra.xlsx .
output_excel_ocr_text_per_page_and_file_name = "ocr_text_per_page_from_website_files.xlsx"

output_excel_merged_text_per_page_and_file_name= "merged_text_per_page_from_website_files.xlsx"#subset of pages needing ocr

#NOTE: Don't use as input to model because excel cell has character limit which truncates text. Instead load in .txt file
output_excel_merged_text_final_per_page_and_file_name = "pdf_merged_text_per_page_from_website_files DONT USE TRUNCATES.xlsx"#full df with all rows potentially join all pages into one df

#page links found in the pdf. Context data/feature for model. Could lookup if page is captured on website page scrapper.
output_excel_file_page_links = "website_files_page_links.xlsx"


#Debug structure for valid words to ensure OCR page orientation is operating correctly.
output_excel_valid_words_rotated_ocr_text_per_page_and_file_name = "valid_words_rotated_ocr_text_per_page_from_website_files.xlsx"#text which was embedded in PDF( not a scan)
output_excel_valid_words_ocr_text_per_page_and_file_name = "valid_words_ocr_text_per_page_from_website_files.xlsx"#text which was embedded in PDF( not a scan)


output_excel_file_rotated_ocr_text_name = "rotated_ocr_text.xlsx"
output_excel_file_not_rotated_ocr_text_name = "not_rotated_ocr_text.xlsx"
#Stats on which text was used for a given page: rotated page or not, and word counts for those pages.
output_excel_file_page_orientation_stats = "page_orientation_stats.xlsx"
output_excel_file_page_orientation_stats_extra = "page_orientation_stats_extra.xlsx"




output_excel_file_per_page_merged_text_file_name = "merged_text_final_df.xlsx"



#will not perform ocr on PDFs with more than this number of pages.
#Will alter methods if a count is applied to a dir because this will filter some pdfs out of dir if longer than page limit
#Not really needed because ocr_page_early_stop_limit could at least OCR this many pages and just use that to represent the PDF file rather than not applying OCR entirely 
ocr_page_count_limit = 4000

#OCR stopping by only apply OCR to first n pages 
ocr_page_early_stop_limit = 400

skipped_pdf_file_name = "skipped_pdfs_too_long.xlsx"


"""
NOTE: Keep minimial_file_output = False
because some files here are needed for model. 
Page links or PDF metadata for example, and the other files are needed for review of OCR conversion
"""

minimial_file_output = False#debug mode, more files generated to see the steps/computation and why certain pages are being used, metadata from PDF, etc
#minimial_file_output = True#output minimial files needed which is the ocr dir with plain text for each PDF file.

"""
Keep old_output = True, and could remove code/variables associated with this if need to improve speed/lower compute. 
"""
#Can likely clean up code which relates to this field/flag
#old_output = False#Include files which have been replaced by better methods/views of data.
old_output = True#Don't include files which have been replaced by better methodsmethods/views of data.


#Likely keep False unless to verify the joining process between files. If true then more files will be made which include redundant info
#if true then see as seperate files prior to join, if false then don't write out files because the data is in another file. 
see_not_joined_files = False
#see_not_joined_files = True


# In[7]:


"""
Attempt with os.listdir after gathering all ocr_folders 
with path_to_website_files_to_ocr = r"{}\{}".format(path_to_municipality,website_files_to_ocr_dir_name)
Uses the fact that the folder will be named  website_files_to_ocr_dir_name = "website_files_to_ocr"
"""

ocr_folder_paths = []
municipality_and_ocr_files_paths_dict = {}#key = path to base of a single municipality, value = path to municipality/website_files_to_ocr
"""
Creating output folders for OCR results and metadata
"""
for municipality in os.listdir(path_to_base_dir):
    #Base of folder. Used to create output folders.
    path_to_municipality = r"{}\{}".format(path_to_base_dir,municipality)
    print("path_to_municipality={}\n".format(path_to_municipality))
    #Folder with all OCRED files one txt file for each input PDF. 
    path_to_ocr_output_dir = r"{}\\{}".format(completed_dir,municipality)
    print("path_to_ocr_output_dir={}\n".format(path_to_ocr_output_dir))

    
    path_to_website_files_to_ocr = r"{}\{}".format(path_to_municipality,website_files_to_ocr_dir_name)
    municipality_and_ocr_files_paths_dict[path_to_municipality] = [path_to_website_files_to_ocr,municipality]
    print("path_to_website_files_to_ocr={}\n".format(path_to_website_files_to_ocr))

    #Folder to store metadata collected when converting PDFs into plain text.
    path_to_ocr_metadata_output_dir = r"{}\\{}".format(completed_meta_dir,municipality)
    print("path_to_ocr_metadata_output_dir={}\n".format(path_to_ocr_metadata_output_dir))

    try:
        os.mkdir(path_to_ocr_output_dir)
        os.mkdir(path_to_ocr_metadata_output_dir)

    except FileExistsError:
        print("Dir already made")





for base_municipality_path,[ocr_folder_path,municipality] in municipality_and_ocr_files_paths_dict.items():
    print(ocr_folder_path)
    #dir_name = strip right most from ocr_folder_path
    #norm_ocr_folder_path = os.path.normpath(ocr_folder_path)
    #dir_name = os.path.basename(norm_ocr_folder_path)
    path_to_municipality = r"{}".format(base_municipality_path)
    print("path_to_municipality={}\n".format(path_to_municipality))
    path_to_ocr_output_dir = r"{}\\{}".format(completed_dir,municipality)
    print("path_to_ocr_output_dir={}\n".format(path_to_ocr_output_dir))
    path_to_ocr_metadata_output_dir = r"{}\\{}".format(completed_meta_dir,municipality)
    print("path_to_ocr_metadata_output_dir={}\n".format(path_to_ocr_metadata_output_dir))
    



    excel_output_path = r"{}\{}".format(path_to_ocr_metadata_output_dir,output_excel_file_name)#ocred website files
    print(excel_output_path)
    #per page and file below
    excel_output_embedded_text_per_page_and_file_path = r"{}\{}".format(path_to_ocr_metadata_output_dir,output_excel_embedded_text_per_page_and_file_name)#original embedded text in pdfs
    excel_output_ocr_text_per_page_and_file_path = r"{}\{}".format(path_to_ocr_metadata_output_dir,output_excel_ocr_text_per_page_and_file_name)#original embedded text in pdfs


    #location to valid words on each pages for rotational OCR vs normal orientation OCR.
    excel_output_valid_words_rotated_ocr_text_per_page_and_file_path = r"{}\{}".format(path_to_ocr_metadata_output_dir,output_excel_valid_words_rotated_ocr_text_per_page_and_file_name)#original embedded text in pdfs
    excel_output_valid_words_ocr_text_per_page_and_file_path = r"{}\{}".format(path_to_ocr_metadata_output_dir,output_excel_valid_words_ocr_text_per_page_and_file_name)#original embedded text in pdfs




    excel_output_original_text_path = r"{}\{}".format(path_to_ocr_metadata_output_dir,output_excel_original_text_file_name)

    """
    TODO: Use metadata and links as features to re-ranking/ ML model.
    excel_metadata_output_path & excel_path_links_output_path
    """
    #metadata for the pdf files
    excel_metadata_output_path = r"{}\{}".format(path_to_ocr_metadata_output_dir,output_excel_file_metadata_name)
    #Links found in PDF likely Important 
    excel_path_links_output_path = r"{}\{}".format(path_to_ocr_metadata_output_dir,output_excel_file_page_links)

    excel_output_merged_text_path = r"{}\{}".format(path_to_ocr_metadata_output_dir,output_excel_merged_text_per_page_and_file_name)


    excel_output_skipped_pdfs_path = r"{}\{}".format(path_to_ocr_metadata_output_dir,skipped_pdf_file_name)

    excel_output_page_orientation_stats_path = r"{}\{}".format(path_to_ocr_metadata_output_dir,output_excel_file_page_orientation_stats)
    excel_output_page_orientation_stats_extra_path= r"{}\{}".format(path_to_ocr_metadata_output_dir,output_excel_file_page_orientation_stats_extra)

    #ocr results for the page being rotated or not
    excel_output_rotated_ocr_text_path = r"{}\{}".format(path_to_ocr_metadata_output_dir,output_excel_file_rotated_ocr_text_name)
    excel_output_not_rotated_ocr_text_path = r"{}\{}".format(path_to_ocr_metadata_output_dir,output_excel_file_not_rotated_ocr_text_name)


    """
    NOTE: Fix
    excel_output_merged_text_per_page_before_joining_to_file_path
    Same output_excel_merged_text_final_per_page_and_file_name
    """
    #results of each page of text per document prior to merging into one file per PDF
    excel_output_merged_text_per_page_before_joining_to_file_path = r"{}\{}".format(path_to_ocr_metadata_output_dir,output_excel_merged_text_final_per_page_and_file_name)

    output_excel_merged_text_final_name = "not_used_file.xlsx"
    excel_output_merged_final_text_path = r"{}\{}".format(path_to_ocr_metadata_output_dir,output_excel_merged_text_final_name)
    
    




    """
    NOTE USE FOR OCR ON PDFS 
    """


    """
    OrderedDict() used below to link up the page text differences from embedded and OCRED text. 
    file_original_text_dict = {}#store text embedded in the PDF file already. 

    file_dict = {}
    file_metadata_dict = {}
    """

    #OrderedDict not used or needed in current version: could be normal dict
    file_original_text_dict = OrderedDict()#store text embedded in the PDF file already. 
    file_dict = OrderedDict()
    file_metadata_dict = OrderedDict()
    file_page_links_dict = defaultdict(list)#default dict with list as key. page : [links on page] 

    ocr_text_per_file_and_page = {}# key = file_name , value = dict(key = page number, value = text on page)
    embedded_text_per_file_and_page = {}# key = file_name , value = dict(key = page number, value = text on page)

    ocr_text_orientation_stats_per_file_and_page = {}# key = file_name , value = dict(key = page number, value = tuple( rotated words count, not rotated words count ))

    #Used to save the docx and txt files into the save folder as OCRED converted PDFs
    docx_file_dict ={}
    txt_file_dict ={}


    if not minimial_file_output:
        rotated_ocr_text_per_file_and_page = {}# key = file_name , value = dict(key = page number, value = text on page)
        not_rotated_ocr_text_per_file_and_page = {}# key = file_name , value = dict(key = page number, value = text on page)
        rotated_ocr_text_valid_words_per_file_and_page = defaultdict(list)# key = file_name , value = dict(key = page number, value = list[valid words on page])
        ocr_text_valid_words_per_file_and_page = defaultdict(list)# key = file_name , value = dict(key = page number, value = list[valid words on page])



    skipped_pdf_dict = {}#key = file, value = page count
    number = len(os.listdir(ocr_folder_path))
    if number == 0:
        continue
    for file in os.listdir(ocr_folder_path):
        #All the documents collected from a single webpage/municipality
       # 
       # if sys.argv[1] != file:
        #    continue
        print(file)
        if ".pdf" in file: 
            #pdf_url_path = r"{}\{}".format(path_to_website_documents,file)
            pdf_url_path = r"{}\{}".format(ocr_folder_path,file)
            print(pdf_url_path)
            #print(pdf_url_path)
            doc = fitz.open(pdf_url_path)

            if not minimial_file_output:
                #debugging differences between word return with rotating page vs not
                rotated_ocr_text_per_page = {}
                not_rotated_ocr_text_per_page = {}
                rotated_ocr_text_valid_words_per_page = defaultdict(list)# key = file_name , value = dict(key = page number, value = list[valid words on page])
                ocr_text_valid_words_per_page = defaultdict(list)# key = file_name , value = dict(key = page number, value = list[valid words on page])




            # Page count check will bypass pdf if larger than ocr_page_count_limit
            if doc.page_count > ocr_page_count_limit: 
                skipped_pdf_dict[file]=doc.page_count
                continue


            #Save document metadata and relate to OCR errors.
            file_metadata_dict[file] = doc.metadata # nested dict: with  (file_name : dict(PDF arg/feature name: value))

            ocr_text_per_page_dict = {}#key == page number, value = text on page
            embedded_text_per_page_dict = {}#key == page number, value = text on page
            ocr_text_orientation_stats_per_page_dict = {}#key == page number, value = tuple( rotated words count, not rotated words count ))


            page_number = 1

            file_text = ""#orced text
            file_original_text =""#embedded text



            page_img_ocr_save_path = r"{}\{}.txt".format(path_to_ocr_output_dir,file)
            print(page_img_ocr_save_path)
            page_agg = []#each page of OCR text
            page_agg_original_text =[]#each page of embedded text




            #print(os.environ["TESSDATA_PREFIX"])
            for page in doc:#pages of pdf




                if page_number > ocr_page_early_stop_limit:
                    #file closed outside loop. 
                    break#break into the outter loop here thus stopping OCR on the current file.




                """
                TODO: Could include positional encoding of words with TextPage.extractDict()
                #https://pymupdf.readthedocs.io/en/latest/page.html#Page.get_text
                """

                #Saving all links found on a page.
                if len(page.get_links()) > 0:
                    file_page_links_dict[file].append(page.get_links())


                """
                Check if page is not a scan(image) and instead already has text recognized in the PDF( i.e. if text is in document then don’t need to apply OCR).
                If already text and OCR is applied then there is a chance of OCR error, especially if the text font is weird. 
                Can use this also to record OCR errors if text is in document. These errors could be made into rules to transform text if a OCRED word is incorrect.

                .get_text() 
                .get_textpage()
                """

                #
                #original_text_page=page.get_textpage()#see if PDF is not a scan and has the text already.
                #embedded_text=original_text_page.extractTEXT()
                embedded_text=page.get_text()
                page_agg_original_text.append(embedded_text)
                embedded_text_per_page_dict[page_number] = embedded_text


                """
                No page rotation originally. 
                Check word counts for rotation and not rotated.
                """



                text_page=page.get_textpage_ocr(flags=3, language="eng", dpi=72, full=True)# full=True
                #text_page=page.get_textpage_ocr(flags=3, language="eng", dpi=72, full=False)#partial ocr, worse results can't handle color background
                ocr_text_from_page=text_page.extractText()
                list_of_terms=ocr_text_from_page.split()#default split is on space

                #NOTE: Could use to tokenize IR inverted index
                #replace anything not a alphabet character with nothing. Using to tokenize
                clean_list_of_terms = [re.sub("[^a-zA-z]+","",term) for term in list_of_terms]#these aren't valid words.
                clean_list_of_terms = [re.sub("[\[\]-_]+","",term) for term in clean_list_of_terms]#removing [, ], _ , -
                clean_list_of_terms = [term.lower() for term in clean_list_of_terms]

                #print("tokenized text:clean_list_of_terms. NOT valid words just cleaned tokens\n")
                #print(set(clean_list_of_terms))


                unique_valid_words = set()


                valid_word_count = 0
                #for term in list_of_terms:
                for term in clean_list_of_terms:
                    if term in word_set:
                        #print("valid word found in not rotated page:{}".format(term))
                        ocr_text_valid_words_per_page[page_number].append(term)
                        valid_word_count+= 1
                        unique_valid_words.add(term)


                if not minimial_file_output:
                    not_rotated_ocr_text_per_page[page_number] = ocr_text_from_page

                unique_valid_word_count = len(unique_valid_words)
                """
                Doing rotation word check on each page. 
                Double compute but handles rotated PDFs
                https://pymupdf.readthedocs.io/en/latest/page.html#Page.set_rotation

                page.set_rotation(90) is inplace

                rotated page below  with page.set_rotation(90)
                """
                page.set_rotation(90)

                #See if dpi of the file is returned from metadata or elsewhere and use to pass in here for best OCR settings.
                #DPI not in document details.
                #if file_metadata_dict[file][dpi] number then
                #dpi = file_metadata_dict[file][dpi]
                rotated_text_page= page.get_textpage_ocr(flags=3, language="eng", dpi=72, full=True)
                rotated_ocr_text_from_page=rotated_text_page.extractText()
                list_of_rotated_terms=rotated_ocr_text_from_page.split()#default is on space

                #replace anything not a alphabet character with nothing. Using to tokenize
                clean_list_of_rotated_terms = [re.sub("[^a-zA-z]+","",term) for term in list_of_rotated_terms]
                clean_list_of_rotated_terms = [re.sub("[\[\]-_]+","",term) for term in clean_list_of_rotated_terms]
                clean_list_of_rotated_terms = [term.lower() for term in clean_list_of_rotated_terms]
                #print("tokenized text:clean_list_of_rotated_terms. NOT valid words just cleaned tokens\n")
                #print(set(clean_list_of_rotated_terms))
                rotated_valid_word_count = 0


                #Unique words could still bias the data because the OCR junk could have many short unique words.
                unique_rotated_valid_words = set()
                #for term in list_of_rotated_terms:
                for term in clean_list_of_rotated_terms:
                    if term in word_set:
                        #print("valid word found in rotated page:{}".format(term))
                        rotated_valid_word_count+= 1
                        rotated_ocr_text_valid_words_per_page [page_number].append(term)#valid word
                        unique_rotated_valid_words.add(term)


                unique_rotated_valid_word_count = len(unique_rotated_valid_words)


                if not minimial_file_output:
                    #Storing both rotated and not rotated OCR results per file per page as a feature
                    rotated_ocr_text_per_page[page_number] = rotated_ocr_text_from_page


                #print("rotated word count:{}\t unrotated word count:{}".format(rotated_valid_word_count,valid_word_count))

                #Factor in unique words( set count) and word character length as weighted sum because OCR junk of 1-2 char words are overwhelming actual words when page is not rotated
                rotated_word_percent_bias = .60
                word_length_weight = 1#NOTE two of these variables make sure they are the same starting value or else it could bias results

                #subtracted away
                rotated_word_bias = int(rotated_valid_word_count *rotated_word_percent_bias)#to not overcount words if rotated since most documents aren't rotated. 

                #Weighted sum based on character count

                rotated_word_char_length_dict = defaultdict(list)#key = word length, value = [words with key as word length] 
                for valid_rotated_word in rotated_ocr_text_valid_words_per_page[page_number]:
                    key = len(valid_rotated_word)
                    rotated_word_char_length_dict[key].append(valid_rotated_word)


                rotated_word_char_length_weighted_value = 0

                ordered_keys_rotated_word_char_length_dict = collections.OrderedDict(sorted(rotated_word_char_length_dict.items()))
                for rotated_word_length, word_list in ordered_keys_rotated_word_char_length_dict.items():
                #for rotated_word_length, word_list in rotated_word_char_length_dict.items():
                    #print("rotated_word_char_length_dict:{}\n".format(rotated_word_length))
                    word_length_weight +=.3
                    #word_length_weight +=1 #same weight for rotated vs not, but getting too many false positives so lowering rotated weight to hopefully reduce FP
                    rotated_word_char_length_weighted_value+= (rotated_word_length * word_length_weight)* len(word_list)


                #not rotated words 
                word_char_length_dict = defaultdict(list)#key = word length, value = [words with key as word length] 
                for valid_word in ocr_text_valid_words_per_page[page_number]:
                    key = len(valid_word)
                    word_char_length_dict[key].append(valid_word)


                word_char_length_weighted_value = 0
                word_length_weight = 1#Needed to reset word_length_weight after rotated added to the variable above.

                ordered_keys_word_char_length_dict = collections.OrderedDict(sorted(word_char_length_dict.items()))
                for word_length, word_list in ordered_keys_word_char_length_dict.items():
                #for word_length, word_list in word_char_length_dict.items():
                    #print("word_char_length_dict:{}\n".format(word_length))

                    #NOTE word_length_weight assumes sorted  Keys == char length,  needs to be sorted in ascending order. Otherwise higher weights could be assigned to random length words. 
                    #The longer the word the higher the weight it receives because OCR error for longer words is less likely but still possible.
                    word_length_weight +=1#giving higher weight to longer words. Keys == char length,  needs to be sorted in ascending order
                    word_char_length_weighted_value+= (word_length * word_length_weight)* len(word_list)



                rotated_page_metric = rotated_valid_word_count - rotated_word_bias + unique_rotated_valid_word_count + rotated_word_char_length_weighted_value
                not_rotated_page_metric = valid_word_count + unique_valid_word_count + word_char_length_weighted_value





                #if (rotated_valid_word_count - rotated_word_bias) > valid_word_count:#too many false positives
                # if rotated_page_metric larger then most likely the PDF page is rotated( LHS to RHS) rather than read from top to bottom. 

                # if both scores are within 10 points of each other add both words for the given page.
                degree_of_closeness = 10

                #True== Use both rotated text and unrotated as the page to not lose text. False clearer winner of the correct page orientation
                within_degree_of_closeness = abs(rotated_page_metric - not_rotated_page_metric) < degree_of_closeness



                #ocr_text_orientation_stats_per_page_dict[page_number] = (rotated_valid_word_count,valid_word_count)
                ocr_text_orientation_stats_per_page_dict[page_number] = (rotated_valid_word_count,valid_word_count,unique_rotated_valid_word_count,unique_valid_word_count,rotated_word_char_length_weighted_value,word_char_length_weighted_value,rotated_page_metric,not_rotated_page_metric,within_degree_of_closeness)



                #TODO:Check if this throws off other metadata files?
                #Add both rotated and not rotated text info for a given page as it is too close to tell the correct page orientation.
                #if (rotated_page_metric - not_rotated_page_metric) < degree_of_closeness:
                if within_degree_of_closeness:
                    #print(rotated_ocr_text_from_page)
                    page_agg.append(rotated_ocr_text_from_page)
                    #page_agg.append("\n")#TODO: Check if only one new line needed for diff obj
                    """
                    NOTE: Potential place for low quality OCR by adding both text from two orientations.
                    """
                    ocr_text_per_page_dict[page_number] = rotated_ocr_text_from_page + ocr_text_from_page
                    page_number += 1



                #NOTE: Below likely includes false positives.
                #Page is considered rotated: adding rotated text
                elif rotated_page_metric > not_rotated_page_metric:
                #if rotated_page_metric > not_rotated_page_metric:
                    #print(rotated_ocr_text_from_page)
                    page_agg.append(rotated_ocr_text_from_page)
                    #page_agg.append("\n")#TODO: Check if only one new line needed for diff obj
                    ocr_text_per_page_dict[page_number] = rotated_ocr_text_from_page
                    page_number += 1



                else:
                    #more words not rotated according to orientation metric so use the unrotated PDF OCRED page text 
                    #print(ocr_text_from_page)
                    page_agg.append(ocr_text_from_page)
                    ocr_text_per_page_dict[page_number] = ocr_text_from_page
                    #page_agg.append("\n")
                    page_number += 1







            #Document/PDF file done, join text and save into dict

            """
            Document/PDF file done, join text and save into dict
            Note: file_text, file_dict not needed anymore.  
            Using embedded_text_per_file_and_page & ocr_text_orientation_stats_per_file_and_page 
            for more complex joins later on in program.
            """
            #file_text=file_text.replace(u"\ufffd", "*")#replace 
            file_text = " ".join(page_agg)#joining strings from each page all at once, at the end.
            file_dict[file] = file_text

            #Save each page so joins are page based: this way can support PDF with scan pages & embedded text pages
            ocr_text_per_file_and_page[file]= ocr_text_per_page_dict

            embedded_text_per_file_and_page[file] = embedded_text_per_page_dict


            ocr_text_orientation_stats_per_file_and_page[file]= ocr_text_orientation_stats_per_page_dict





            if not minimial_file_output:
                rotated_ocr_text_per_file_and_page[file] = rotated_ocr_text_per_page# key = file_name , value = dict(key = page number, value = text on page)
                not_rotated_ocr_text_per_file_and_page[file] = not_rotated_ocr_text_per_page# key = file_name , value = dict(key = page number, value = text on page)
                rotated_ocr_text_valid_words_per_file_and_page[file] = rotated_ocr_text_valid_words_per_page
                ocr_text_valid_words_per_file_and_page[file] = ocr_text_valid_words_per_page





            file_original_text = " ".join(page_agg_original_text)
            file_original_text_dict[file] = file_original_text
            doc.close()

        elif ".docx" in file:
            """
            TODO: Verify that metadata files and other files are being applied to .docx & .txt as well--where relevant.
            """

            """
            Not a scan, text already in format computer understands: i.e. OCR not needed. 
            Read in text and output to file.
            """
            #word_url_path = r"{}\{}".format(path_to_website_documents,file)
            word_url_path = r"{}\{}".format(ocr_folder_path,file)

            file_text=docx2txt.process(word_url_path)
            file_dict[file] = file_text#old use docx_file_dict[file]

            docx_file_dict[file] = file_text









        elif ".txt" in file:

            """
            Not a scan, text already in format computer understands: i.e. OCR not needed. 
            Read in text and output to file.
            """
            #path_to_doc = r"{}\{}".format(path_to_website_documents,file)
            path_to_doc = r"{}\{}".format(ocr_folder_path,file)
            
            #txt_file=open(path_to_doc,"r+")
            #txt_file=open(path_to_doc,"r+",encoding='utf-8', errors='ignore')
            txt_file=open(path_to_doc,"r+",encoding='utf-8', errors='surrogateescape')
            

            all_lines_of_txt_file=txt_file.read()#string for each file.
            file_dict[file] = all_lines_of_txt_file#old use txt_file_dict[file]
            txt_file.close()

            txt_file_dict[file] = all_lines_of_txt_file

        else:
            print("Not supported file type:{}".format(file))

        os.remove(r"{}\{}".format(ocr_folder_path,file))



    #printing OCR warning. 
    print(fitz.TOOLS.mupdf_warnings())
    #End of apply OCR to a single municipality

    """
    Saving results of applying OCR to folder of files.
    """


    """
    Saving metadata from PDF file.
    pdf_field == Title useful e.g. "WORKERS’ COMPENSATION REQUIREMENTS UNDER WCL SECTION  57"
    ('Excavation Application (1).pdf', 'title')	STREET OPENING PERMIT APPLICATION

     'subject' might be useful: doesn't seem to be filled often.


    ('clergy_exemption_form.pdf', 'subject')	Application for partial tax exemption for real property of members of the clergy
    ('clergy_exemption_form.pdf', 'keywords')	Application for partial tax exemption for real property of members of the clergy

    Could relate 'creator' & 'producer' to OCR results. 
    ('Employment Application - Lake Isle2020.pdf', 'creator')	Microsoft® Word 2013
    ('Employment Application - Lake Isle2020.pdf', 'producer')	Microsoft® Word 2013

    'keywords' useful when provided
    ('GrievanceFormRP524.pdf', 'keywords')	"Complaint, Real Property Assessment, Board of Assessment Review"

    ('mv6641 (1).pdf', 'title')	How To Apply For A Parking Permit Or License Plates For Persons With Severe Disabilities 
    ('mv6641 (1).pdf', 'author')	New York State Department of Motor Vehicles
    ('mv6641 (1).pdf', 'subject')	Application For A Parking Permit Or License Plates, For Persons With Severe Disabilities
    ('mv6641 (1).pdf', 'keywords')	How, To, Apply, For, A, Parking, Permit, Or, License, Plates, Persons, With, Severe, Disabilities, Application, New, York, State, Department, of, Motor, Vehicles

    ('rp425b_fill_in 2018.pdf', 'title')	Form RP-425-B:6/18:Application for Basic STAR Exemption for the 2019-2020 School Year:rp425b

    """


    metadata_df=pd.DataFrame.from_dict({(i,j): file_metadata_dict[i][j] 
                               for i in file_metadata_dict.keys() 
                               for j in file_metadata_dict[i].keys()},
                           orient='index', columns=['pdf_field_value'])

    #Don't need to save below: redundant info. Unpack tuple in index into columns
    #metadata_df['file_name'], metadata_df['pdf_field'] = zip(*metadata_df.index)

    metadata_df.index.name = "file name and PDF field"


    metadata_df.to_excel(excel_metadata_output_path)


    """

    Saving PDF page links. 

    converting to df then exporting

    https://pymupdf.readthedocs.io/en/latest/page.html#Page.get_text

    kind values below
    https://pymupdf.readthedocs.io/en/latest/vars.html#linkdest-kinds

    Saving page links found in the PDF. Check these links/use as a feature/context to the model.

    'uri': 'mailto:PDtraffic@eastchester.org' might be needed reference to contact?

    """

    #file_page_links_dict


    #format
    #page_links=pd.Series(file_page_links_dict).rename_axis('key_column').reset_index(name='value_column')
    page_links=pd.Series(file_page_links_dict).rename_axis('file_name').reset_index(name='link_info')

    #print(page_links)
    page_links_df=pd.DataFrame(page_links)

    page_links_df.to_excel(excel_path_links_output_path)



    """
    NOT needed more complex joins used in merged_text_final.
    Nevermind need to use file_dict because the other word and .txt files use this
    """

    #converting ocred file text into needed format: to pandas df then writing to excel file.
    #file_dict[file]
    #Conversion of the whole file and join rather than the individual  pages actually used.

    df = pd.DataFrame.from_dict(file_dict,orient="index",columns=["text"])
    df.index.name = "file_name"

    #Removing any formulas( cell starting with =) so excel file doesn't error.
    df['text'].replace(to_replace='^=', value=' =',regex=True, inplace=True)
    #if not old_output:

    if not old_output:
    #if not minimial_file_output:
        df.to_excel(excel_output_path)


    """
    Saving embedded text from the PDF files. 
    Use this text over the OCRED text. Can check if the text feature is Nan, if so then use OCR field.
    """


    #original text from file into needed format: to pandas df then writing to excel file.
    #file_dict[file]

    df = pd.DataFrame.from_dict(file_original_text_dict,orient="index",columns=["text"])
    df.index.name = "file_name"

    #Removing any formulas( cell starting with =) so excel file doesn't error.
    df['text'].replace(to_replace='^=', value=' =',regex=True, inplace=True)

    if not old_output:
    #if not minimial_file_output:
        df.to_excel(excel_output_original_text_path)



    """
    Checking results from OCR rotated text vs not. These are different from 'orientation_metric_ocred_page_text'. 

    'orientation_metric_ocred_page_text' determines which OCR results( rotated, not, or both because too close to tell) are to be used. 

    rotated_ocr_text_per_file_and_page[file] = rotated_ocr_text_per_page# key = file_name , value = dict(key = page number, value = text on page)
    not_rotated_ocr_text_per_file_and_page[file] = not_rotated_ocr_text_per_page# key = file_name , value = dict(key = page number, value = text on page)


    """

    see_not_joined_files = False

    if not minimial_file_output:
        rotated_ocr_text_per_file_and_page_df=pd.DataFrame.from_dict({(i,j): rotated_ocr_text_per_file_and_page[i][j] 
                                   for i in rotated_ocr_text_per_file_and_page.keys() 
                                   for j in rotated_ocr_text_per_file_and_page[i].keys()},
                               orient='index', columns=['rotated_ocr_page_text'])#rotated_ocr_text 'page_text'




        not_rotated_ocr_text_per_file_and_page_df=pd.DataFrame.from_dict({(i,j): not_rotated_ocr_text_per_file_and_page[i][j] 
                                   for i in not_rotated_ocr_text_per_file_and_page.keys() 
                                   for j in not_rotated_ocr_text_per_file_and_page[i].keys()},
                               orient='index', columns=['not_rotated_ocred_page_text'])




        rotated_ocr_text_per_file_and_page_df['rotated_ocr_page_text'].replace(to_replace='^=', value=' =',regex=True, inplace=True)

        rotated_ocr_text_per_file_and_page_df['file_name'],rotated_ocr_text_per_file_and_page_df['page_number'] = zip(*rotated_ocr_text_per_file_and_page_df.index)
        #Reset index to get a random index and then be able to do the merging.
        rotated_ocr_text_per_file_and_page_df.reset_index( drop=True, inplace=True)


        """

        """
        if see_not_joined_files:
            #NOTE: Could remove individual write out as data in page_orientation_stats_extra.xlsx
            rotated_ocr_text_per_file_and_page_df.to_excel(excel_output_rotated_ocr_text_path)


        not_rotated_ocr_text_per_file_and_page_df['not_rotated_ocred_page_text'].replace(to_replace='^=', value=' =',regex=True, inplace=True)
        not_rotated_ocr_text_per_file_and_page_df['file_name'],not_rotated_ocr_text_per_file_and_page_df['page_number'] = zip(*not_rotated_ocr_text_per_file_and_page_df.index)

        not_rotated_ocr_text_per_file_and_page_df.reset_index( drop=True, inplace=True)
        if see_not_joined_files:
            #NOTE: Could remove individual write out as data in page_orientation_stats_extra.xlsx
            not_rotated_ocr_text_per_file_and_page_df.to_excel(excel_output_not_rotated_ocr_text_path)




    """

    NOTE: Wouldn't have to store these two dfs alone, could always store the joined result.
    Save below or just use to join ocr text per page

    page_text == embedded_text already stored in PDF file

    'merged_text' == column used to blend the PDFs pages together into one file as .txt file.

    1: Check for Nan values in embedded_text_per_file_and_page_df. If Nan then merge with whatever is in ocr_text_per_file_and_page

    2: Check for set differences in embedded_text to see if there is a part of a page which is a scan/image which the OCR caught
    If OCR caught then add that to the page.



    """

    embedded_text_per_file_and_page_df=pd.DataFrame.from_dict({(i,j): embedded_text_per_file_and_page[i][j] 
                               for i in embedded_text_per_file_and_page.keys() 
                               for j in embedded_text_per_file_and_page[i].keys()},
                           orient='index', columns=['page_text'])



    #result of applying the page orientation metric in deciding if the page is rotated or not.
    #This determines which OCR results( rotated, not, or both because too close to tell) are to be used. 
    ocr_text_per_file_and_page_df=pd.DataFrame.from_dict({(i,j): ocr_text_per_file_and_page[i][j] 
                               for i in ocr_text_per_file_and_page.keys() 
                               for j in ocr_text_per_file_and_page[i].keys()},
                           orient='index', columns=['orientation_metric_ocred_page_text'])
    #


    #errors with excel it thinks there is a formula 
    embedded_text_per_file_and_page_df['page_text'].replace(to_replace='^=', value=' =',regex=True, inplace=True)

    embedded_text_per_file_and_page_df['file_name'],embedded_text_per_file_and_page_df['page_number'] = zip(*embedded_text_per_file_and_page_df.index)
    #Reset index to get a random index and then be able to do the merging.
    embedded_text_per_file_and_page_df.reset_index( drop=True, inplace=True)



    #Verify this should handle Plumbing-Permit-1.pdf page 2 and use its OCR. It does.
    embedded_text_per_file_and_page_df.loc[:,'use_ocr']= embedded_text_per_file_and_page_df.apply(use_ocr_text, axis=1)

    #if not minimial_file_output:
    if see_not_joined_files:
        embedded_text_per_file_and_page_df.to_excel(excel_output_embedded_text_per_page_and_file_path)

    """
    NOTE: orientation_metric_ocred_page_text is only provided for use_ocr_x == True due to join. 
    """
    ocr_text_per_file_and_page_df['orientation_metric_ocred_page_text'].replace(to_replace='^=', value=' =',regex=True, inplace=True)
    ocr_text_per_file_and_page_df['file_name'],ocr_text_per_file_and_page_df['page_number'] = zip(*ocr_text_per_file_and_page_df.index)

    ocr_text_per_file_and_page_df.reset_index( drop=True, inplace=True)

    #if not minimial_file_output:
    if see_not_joined_files:
        ocr_text_per_file_and_page_df.to_excel(excel_output_ocr_text_per_page_and_file_path)


    # Case 1: Check for Nan values in embedded_text_per_file_and_page_df. If Nan then merge with whatever is in ocr_text_per_file_and_page
    #join OCR page text on embedded_text_empty_pages
    embedded_text_empty_pages=embedded_text_per_file_and_page_df.loc[embedded_text_per_file_and_page_df['use_ocr'], ]

    #Note page_text blank on this due to setup and could be removed as column.
    # Good test example is '2017-Policy-and-Application-Town-Co-Sponsorship-Events-3.pdf'
    merged_ocr_text=pd.merge(embedded_text_empty_pages.copy(),ocr_text_per_file_and_page_df, how='left', on=['page_number','file_name'])

    #Removing suffix by not including in merge
    #merged_ocr_text=pd.merge(embedded_text_empty_pages[,['page_number','file_name']].copy(),ocr_text_per_file_and_page_df, how='left', on=['page_number','file_name'])



    #Empty embedded text pages will use whatever the OCR captured.
    merged_ocr_text['merged_text'] = merged_ocr_text['orientation_metric_ocred_page_text']
    #print("merged_ocr_text columns:{}".format(merged_ocr_text.columns))

    if not old_output:
    #if not minimial_file_output:
        merged_ocr_text.to_excel(excel_output_merged_text_path)



    merged_text_final=pd.merge(embedded_text_per_file_and_page_df, merged_ocr_text,how='left', on=['file_name','page_number'])


    #print(merged_text_final.columns)
    #NOTE: _x suffix just denotes leftmost df columns which are common to the other df used in join. _y is rightmost df columns 
    #NOTE: Program currently only supports using embedded text xor OCR text for a single page--can't use both or combine the results of both with below.
    # if 'use_ocr_x' == False then use the embedded text for that row as merged_text column
    #merged_text_final.loc[merged_text_final['use_ocr_x'] == False,'merged_text'] = merged_text_final.loc[merged_text_final['use_ocr_x'] == False,'page_text_x']#not joining
    merged_text_final.loc[merged_text_final['use_ocr_x'] == False,'merged_text'] = merged_text_final['page_text_x']
    #merged_text_final.loc[merged_text_final['merged_text'].isnull(),'merged_text'] = merged_text_final.loc[merged_text_final['page_text_x'].notnull(),'page_text']


    """
    32,767 characters
    Microsoft Excel has a character limit of 32,767 characters in each cell.

    The excel file wasn't showing all of the text in a given cell but if one double clicks on the cell then more shows up. 
    TODO: Load in and try to print and see if all the text prints or if excel limit comes into play.
    Could always just save the files as plain text and load in when needed if excel loses part of the string due to char limit 
    """

    #show all text for a given document in the 'all_pages_merged' column
    #merged_text_final['all_pages_merged'] = merged_text_final[['file_name','merged_text']].groupby('file_name')['merged_text'].transform(lambda x: ' '.join(x))

    #debugging
    if not old_output:
    #if not minimial_file_output:
        merged_text_final.to_excel(excel_output_merged_text_per_page_before_joining_to_file_path)


    #merge all pages into one df. Need 'ocred_page_text' + 'page_text' 
    #merged_text_final.groupby("file_name")['text'].apply(list)

    #Keeping as a feature 
    #merged_text_final.groupby("file_name", as_index = False).agg({'merged_text': list})
    #merged_text_final.groupby("file_name", as_index = True).agg({'merged_text': list})


    #Here the ' ' would be the char joining pages of the PDF
    #Apply below to make sure pages are coming out in correct order
    merged_text_final.sort_values(by = ["file_name","page_number"], inplace=True)
    document_pages_merged=merged_text_final.groupby("file_name")['merged_text'].apply(' '.join).reset_index()
    #document_pages_merged=merged_text_final.groupby("file_name")['merged_text'].apply(' '.join).reset_index().copy()
    #print(merged_text_final.groupby("file_name")['merged_text'].apply(' '.join))
    #merged_text_final.groupby("file_name")['merged_text'].apply(' '.join).reset_index().to_excel(excel_output_merged_final_text_path)#only joins first two pages. 
    #document_pages_merged=merged_text_final.groupby("file_name")['merged_text'].apply(' '.join).reset_index()

    #
    see_truncation = False
    if see_truncation:
        document_pages_merged.to_excel(excel_output_merged_final_text_path)#DON'T DO THIS truncates past char limit
    #merged_text_final.to_excel(excel_output_merged_final_text_path)
    #TODO: Case 2 check if anything is in the OCR page which isn't in the embedded text page. Might not need this case. Haven't seen an example of this with embedded text and then a scan/image of text on the same page. Have seen a logo though as a scan and text. 
    """
    for key in document_pages_merged:
        print("Group:{}\tValues:{}".format(key,document_pages_merged.get_group(key)))

    """




    """
    Write each row of document_pages_merged( a PDF file from website) to a file to avoid the char limit of excel.

    This plain text files will be the input for the tokenizer and model.Need to put them into df. 

    .pdf.txt is file name change unless file_name_wo_extension is used.

    Plain text and word documents should also be written to this folder as plain text.
    """

    for i, row in document_pages_merged.iterrows():
        file_merged_text = '{}'.format(row.merged_text)
        file_name_wo_extension=row.file_name.replace(".pdf","")

        #print("Row columns:{}".format(row.columns))
        #print(row)
        #NOTE: downstream in program will attempt to match on file name notice the .txt is added to the end of the .pdf
        #Could remove .pdf then match on file name without extension
        #file_name_wo_extension=row.file_name.replace(".pdf","")
        #with open(r'{}\{}.txt'.format(path_to_ocr_output_dir, file_name_wo_extension), 'w',errors='surrogateescape') as text_file:

        #TODO: Check if unknown chars are in the individual output files 
        #Note  errors='ignore' writes out but then keeps the char in the file. Attempting other args for errors
        #with open(r'{}\{}.txt'.format(path_to_ocr_output_dir, row.file_name), 'w',errors='replace') as text_file:
        with open(r'{}\{}.txt'.format(path_to_ocr_output_dir, file_name_wo_extension), 'w',errors='backslashreplace') as text_file:

            text_file.write(file_merged_text)
            text_file.close()

    #Writing out docx files to the same folder as the converted PDFs to text & .txt files
    for file_name, text in docx_file_dict.items():
        file_name_wo_extension=file_name.replace(".docx","")
        #with open(r'{}\{}.txt'.format(path_to_ocr_output_dir, file_name), 'w',errors='replace') as text_file:
        with open(r'{}\{}.txt'.format(path_to_ocr_output_dir, file_name_wo_extension), 'w',errors='backslashreplace') as text_file:

            text_file.write(text)
            text_file.close()


    #Writing out txt files to the same folder as the converted PDFs to text & docx files
    for file_name, text in txt_file_dict.items():

        with open(r'{}\{}.txt'.format(path_to_ocr_output_dir, file_name), 'w',errors='backslashreplace') as text_file:

            text_file.write(text)
            text_file.close()



    """
    File denoting skipped pdfs due to page length limit
    """
    skipped_pdfs = pd.DataFrame.from_dict(skipped_pdf_dict,orient="index",columns=["page_count"])
    skipped_pdfs.index.name = "file_name"
    skipped_pdfs.to_excel(excel_output_skipped_pdfs_path)




    """
    Stats on which text was used for a given page: rotated page or not, and word counts for those pages.

    ocr_text_orientation_stats_per_file_and_page[file]

    ocr_text_orientation_stats_per_page_dict[page_number] = (rotated_valid_word_count,valid_word_count)


                ocr_text_orientation_stats_per_page_dict[page_number] = (rotated_valid_word_count,valid_word_count,unique_rotated_valid_word_count,unique_valid_word_count,rotated_word_char_length_weighted_value,word_char_length_weighted_value,rotated_page_metric,not_rotated_page_metric,within_degree_of_closeness)

    valid_word_count == not rotated valid word count
    rotated_valid_word_count == rotated page and applied ocr to it, and counted valid words from brown nlp dataset
    """
    #With 'within_degree_of_closeness'
    ocr_text_columns = ['rotated_valid_word_count','valid_word_count','unique_rotated_valid_word_count','unique_valid_word_count','rotated_word_char_length_weighted_value','word_char_length_weighted_value','rotated_page_metric','not_rotated_page_metric','within_degree_of_closeness']

    #Without 'within_degree_of_closeness' incase joining rotated and original is a bad idea. 
    #ocr_text_columns = ['rotated_valid_word_count','valid_word_count','unique_rotated_valid_word_count','unique_valid_word_count','rotated_word_char_length_weighted_value','word_char_length_weighted_value','rotated_page_metric','not_rotated_page_metric','within_degree_of_closeness']


    text_orientation_stats=pd.DataFrame.from_dict({(i,j): ocr_text_orientation_stats_per_file_and_page[i][j] 
                               for i in ocr_text_orientation_stats_per_file_and_page.keys() 
                               for j in ocr_text_orientation_stats_per_file_and_page[i].keys()},
                           orient='index', columns=ocr_text_columns)


    #text_orientation_stats['used_rotated_page'] = text_orientation_stats['rotated_page_metric'] > text_orientation_stats['not_rotated_page_metric']

    #text_orientation_stats.to_excel(excel_output_page_orientation_stats_path)

    #TODO: Join the used_ocr on this to see if there was not embedded text on the page and if the system had to depend on the rotation being correct.
    #In embedded_text_per_page but will need to make the features file_name & page_number into tuple and join on index of this.

    text_orientation_stats['file_name'],text_orientation_stats['page_number'] = zip(*text_orientation_stats.index)



    #Shows which pages got used for OCR debugging
    text_orientation_stats_used_pages=pd.merge(merged_text_final, text_orientation_stats,how='left', on=['file_name','page_number'])
    #OCR is needed and rotated page has a higher metric.
    text_orientation_stats_used_pages.loc[text_orientation_stats_used_pages['use_ocr_x']== True,'used_rotated_page_ocr'] = text_orientation_stats_used_pages['rotated_page_metric'] > text_orientation_stats_used_pages['not_rotated_page_metric']


    #text_orientation_stats_used_pages.loc[text_orientation_stats_used_pages['page_text_x'].isnull() and text_orientation_stats_used_pages['ocred_page_text'].isnull(),'no_text_ocr_or_embedded'] = True
    #Seeing if both columns are Nan and if they are then set to True. Will also check if the columns are the same in general, but the code shouldn't allow this to be the case beyond the Nan Example.
    #NOTE better way to do below.
    text_orientation_stats_used_pages['no_text_ocr_or_embedded'] = text_orientation_stats_used_pages.apply(lambda row: True if str(row['page_text_x']) == str(row['orientation_metric_ocred_page_text']) else False, axis=1)

    if not old_output:
        text_orientation_stats_used_pages.to_excel(excel_output_page_orientation_stats_path)

    #merging page OCR text rotated and not rotated into this df.  All page stats in one place.
    #merging OCR rotated and not rotated into page stats
    text_orientation_stats_used_pages_merged_condensed=pd.merge(text_orientation_stats_used_pages,not_rotated_ocr_text_per_file_and_page_df ,how='left', on=['file_name','page_number'])
    text_orientation_stats_used_pages_merged_wo_embedded=pd.merge(text_orientation_stats_used_pages_merged_condensed,rotated_ocr_text_per_file_and_page_df ,how='left', on=['file_name','page_number'])
    text_orientation_stats_used_pages_merged_wo_embedded.to_excel(excel_output_page_orientation_stats_extra_path)


    #merging embedded text into page stats. Don't need because original leftmost df has embedded text in it
    #CHECK do I not need embedded because in the original df in column page_text_x?
    #text_orientation_stats_used_pages_merged_full=pd.merge(text_orientation_stats_used_pages_merged_wo_embedded, embedded_text_per_file_and_page_df,how='left', on=['file_name','page_number'])
    #Should be all the per page details for all files. Text limit shouldn't matter because a page is unlikely to go over the 35k char limit per excel cell.
    #text_orientation_stats_used_pages_merged_full.to_excel(excel_output_page_orientation_stats_extra_path)


    """
    Saving out valid words( in Brown nlp corpus/ nltk) per file per page.

    Debug structure for valid words to ensure OCR page orientation is operating correctly.

    rotated_ocr_text_valid_words_per_file_and_page 
    ocr_text_valid_words_per_file_and_page 

    excel_output_valid_words_rotated_ocr_text_per_page_and_file_path
    excel_output_valid_words_ocr_text_per_page_and_file_path = r"{}\{}".format(path_to_ocr_metadata_output_dir,output_excel_valid_words_ocr_text_per_page_and_file_name)#original embedded text in pdfs


    """


    #print(rotated_ocr_text_valid_words_per_file_and_page)


    if not minimial_file_output:
        valid_words_rotated_ocr_text=pd.DataFrame.from_dict({(i,j): rotated_ocr_text_valid_words_per_file_and_page[i][j] 
                                   for i in rotated_ocr_text_valid_words_per_file_and_page.keys() 
                                   for j in rotated_ocr_text_valid_words_per_file_and_page[i].keys()},
                               orient='index')

        valid_words_rotated_ocr_text.index.name = "file name and page number"

        valid_words_rotated_ocr_text.to_excel(excel_output_valid_words_rotated_ocr_text_per_page_and_file_path)








        valid_words_ocr_text=pd.DataFrame.from_dict({(i,j): ocr_text_valid_words_per_file_and_page[i][j] 
                                   for i in ocr_text_valid_words_per_file_and_page.keys() 
                                   for j in ocr_text_valid_words_per_file_and_page[i].keys()},
                               orient='index')
        valid_words_ocr_text.index.name = "file name and page number"


        valid_words_ocr_text.to_excel(excel_output_valid_words_ocr_text_per_page_and_file_path)



# In[ ]:





# In[8]:



end_time = time.perf_counter()
print(f"OCR conversion done in: {end_time - start_time:0.4f} seconds")


# In[ ]:





# In[ ]:




