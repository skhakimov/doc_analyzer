import os
import numpy as np
import collections
import csv
import codecs
import argparse
import PyPDF2
import optparse

from pdfminer.pdfinterp import PDFResourceManager, process_pdf
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from cStringIO import StringIO
from PyPDF2 import PdfFileReader

CSV_NAME = 'Similarity Table.csv'
ERROR_SUMMARY = 'Error Files.csv'
METADATA = 'PDF metadata.csv'

# prime number, rough guess for the number of unique words
prime = 33317

ignore = ['a','about', 'all', 'along', 'also', 'although', 'among','an',
'and', 'any','anyone', 'anything', 'are', 'around','at', 'be','because', 'been', 'before',
'being', 'both', 'but', 'came', 'come', 'coming', 'could', 'did', 'each',
'else', 'every', 'for', 'from', 'get', 'getting', 'going', 'got',
'gotten', 'had', 'has', 'have', 'having', 'her', 'here', 'hers', 'him',
'his', 'how', 'however','i', 'into', 'its', 'like', 'may', 'most', 'next','not',
'now', 'only', 'our', 'out', 'particular', 'same', 'she', 'should',
'some', 'take', 'taken', 'taking', 'than', 'that', 'the', 'then', 'there',
'these', 'they', 'this', 'those', 'throughout', 'to','too', 'took', 'very',
'was', 'went', 'what', 'when', 'which', 'while', 'who', 'why', 'will',
'with', 'without', 'would', 'yes', 'yet', 'you', 'your']

#extracts metadata from a folder with pdf files and stores results as a csv file
def getMetacsv(folder):	
  csvfile = open(METADATA,'w')	
  csvwriter = csv.writer(csvfile, dialect = 'excel')
  csvwriter.writerow(['FILENAME','Author', 'Company','Producer','Title','Creator','Creation Date','Modified Date','Subject','Keywords'])
	
  for filename in os.listdir(folder):
    metadata=[]
    try:
      if '.pdf' in filename:		
        pdfFile = PdfFileReader(file(folder+'/'+filename, 'rb'))
        docInfo = pdfFile.getDocumentInfo()
	  if '/Author' in docInfo:
	    metadata.append(docInfo['/Author'].strip())
	  else:
	    metadata.append('')
	  if '/Company' in docInfo:
	    metadata.append(docInfo['/Company'].strip())
	  else:
	    metadata.append('')		
	  if '/Producer' in docInfo:
		metadata.append(docInfo['/Producer'].strip())
	  else:
	    metadata.append('')
	  if '/Title' in docInfo:
	    metadata.append(docInfo['/Title'].strip())
	  else:
	    metadata.append('')
	  if '/Creator' in docInfo:
		metadata.append(docInfo['/Creator'].strip())
	  else:
		metadata.append('')
	  if '/CreationDate' in docInfo:
		metadata.append(docInfo['/CreationDate'].strip())
	  else:
		metadata.append('')	
	  if '/ModDate' in docInfo:
		metadata.append(docInfo['/ModDate'].strip())
	  else:
		metadata.append('')
	  if '/Subject' in docInfo:
		metadata.append(docInfo['/Subject'].strip())
	  else:
		metadata.append('')	
	  if '/Keywords' in docInfo:
        metadata.append(docInfo['/Keywords'].strip())
      else:
        metadata.append('')	
				
      csvwriter.writerow([filename]+metadata)
				
    except:
      continue
		
#takes a pdf file as input and outputs text contents of the file as a string
def convert_pdf(path):
  rsrcmgr = PDFResourceManager()
  retstr = StringIO()
  codec = 'utf-8'
  laparams = LAParams()
  device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)

  fp = file(path, 'rb')
  process_pdf(rsrcmgr, device, fp)
  fp.close()
  device.close()

  str = retstr.getvalue()
  retstr.close()
  
  return str
    
    
#hash function for our purposes.  This is to assign to each word a unique hash value
def word_hash(string):
	fac = 256
	h=0
	for i in string:
		h = (h*fac+ord(i))%prime 
	h=h+1   # hash values. ranging [1 to prime]
	
	return h
	
	
# computes cosine given two vectors in an array format.  Ranges 0-1	
def cos_vector(input1, input2):	
	#dot product divided by products of norms
	result = round(np.dot(input1, input2) / (np.linalg.norm(input1)*np.linalg.norm(input2)),3) 

	return result
	
		
# generates a dictionary of hash values given an input of a list consisting of words. E.g. result['apple']=1234, result['orange']=9304, etc.
def dictionary_hash(data):
  result = {}
  for i in xrange(len(data)):#for all unique words
    j = word_hash(data[i])#produce a hash of each word and assign an index j
    result[data[i]]=j-1#hash values of all unique words

	return result	
	
def dictionary_vector(dict_hash, dict_freq):
  result = {} #key:filename, value: vector form of a file
  for filename in dict_freq:
    vector = np.zeros(prime)
    for word in dict_freq[filename]: #for every unique word in the file
      vector[dict_hash[word]]=dict_freq[filename][word] #the specific dimension of the vector is frequency of that word in a file
      result[filename]=vector
      
      return result	
      		
	
def word_frequency(folder):
  unique_all = np.zeros(1)		
  dict_freq = {}
  error_files = []
  text_files = os.listdir(folder)
	
  if '.DS_Store' in text_files:
    text_files.remove('.DS_Store')

  for filename in os.listdir(folder):		
    try:
      if '.pdf' in filename:
        text = convert_pdf(folder+'/'+filename).lower()#extract contents of a pdf file as a lowercase string
      else:		
        text = open(folder+'/'+filename).read().lower()#extract contents of a file as a lowercase string
		
    except Exception,e: 
      text_files.remove(filename)
      error_files.append(filename)
      print '[-] '+ filename + ' UNABLE TO PROCESS: ' + str(e)
      continue

  words = ''.join(c for c in text if c.isalpha() or c.isspace()).split()#create a list of words from a file

  words_filter = [c for c in words if len(c)>2 and len(c)<20 and c not in ignore] #remove words less than 3 and greater than 20 in lengths

  words_array = np.array(words_filter)#convert the list of words in a file into an np array

  unique = np.unique(words_array)#extract unique words from a file

  unique_all = np.concatenate((unique_all, unique),axis=0)#create an array of unique words from all files

  words_array_freq = collections.Counter(words_array)#calculate frequencies of each word in a file by creating a dictionary with words as keys

  dict_freq[filename] = words_array_freq#assign frequency array to a filename which is a key
		
  final_unique = np.unique(unique_all)#remove duplicates to create an array of unique words from all files

  dict_hash = dictionary_hash(final_unique)#dictionary of indexes assigned to all unique words.  i.e. dict_hash['apple'] = 1234, etc.
	
  # creates a dictionary of filenames and their vector-form representation
  dict_vector = dictionary_vector(dict_hash, dict_freq)
	
  dict_similarity = {}

  dict_similarity['filenames'] = text_files
	
  if len(error_files)>0:
    dict_similarity['error_files'] = error_files

  for file in text_files:
    filevector1 = dict_vector[file]
    similarity_list = []
    for other in text_files:
      filevector2 = dict_vector[other]
      cosine = cos_vector(filevector1, filevector2)	
      similarity_list.append(cosine)		
      dict_similarity[file] = similarity_list
	
  return dict_similarity	


def save_csv(data):
  csvfile = open(CSV_NAME,'w')
  csvwriter = csv.writer(csvfile, dialect='excel')
  first_row = [''] + data['filenames']
  csvwriter.writerow(first_row)
  
  for file_name in data['filenames']:	
    row = [file_name] + data[file_name]
    csvwriter.writerow(row)
		
  if 'error_files' in data:		
    csvfile = open(ERROR_SUMMARY,'w')
	csvwriter = csv.writer(csvfile, dialect = 'excel')
	
  for filename in data['error_files']:
    csvwriter.writerow([filename])
	
def main():

	parser = argparse.ArgumentParser(description = "Script to quantify similarity relationships between documents and produce pdf metadata summary. \
	Outputs two csv files: Similarity(correlation) table and summary table with pdf metdata. \
	If could not access any files, output a list in a csv file\
	For more info on the technique: http://en.wikipedia.org/wiki/Latent_semantic_analysis")

	parser.add_argument('inputfolder', metavar = 'INPUTFOLDER', type = str, help = 'Folder with files to process and analyze')

	
	args = parser.parse_args()

	
	if not os.path.isdir(args.inputfolder):
		print '[-] Folder ' + args.inputfolder + ' does not exist.'
		exit(0)
		
	
	data = word_frequency(args.inputfolder)
	
	save_csv(data)	
	
	getMetacsv(args.inputfolder)
	
	    
if __name__ =="__main__":
  main()