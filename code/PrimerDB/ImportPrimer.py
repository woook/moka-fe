'''
Created on 23 Jan 2017

@author: ajones7
'''
import pyodbc
import os

# import from config file
from ImportPrimerconfig import *

class Import_primer():

	def __init__(self):
		# variables for the database connection
		self.cnxn = pyodbc.connect(mokadata_connection)
		#self.cnxn = pyodbc.connect(devdatabase_connection)
		self.cursor = self.cnxn.cursor()
		
		
		# Use of bed file to get gene symbol
		self.bedfile_dict={}
		
		# dict for each primer insert statement
		self.dict={}
		
		# default tag values
		self.FTag=FTag
		self.RTag=RTag
		# status
		self.status_string="\'"+status_string+"\'" 
		self.statusID=""
		# purification
		self.purificationText="\'"+purificationText+"\'"
		self.purificationID=""
		#scale_of_synth
		self.scaleofsynthText="\'"+scaleofsynthText+"\'"
		self.scaleofsynthID=""
		#genome build
		self.GenomeBuildText="\'"+GenomeBuildText+"\'"
		self.GenomeBuildID=""
		
		# variables to hold queries and exceptions
		self.select_qry = ""
		self.insert_query = ""
		self.select_qry_exception = ""
		self.backup_qry = ""
	
	
	def get_gene(self):
		# read the bedfile and populate a dict with a chr, start, stop tuple as key and entrezid as value
		bedfile=open(pan493bedfile,'r')
		for line in bedfile:
			splitline=line.split("\t")
			self.bedfile_dict[(splitline[0],splitline[1],splitline[2])]= splitline[-1].rstrip()
		
	def get_moka_lookup_values(self):
		'''read lookup tables to extract the values to insert'''
		# statusID
		qry="select statusID from status where status = " + self.status_string
		
		result=self.fetchone(qry)
		self.statusID = str(result[0])
		# purification
		qry="select PrimerLookupID from PrimerLookup where Name = " + self.purificationText
		
		result=self.fetchone(qry)
		self.purificationID= str(result[0])
		# ScaleOfSynth
		qry="select PrimerLookupID from PrimerLookup where Name = " + self.scaleofsynthText
		result=self.fetchone(qry)
		self.scaleofsynthID= str(result[0])
		# GenomeBuild
		qry="select PrimerLookupID from PrimerLookup where Name = " + self.GenomeBuildText
		result=self.fetchone(qry)
		self.GenomeBuildID= str(result[0])
		
	def parse_input(self):
		#count files parsed and primers imported
		filecount=0
		primer_count=0
		
		# for each file in the directory
		for file in os.listdir(primer_design_files):
			if file.endswith('.txt'):
				filecount+=1
				#open file and loop through
				with open(primer_design_files+"//"+file,'r') as f:
					for i, line in enumerate(f):
						# for each primer:
						if i > 0 and len(line) >3:
							splitline=line.split("\t")
							#first element is in format chr16_70534805-70535059
							Chr=splitline[0].split("_")[0]
							sta=splitline[0].split("_")[1].split("-")[0]
							sto=splitline[0].split("_")[1].split("-")[1]
							#capture primer sequences
							FSeq=splitline[6]
							RSeq=splitline[7]
							
							#convert the chromosome to numeric form, without 'chr' 
							if Chr=='chrX':
								cleaned_chr=23
							elif Chr=='chrY':
								cleaned_chr=24
							else:
								cleaned_chr=Chr.replace('chr','')
							
							#empty EntrezGeneID variable
							EntrezGeneIDmapped=''
							
							# loop through the bedfile dictionary to match the primer to an entrez id 
							for gene in self.bedfile_dict:
								#use Chr rather than the cleaned values as the bedfile uses X and Y
								if gene[0]==Chr.replace("chr","") and gene[1]<sto and gene[2]>sta:
									EntrezGeneIDmapped=self.bedfile_dict[gene]
							
							#create a dict of all variables and pass to self.build_insert module
							dict={"Chr":str(cleaned_chr),"start":str(sta),"stop":str(sto),"gene":str(EntrezGeneIDmapped),"FSeq":FSeq,"RSeq":RSeq,"FTag":self.FTag,"RTag":self.RTag,"status":self.statusID,"purification":self.purificationID,"SoS":self.scaleofsynthID,"GenomeBuild":self.GenomeBuildID}
							self.build_insert(dict)
							primer_count+=1
				
				# move the file into an archive folder
				os.rename(primer_design_files+"//"+file, archived_primer_design_files+"//"+file)
				#print "archiving file"
		
		# report progress to message box
		if filecount == 0:
			print "no files found to import. Please ensure that files are in " + primer_design_files
		else:
			# report progress to message box
			print str(primer_count) + " primers imported from " + str(filecount) + " files"

	
	def build_insert(self,dict):
		#build sql insert statement
		self.insert_query='INSERT INTO "dbo"."PrimerAmplicon" ("EntrezGeneIDmapped", "ForwardSeq", "ReverseSeq", "ChromosomeID", "Start", "Stop","status","FTag","RTag","ScaleOfSynth","Purification","GenomeBuild") VALUES ('+dict['gene']+',\''+dict['FSeq']+'\',\''+dict['RSeq']+'\','+dict['Chr']+','+dict['start']+','+dict['stop']+','+dict['status']+','+dict['FTag']+','+dict['RTag']+','+dict['SoS']+','+dict['purification']+','+dict['GenomeBuild']+')'
		#print self.insert_query
		# call module to execure insert query
		self.insert_query_function()

	
	def insert_query_function(self):
		'''This function executes an insert query'''
		# execute the insert query
		self.cursor.execute(self.insert_query)
		self.cursor.commit()
		#print "bosh"
	
	
	def fetchone(self, query):
		'''perform select query but return single result'''
		#perform query
		self.cursor.execute(query)
		#capture result
		result = self.cursor.fetchone()
		
		#yield result
		if result:
			return result
		else:
			print "unable to return result for query " + query
		
if __name__ == '__main__':
	a=Import_primer()
	a.get_gene()
	a.get_moka_lookup_values()
	a.parse_input()
