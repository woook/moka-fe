"""
Created on 23 Jan 2017

This script is used to import primers designed by the PHIGIS program into MOKA.

PHIGIS exports primers in a set format. These files are stored in a set location on the shared drive and this script will loop through the files and import these into the database.

It does not matter what order the primers are imported as the database assigns a incrementing value as the primer name.

Default values are specified in the config file. These values are in the primerlookup table and the keys for these values are pulled out and used when inserting the primer into the database

@author: ajones7
"""

import pyodbc
import os

# import from config file
from ImportPrimerconfig import *

class ImportPrimer():
	def __init__(self):
		# using pyodbc specify database connection details
		#self.cnxn = pyodbc.connect(mokadata_connection) # use the mokadata database connection details specified in the config file
		self.cnxn = pyodbc.connect(dev_database_connection) # use the devdatabase database connection details specified in the config file
		# create a cursor to connect to database 
		self.cursor = self.cnxn.cursor()
		
		# dictionary used to record the coordinates for each gene in the Pan493 bed file
		self.bedfile_dict={}
		
		# build string to search primerlookup table using value in config file
		self.status_string="\'"+status_string+"\'" 
		# variable to hold the lookup value returned
		self.status_id=""
		# build string to search primerlookup table using value in config file
		self.purification_text="\'"+purification_text+"\'"
		# variable to hold the lookup value returned
		self.purification_id=""
		# build string to search primerlookup table using value in config file
		self.scale_of_synth_text="\'"+scale_of_synth_text+"\'"
		# variable to hold the lookup value returned
		self.scale_of_synth_id=""
		# build string to search primerlookup table using value in config file
		self.genome_build_text="\'"+genome_build_text+"\'"
		# variable to hold the lookup value returned
		self.genome_build_id=""
		
		#HighestPrimerName
		self.highest_primer_name=0
		
		# variables to hold queries and exceptions
		self.insert_query = ""
	
	
	def get_gene(self):
		"""
		The Pan493 bedfile is used to provide all the genes in the exome and the genomic coordinates
		The bed file has one exon per line, with the chromosome, start, stop and entrezID in a tab delimited file. 
		This is entered into a dictionary so if the primer overlaps with a gene the gene can be recorded.
		"""
		# open the bedfile as read only using the file path from the config file
		with open(pan493_bedfile,'r') as bedfile:
			# loop through each line
			for line in bedfile:
				#split the line on tab
				splitline=line.split("\t")
				# append the line to the dictionary with a tuple of the chromosome, start and stop as a key and the entrez id as the value (with the new line character stripped from the entrezid)
				self.bedfile_dict[(splitline[0],splitline[1],splitline[2])] = splitline[-1].rstrip()
		
	def get_moka_lookup_values(self):
		"""
		The human readable strings for some fields with default values are specified in the config file. 
		Select queries are used to return the keys for each value
		"""
		
		# set the query to extract the status_string
		query="select StatusID from status where status = " + self.status_string
		# pass the query to the fetch one function which returns the query as a list with a single object (using pyodbc fetchone)
		result=self.fetchone(query)
		# extract the value from the list, convert to a string and assign to the variable 
		self.status_id = str(result[0])
		
		# set the query to extract the purification
		query="select PrimerLookupID from PrimerLookup where Name = " + self.purification_text
		# pass the query to the fetch one function which returns the query as a list with a single object (using pyodbc fetchone)
		result=self.fetchone(query)
		# extract the value from the list, convert to a string and assign to the variable 
		self.purification_id= str(result[0])
		
		# set the query to extract the ScaleOfSynth
		query="select PrimerLookupID from PrimerLookup where Name = " + self.scale_of_synth_text
		# pass the query to the fetch one function which returns the query as a list with a single object (using pyodbc fetchone)
		result=self.fetchone(query)
		# extract the value from the list, convert to a string and assign to the variable 
		self.scale_of_synth_id= str(result[0])
		
		# set the query to extract the GenomeBuild
		query="select PrimerLookupID from PrimerLookup where Name = " + self.genome_build_text
		# pass the query to the fetch one function which returns the query as a list with a single object (using pyodbc fetchone)
		result=self.fetchone(query)
		# extract the value from the list, convert to a string and assign to the variable 
		self.genome_build_id= str(result[0])

		
		
	def parse_input(self):
		"""
		
		"""
		#count files parsed and primers imported
		file_count=0
		primer_count=0
		
		# for each file in the directory
		for file in os.listdir(primer_design_files):
			if file.endswith('.txt'):
				file_count+=1
				#open file and loop through
				with open(primer_design_files+"//"+file,'r') as primer_designs:
					for line_number, line in enumerate(primer_designs):
						# for each primer:
						if line_number > 0 and len(line) >3:
							splitline=line.split("\t")
							#first element is in format chr16_70534805-70535059
							chr=splitline[0].split("_")[0]
							primer_start=splitline[0].split("_")[1].split("-")[0]
							primer_stop=splitline[0].split("_")[1].split("-")[1]
							
							#second element is the variant id to go into notes
							variant_id=splitline[1]
							
							#capture primer sequences
							forward_sequence=splitline[6]
							reverse_sequence=splitline[7]
							
							#convert the chromosome to numeric form, without 'chr' 
							if chr=='chrX':
								cleaned_chromosome_number=23
							elif chr=='chrY':
								cleaned_chromosome_number=24
							else:
								cleaned_chromosome_number=chr.replace('chr','')
							
							#empty EntrezGeneID variable
							entrez_gene_id_mapped=''
							
							# loop through the bedfile dictionary to match the primer to an entrez id 
							for gene in self.bedfile_dict:
								#use chr rather than the cleaned values as the bedfile uses X and Y
								if gene[0]==chr.replace("chr","") and gene[1]<primer_stop and gene[2]>primer_start:
									entrez_gene_id_mapped=self.bedfile_dict[gene]
							
							#create a dict of all variables and pass to self.build_insert module
							primer_values_to_insert_dict={"chr":str(cleaned_chromosome_number),"start":str(primer_start),"stop":str(primer_stop),"gene":str(entrez_gene_id_mapped),"FSeq":forward_sequence,"RSeq":reverse_sequence,"FTag":f_tag,"RTag":r_tag,"status":self.status_id,"purification":self.purification_id,"SoS":self.scale_of_synth_id,"GenomeBuild":self.genome_build_id,"notes":variant_id}
							
							self.build_insert(primer_values_to_insert_dict)
							primer_count+=1
				
				# move the file into an archive folder
				os.rename(primer_design_files+"//"+file, archived_primer_design_files+"//"+file)
				#print "archiving file"
		
		# report progress to message box
		if file_count == 0:
			print "no files found to import. Please ensure that files are in " + primer_design_files
		else:
			# report progress to message box
			print str(primer_count) + " primers imported from " + str(file_count) + " files"

	
	def build_insert(self,dict):
		# HighestGeneName
		query="select PrimerName from PrimerAmplicon union select '0' as PrimerName from dbo.PrimerLookup order by PrimerName desc"
		result=self.fetchone(query)
		highest_primer_name= result[0]+1
		self.highest_primer_name=highest_primer_name
		
		#build sql insert statement
		self.insert_query='INSERT INTO "dbo"."PrimerAmplicon" ("EntrezGeneIDmapped", "ForwardSeq", "ReverseSeq", "ChromosomeID", "Start", "Stop","status","FTag","RTag","ScaleOfSynth","Purification","GenomeBuild", "notes","PrimerName") VALUES ('+dict['gene']+',\''+dict['FSeq']+'\',\''+dict['RSeq']+'\','+dict['chr']+','+dict['start']+','+dict['stop']+','+dict['status']+','+dict['FTag']+','+dict['RTag']+','+dict['SoS']+','+dict['purification']+','+dict['GenomeBuild']+',\''+dict['notes']+'\','+str(highest_primer_name)+')'
		print self.insert_query
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

def main():
	a=ImportPrimer()
	a.get_gene()
	a.get_moka_lookup_values()
	a.parse_input()

	
if __name__ == '__main__':
	main()
