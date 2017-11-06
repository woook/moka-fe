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
		Primer designs are read from files within a specific folder.
		This function captures all the required information needed to insert the primer into the database and passes this information to a function which executes the insert query, one primer at a time.
		"""
		# count files parsed and primers imported to be reported at the end
		file_count=0
		primer_count=0
		
		# for each file in the directory of primer designs
		for file in os.listdir(primer_design_files):
			# look for all the text files
			if file.endswith('.txt'):
				# add to the file count
				file_count+=1
				# open file as read only
				with open(primer_design_files+"//"+file,'r') as primer_designs:
					# loop through line by line, using enumerate so the first line can be skipped
					for line_number, line in enumerate(primer_designs):
						# skipping first line, and making sure the line isn't empty:
						if line_number > 0 and len(line) >3:
							# split the line on tabs into a list 
							splitline=line.split("\t")
							# take the first column in the file, which is genomic coordinates in the format chr16_70534805-70535059
							# everything before the underscore is the chromosome
							chr=splitline[0].split("_")[0]
							# everything after the underscore, but before the dash is the start position
							primer_start=splitline[0].split("_")[1].split("-")[0]
							# everything after the underscore, and after the dash is the stop position
							primer_stop=splitline[0].split("_")[1].split("-")[1]
							
							# second column in the file is the variant id  - this is to go into notes column
							variant_id=splitline[1]
							
							# capture forward and reverse primer sequences from column 7 and 8
							forward_sequence=splitline[6]
							reverse_sequence=splitline[7]
							
							# to enable a join with the chromosome table convert the sex chromosomes to numeric form, and remove 'chr' 
							# change X to 23 and remove chr
							if chr=='chrX':
								cleaned_chromosome_number=23
							# change Y to 24 and remove chr
							elif chr=='chrY':
								cleaned_chromosome_number=24
							# remove 'chr' from all autosomes
							else:
								cleaned_chromosome_number=chr.replace('chr','')
							
							# create an empty EntrezGeneID variable
							entrez_gene_id_mapped=''
							
							# loop through the bedfile dictionary to match the primer to an entrez id 
							for gene in self.bedfile_dict:
								# The bedfile contains chromosome names in the format 'X' and 'Y' so we must convert the chromosome names from the design file to match this
								# each gene in the bedfile dict has a tuple key of (chr,start,stop), the entrez gene id is the value
								# for each gene in Pan493 look if the primer overlaps with the gene 
								if gene[0]==chr.replace("chr","") and gene[1]<primer_stop and gene[2]>primer_start:
									# capture the entrez gene id
									entrez_gene_id_mapped=self.bedfile_dict[gene]
							# NB if a primer overlaps with multiple genes only one will be captured
							
							
							#create a dict of all variables to be inserted to the database and pass to self.build_insert module
							primer_values_to_insert_dict={"chr":str(cleaned_chromosome_number),"start":str(primer_start),"stop":str(primer_stop),"gene":str(entrez_gene_id_mapped),"FSeq":forward_sequence,"RSeq":reverse_sequence,"FTag":f_tag,"RTag":r_tag,"status":self.status_id,"purification":self.purification_id,"SoS":self.scale_of_synth_id,"GenomeBuild":self.genome_build_id,"notes":variant_id}
							
							# pass dictionary to insert query
							self.build_insert(primer_values_to_insert_dict)
							# add to count of inserted primers
							primer_count+=1
				
				# move the file design file into an archive folder
				os.rename(primer_design_files+"//"+file, archived_primer_design_files+"//"+file)
				# NB errors may occur if a file with the same name already exists in the archive - but this shouldn't happen routinely
		
		# when all files  have been read report progress (this should be displayed in a moka message box)
		if file_count == 0:
			# if no files were found to be imported 
			print "no files found to import. Please ensure that files are in " + primer_design_files
		else:
			# return the file and primer counts
			print str(primer_count) + " primers imported from " + str(file_count) + " files"

	
	def build_insert(self,dict):
		"""
		A dictionary of values to be inserted is recieved and a SQL insert statement built and executed
		"""
		# The primer name is an autoincrementing number. A select query (ordered by primer name descending) returns the current highest number. this is a union query returning 0 should this be the first import
		query="select PrimerName from PrimerAmplicon union select '0' as PrimerName from dbo.PrimerLookup order by PrimerName desc"
		# execute and capture the query result
		result=self.fetchone(query)
		# add one to create the primer name
		highest_primer_name = result[0]+1
				
		# build sql insert statement
		self.insert_query='INSERT INTO "dbo"."PrimerAmplicon" ("EntrezGeneIDmapped", "ForwardSeq", "ReverseSeq", "ChromosomeID", "Start", "Stop","status","FTag","RTag","ScaleOfSynth","Purification","GenomeBuild", "notes","PrimerName") VALUES ('+dict['gene']+',\''+dict['FSeq']+'\',\''+dict['RSeq']+'\','+dict['chr']+','+dict['start']+','+dict['stop']+','+dict['status']+','+dict['FTag']+','+dict['RTag']+','+dict['SoS']+','+dict['purification']+','+dict['GenomeBuild']+',\''+dict['notes']+'\','+str(highest_primer_name)+')'
		# print self.insert_query
		
		# call module to execute insert query
		self.insert_query_function()

	
	def insert_query_function(self):
		"""This function executes an insert query"""
		# execute the insert query
		self.cursor.execute(self.insert_query)
		# commits to database
		self.cursor.commit()
			
	
	def fetchone(self, query):
		"""perform select query but returns a single result"""
		# perform query
		self.cursor.execute(query)
		# capture result using fetchone to return only the top result
		result = self.cursor.fetchone()
		# if there was a result 
		if result:
			# return result
			return result
		# if no result
		else:
			# print the query to aid troubleshooting
			print "unable to return result for query " + query

def main():
	# create instance of class
	a=ImportPrimer()
	# read pan493 to populate the dictionary of all gene coordinates
	a.get_gene()
	# read the look up values from moka
	a.get_moka_lookup_values()
	# parse the 
	a.parse_input()

	
if __name__ == '__main__':
	main()
