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
import ImportPrimerconfig as config
# import database details from connection file
import ImportPrimerDBConnections as connection

class ImportPrimer():
	def __init__(self):
		# using pyodbc specify database connection details
		self.cnxn = pyodbc.connect(connection.mokadata_connection) # use the mokadata database connection details specified in the connection file
		
		# create a cursor to connect to database 
		self.cursor = self.cnxn.cursor()
		
	def parse_input(self):
		"""
		Primer designs are read from files within a specific folder.
		This function captures all the required information needed to insert each primer into the database and passes this information to a function which executes the insert query, one primer at a time.
		"""
		# count files parsed and primers imported to be reported at the end
		file_count = 0
		primer_count = 0
		
		# for each file in the directory of primer designs
		for file in os.listdir(config.primer_design_files):
			# look for all the text files
			if file.endswith('.txt'):
				# add to the file count
				file_count += 1
				# open file as read only
				with open(config.primer_design_files + "//" + file , 'r') as primer_designs:
					# loop through line by line, using enumerate so the first line can be skipped
					for line_number, line in enumerate(primer_designs):
						# skipping first line, and making sure the line isn't empty:
						if line_number > 0 and len(line) >3:
							# split the line on tabs into a list 
							splitline = line.split("\t")
							# take the first column in the file, which is genomic coordinates in the format chr16_70534805-70535059
							# everything before the underscore is the chromosome
							chr = splitline[0].split("_")[0]
							# everything after the underscore, but before the dash is the start position
							primer_start = splitline[0].split("_")[1].split("-")[0]
							# everything after the underscore, and after the dash is the stop position
							primer_stop = splitline[0].split("_")[1].split("-")[1]
							
							# second column in the file is the variant id  - this is to go into notes column
							variant_id = splitline[1]
							
							# capture forward and reverse primer sequences from column 7 and 8
							forward_sequence = splitline[6]
							reverse_sequence = splitline[7]
							
							# take the chromosome, and remove 'chr' incase it was present.
							chromosome_lookup = chr.replace('chr','')
							# use the chromosome value without the chr to query the chromosome look up table to return key for the chromosome
							query = "select ChrID from Chromosome where Chr='" + chromosome_lookup + "'"
							# pass query to function to query the database
							chromosome_lookup = self.fetchone(query)[0]
							
							# create a dict of all variables to be inserted to the database and pass to self.build_insert module
							primer_values_to_insert_dict = {"chr" : str(chromosome_lookup) , "start" : str(primer_start) , "stop" : str(primer_stop) ,
								"FSeq" : forward_sequence , "RSeq" : reverse_sequence , "f_tag_name" : config.f_tag_name_id , "f_tag_seq" : config.f_tag_seq_id , "f_tag_order_name" : config.tag_ordering_name_id,
								"r_tag_name" : config.r_tag_name_id , "r_tag_seq" : config.r_tag_seq_id , "r_tag_order_name" : config.tag_ordering_name_id , "status" : config.status_id,
								"purification" : config.purification_id , "SoS" : config.scale_of_synth_id , "GenomeBuild" : config.genome_build_id , "notes" : variant_id}						
							
							# pass dictionary to insert query
							self.build_insert(primer_values_to_insert_dict)
							# add to count of inserted primers
							primer_count += 1
				
				# move the file design file into an archive folder
				os.rename(config.primer_design_files + "//" + file, config.archived_primer_design_files + "//" + file)
				# NB errors may occur if a file with the same name already exists in the archive - but this shouldn't happen routinely
		
		# when all files  have been read report progress (this should be displayed in a moka message box)
		if file_count == 0:
			# if no files were found to be imported 
			print "no files found to import. Please ensure that files are in " + config.primer_design_files
		else:
			# return the file and primer counts
			print str(primer_count) + " primers imported from " + str(file_count) + " files"

	
	def build_insert(self,dict):
		"""
		A dictionary of values to be inserted is recieved and a SQL insert statement built and executed.
		This inserts one primer at a time
		"""
		# The primer name is an autoincrementing number. A select query (ordered by primer name descending) returns the current highest number. 
		# if this table is empty the query will not return anything so a union query is used to return a value of 0 (sql requires a different table to the first half of the query)
		query = "select PrimerName from PrimerAmplicon union select '0' as PrimerName from dbo.PrimerLookup order by PrimerName desc"
		
		# execute and capture the query result
		result = self.fetchone(query)
		
		# add one to create the primer name
		highest_primer_name = str(result[0] + 1)
		
		# build the primer order names, eg 1F_T1_01, using the highest primer name, F or R and the textdefined in the config file
		FPrimer_order_name = highest_primer_name + "F_" + config.tag_ordering_name_text
		RPrimer_order_name = highest_primer_name + "R_" + config.tag_ordering_name_text
				
		# build sql insert statement
		insert_query = 'insert into dbo.PrimerAmplicon ("ForwardSeq","ReverseSeq","ChromosomeID","Start19","Stop19","Status","FTagName","FTagSequence","FTagOrderingName","RTagName",\
			"RTagSequence","RTagOrderingName","Notes","ScaleOfSynth","Purification","PrimerName","FPrimerOrderingName","RPrimerOrderingName") VALUES (\'' + dict['FSeq'] + '\',\'' \
			+ dict['RSeq'] + '\',' + dict['chr'] + ',' + dict['start'] + ',' + dict['stop'] + ',' + dict['status'] + ',' + dict['f_tag_name'] + ',' + dict['f_tag_seq'] + ',' \
			+ dict['f_tag_order_name'] + ',' + dict['r_tag_name'] + ',' + dict['r_tag_seq'] + ',' + dict['r_tag_order_name'] + ',\'' + dict['notes'] + '\',' + dict['SoS'] + ',' \
			+ dict['purification'] + ',' + highest_primer_name + ',\'' + FPrimer_order_name + '\',\'' + RPrimer_order_name + '\')'
		#print self.insert_query
		
		# call module to execute insert query
		self.insert_query_function(insert_query)

	
	def insert_query_function(self,insert_query):
		"""This function executes an insert query"""
		# execute the insert query
		self.cursor.execute(insert_query)
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
	a = ImportPrimer()
	# parse the primer design files to build and execute insert statements
	a.parse_input()

	
if __name__ == '__main__':
	main()
