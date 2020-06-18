'''
v1.9 - AB 2019/09/11

Usage:
    Parses an Ingenuity VCF output file and inserts variants and annotations into Moka NGSVariant and NGSVariantAnnotations tables.
    Called from the WES results form [s0901_WESTestResults] or Oncology results form [s0902_OncologyTestResults] in Moka
    Takes the following 3 arguments: comma separated string of VCF file paths, internal patient ID, NGS test ID

Requirements:
    ODBC Connection to Moka
    Python 2.7
    pyodbc module
    pyVCF module
'''
import sys
import os
import time
import math
import vcf
import pyodbc
import gzip
from ConfigParser import ConfigParser

class MokaVCF(object):
    """Parses Ingenuity VCF output file and enters each variant into Moka NGSVariants Table"""
    def __init__(self, moka_server, moka_db_name, patID, ngsTestID, DNA):
        # THE LIST BELOW CONTAINS EACH ID FROM THE INGENUITY OUTPUT VCF INFO FIELD THAT WILL BE ADDED TO DATABASE.
        # TO INCLUDE ADDITIONAL ANNOTATIONS, ADD THE CORRESPONDING INFO FIELD ID TO THE LIST BELOW AND ENSURE THERE
        # IS A FIELD WITH THE SAME NAME IN THE NGSVariantAnnotations TABLE IN MOKA
        self.fields = ('GENE_SYMBOL',
                      'TRANSCRIPT_ID',
                      'HGVS_TRANSCRIPT',
                      'HGVS_PROTEIN',
                      'GENE_REGION',
                      'TRANSLATION_IMPACT',
                      'SIFT_FUNCTION',
                      'POLYPHEN_FUNCTION',
                      'ING_CLASSIFICATION',
                      'CADD',
                      'BLOSUM50',
                      'PHYLOP',
                      'ING_FINDINGS',
                      'ING_PHENOTYPE',
                      'DBSNP',
                      'AFC_AF',
                      'AFC_AF_EAS',
                      'AFC_AF_SAS',
                      'AFC_AF_AFR',
                      'AFC_AF_EUR',
                      'AFC_AF_LAT',
                      'AF_1000G',
                      'ESP_AF',
                      'ESP_AF_AA',
                      'ESP_AF_EA',
                      'END_pos',
                      'SVLEN',
                      'SVTYPE')
        self.vcfPaths = {}
        self.cnxn = pyodbc.connect('DRIVER={{SQL Server}}; SERVER={moka_server}; DATABASE={moka_db_name};'.format(
                    moka_server=moka_server,
                    moka_db_name=moka_db_name
                ),
            autocommit=True
            )
        self.cursor = self.cnxn.cursor()
        self.prevVars = set([])
        self.mokaChr = {}
        self.mokaHGNC = {}
        self.vars = {}
        self.patID = patID
        self.ngsTestID = ngsTestID
        self.DNA = DNA
        self.datetime = time.strftime("%Y%m%d %H:%M:%S %p") # date/time in format: yyyymmdd hh:mm:ss AM/PM

    def makeVCFdict(self, vcfPathLst):
        for vcf in vcfPathLst:
            # pass to the function which will unzip if required and return the new file path
            vcf = self.unzip_vcf(vcf)

            #if oncology sample
            if os.path.basename(vcf).startswith("ONC"):
                self.vcfPaths[(0, "Oncology")] = vcf
            #else if WES sample split to get the panel name (after a dash) but without any extension (before the dot)
            else:
                # remove the file extension from the panel name
                panelName = os.path.basename(vcf).split("-")[1].split(".")[0]
                if panelName.lower() == "primarypanel":
                    self.vcfPaths[(1, "Primary")] = vcf
                elif panelName.lower() == "secondarypanel":
                    self.vcfPaths[(2, "Secondary")] = vcf
                elif panelName.lower() == "phenotype":
                    self.vcfPaths[(3, "Phenotype")] = vcf
                else:
                    self.vcfPaths[(4, panelName)] = vcf

    def unzip_vcf(self,vcfpath):
        # If the vcf is gzipped unzip and write to an uncompressed file. Return the filepath
        # if it's not gzipped
        if not vcfpath.endswith(".gz"):
            # return the filepath without modification
            return vcfpath
        # if it is gzipped
        else:
            # create the name of the unzipped file to be created by removing .gz extension
            new_vcfpath = vcfpath.replace('.gz','')
            # create and open the unzipped file
            with open(new_vcfpath,'w') as vcfunzipped_file:
                # open the gzipped vcf using gzip
                 with gzip.open(vcfpath,'r') as vcf_compressed:
                     # write lines from the compressed file to the uncompressed file
                     vcfunzipped_file.writelines(vcf_compressed.readlines())
            # return the file path
            return new_vcfpath

    def lookupPrevVars(self):
        # Find details of any variants already imported into Moka for this test and add to exclusion list to prevent duplication.
        sqlPrevVars = "SELECT NGSVariant.ChrID, NGSVariant.Position_hg19, NGSVariant.ref, NGSVariant.alt FROM NGSVariant WHERE NGSVariant.NGSTestID = %s" % (self.ngsTestID)
        mokaPrevVars = self.cursor.execute(sqlPrevVars).fetchall()
        for row in mokaPrevVars:
            self.prevVars.add((str(row.ChrID), str(row.Position_hg19), "'{}'".format(str(row.ref)), "'{}'".format(str(row.alt)))) # "'{}'".format() wraps string in single quotes. See below.

    def lookupChr(self):
        # Create a chromosome ID lookup dictionary from Moka Chromosome table
        sqlGetChr = "SELECT ChrID, Chr FROM Chromosome;"
        mokaChrAll = self.cursor.execute(sqlGetChr).fetchall()
        for row in mokaChrAll:
            self.mokaChr[row.Chr] = row.ChrID

    def lookupHGNCID(self):
        # Create an HGNCID lookup dictionary from Moka genesHGNC_current table
        allVCFGenes = set([]) # using set instead of list to prevent gene symbols being added multiple times
        for vcfFile in self.vcfPaths.itervalues(): # for each vcf...
            vcfReader = vcf.Reader(open(vcfFile, 'r')) # read vcf
            for row in vcfReader:
                # Add gene symbol to set. Will error if no gene symbol present, so use try/except to catch this.
                try:
                    allVCFGenes.add(row.INFO['GENE_SYMBOL'][0])
                except KeyError:
                    pass
            vcfReader = None # Releases file (pyvcf reader object has no .close() method)
        if allVCFGenes: # Check the list contains at least one gene symbol
            # Retrieve all HGNCIDs from Moka for the gene symbols in set.
            sqlGetHGNC = "SELECT ApprovedSymbol, HGNCID FROM GenesHGNC_current WHERE ApprovedSymbol IN ('%s')" % ("', '".join(allVCFGenes))
            mokaHGNCAll =  self.cursor.execute(sqlGetHGNC).fetchall()
            # Add each {gene symbol: HGNCID} pair to the self.mokaHGNC dictionary
            for row in mokaHGNCAll:
                self.mokaHGNC[row.ApprovedSymbol] = str(row.HGNCID)

    def getVars(self):
        # Extracts variants and selected annotations from VCF.
        # Builds strings of variant and annotation data for SQL INSERT statements.
        # Each variant string is stored as a dictionary key. A nested list of annotation strings (one for each affected transcript) is stored as the value.
        for panel in sorted(self.vcfPaths.keys()):
        # for each panel...
            # Add any variants already in self.vars (from prev panels) to exclusion list to prevent duplicates
            for variant in self.vars.keys():
                self.prevVars.add(variant[:4])
            # Create vcfReader object
            vcfReader = vcf.Reader(open(self.vcfPaths[panel], 'r'))
            for row in vcfReader:
                #######
                # EXTRACT DATA FOR NGSVariant TABLE...
                #######
                chromosome = row.CHROM[3:] # Removes 'chr' from beginning of string
                mokaChrID = str(self.mokaChr[chromosome]) # Convert to Moka chromosome ID using self.mokaChr dictionary **store numbers as strings so that they can be added to SQL statement
                position = str(row.POS) # genomic position
                ref = "'{}'".format(row.REF) # ref sequence **"'{}'".format used to wrap string in single quotes so that it will be recognised as a string in SQL statement.
                alt = "'{}'".format(str(row.ALT[0])) # alt sequence
                # Only continue if variant is not already associated with this test in Moka.
                if (mokaChrID, position, ref, alt) not in self.prevVars:
                    gt = "'{}'".format(row.genotype(self.DNA)['GT']) # Genotype
                    #If read depth is present, capture it, otherwise set to 'Null'
                    if row.genotype(self.DNA)['DP'] is not None:
                        rd = str(row.genotype(self.DNA)['DP']) # Read depth
                    else:
                        rd = 'Null'
                    #If QUAL is present, capture it, otherwise set to 'Null'
                    if row.QUAL is not None:
                        cq = str(row.QUAL) # Call quality
                    else:
                        cq = 'Null'
                    #Capture the allele fraction
                    af = row.genotype(self.DNA)['ING_AF'] # Ingenuity inferred allele fraction (percentage).
                    #Sometimes allele fraction has an NaN value rather than None, so check for this and record as 'Null' if present.
                    if not af or math.isnan(af):
                        af = 'Null' # Adds Null value to SQL statement
                    #If allele depth for ref allele is present, capture it, otherwise set to 'Null'
                    if row.genotype(self.DNA)['AD'][0] is not None:
                        ref_ad = row.genotype(self.DNA)['AD'][0] # Reference Allele Depth
                    else:
                        ref_ad = 'Null'
                    #If allele depth for alt allele is present, capture it, otherwise set to 'Null'
                    if row.genotype(self.DNA)['AD'][1] is not None:
                        alt_ad = row.genotype(self.DNA)['AD'][1] # Alt Allele Depth
                    else:
                        alt_ad = 'Null'
                    #If genotype quality is present, capture it, otherwise set to 'Null'
                    if row.genotype(self.DNA)['GQ'] is not None:
                        gq = row.genotype(self.DNA)['GQ'] # Genotype quality
                    else:
                        gq = 'Null'
                    # Stores each variant as a string that can be used in VALUES section of SQL insert statement (see below).
                    varCurrent = (mokaChrID, position, ref, alt, self.ngsTestID, self.patID, "'{}'".format(self.datetime), str(panel[0]), "'{}'".format(panel[1]), gt, rd, cq, str(af), str(ref_ad), str(alt_ad), str(gq))
                    #varCurrent = (mokaChrID, position, ref, alt, self.ngsTestID, self.patID, "#"+self.datetime+"#", str(panel[0]), "'{}'".format(panel[1]), gt, rd, cq, str(af), str(ref_ad), str(alt_ad), str(gq))
                    ######
                    #EXTRACT DATA FOR NGSVariantAnnotations TABLE...
                    ######
                    # Lookup HGNCID in self.mokaHGNC dictionary, add HGNCID to start of 'annotations' string for SQL INSERT statement.
                    annot = []
                    try:
                        annot.append("'{}'".format(self.mokaHGNC[row.INFO['GENE_SYMBOL'][0]]))
                    except KeyError:
                        annot.append("Null") # If no HGNCID found, a Null value will be added.
                    if row.genotype(self.DNA)['ING_CH']:
                        annot.append("'{}'".format(row.genotype(self.DNA)['ING_CH'])) # Gets the inferred compound heterozygosity from VCF sample field
                    else:
                        annot.append("Null")
                    if row.genotype(self.DNA)['ING_IA']:
                        annot.append("'{}'".format(row.genotype(self.DNA)['ING_IA'])) # Gets the inferred activity (e.g. gain, loss, normal) from VCF sample field
                    else:
                        annot.append("Null")
                    # Loop through VCF annotation fields and add to string to be used in SQL query. If field is not in VCF, add a NULL value.
                    for i in range(len(self.fields)):
                        vcfTag = self.fields[i]
                        #The 1000G_AF and END vcf tags are not valid field names in MS-SQL (can't start with number and END is reserved word).
                        #Therefore for these fields need to convert from Moka field name to VCF tag
                        if vcfTag == 'AF_1000G':
                            vcfTag = '1000G_AF'
                        elif vcfTag == 'END_pos':
                            vcfTag = 'END'
                        #Get value for the field, if it's not found, insert a Null
                        try:
                            vcfVal = row.INFO[vcfTag]
                        except KeyError:
                            annot.append("Null")
                        else:
                            # CADD can be a single value in list, or a min and max value in list if CADD score less than 10.
                            # Take the last value in the list, so the actual CADD score or max value is taken.
                            if vcfTag == 'CADD':
                                vcfVal = vcfVal[-1]
                            #If the number of items for a field is not specified in the VCF header, pyvcf returns a list.
                            #Every field should only have one value, so if a list is returned just take the first element
                            if type(vcfVal) == list:
                                vcfVal = vcfVal[0]
                            #If the value is a number, need to convert to string for SQL statement.
                            if type(vcfVal) == int or type(vcfVal) == float:
                                if math.isnan(vcfVal): # If value is NaN, add a Null value
                                    annot.append("Null")
                                else:
                                    annot.append(str(vcfVal)) # Need to convert number to string for adding to SQL statement
                            #If the value is a string, need to enclose in single quotes for SQL statement.
                            else:
                                annot.append("'{}'".format(vcfVal.replace("'", "''")))
                    # Each variant string is stored as a dictionary key.
                    # The dictionary value is a nested list of annotations for each assosciated transcript.
                    # This list is added to each time a new row for that variant is encountered in the VCF.
                    self.vars[varCurrent] = self.vars.setdefault(varCurrent, []) + [annot]
            vcfReader = None # Releases file (pyvcf reader object has no .close() method)

    def insertMoka(self):
        # Loops through variant dictionary and inserts into Moka
        sqlIns = "INSERT INTO NGSVariant (Gene, ChrID, Position_hg19, ref, alt, NGSTestID, InternalPatientID, DateAdded, PanelType, PanelTypeName, genotype, ReadDepth, CallQuality, AlleleFraction, RefAlleleDepth, AltAlleleDepth, GenotypeQuality) VALUES (%s, %s)"
        for var in sorted(self.vars.keys()):
            #Retrieve gene symbol(s) associated with each variant so they can be added to NGSVariants table. Separate multiple genes with semi-colon (;).
            genes = ";".join(set([annot[3][1:-1] for annot in self.vars[var] if annot[3] != "Null"]))
            if genes != "":
                genes = "'{}'".format(genes) #Surround in single quotes for SQL
            else:
                genes = "Null" #Insert Null value if vaiant not associated with any genes
            sqlRowIns = sqlIns % (genes, ", ".join(var))
            self.cursor.execute(sqlRowIns) # Insert variant
            ngsVariantID = self.cursor.execute("SELECT @@IDENTITY").fetchone()[0] #Get last inserted ID
            for entry in self.vars[var]:
                sqlInsAnnot = "INSERT INTO NGSVariantAnnotations (NGSVariantID, HGNCID, InferredCompHet, InferredActivity, %s) VALUES (%s, %s)" % (", ".join(self.fields), str(ngsVariantID), ", ".join(entry))
                self.cursor.execute(sqlInsAnnot) # Insert annotations for each variant transcript

if __name__ == "__main__":
    vcfPathLst = sys.argv[1].split(",")
    patID = sys.argv[2]
    ngsTestID = sys.argv[3]
    DNA = sys.argv[4]
    # Read config file
    config = ConfigParser()
    config.read(os.path.join(os.path.dirname(os.path.realpath(__file__)), "config.ini"))
    # Insert variants to Moka
    mv = MokaVCF(config.get("MOKA", "SERVER"), config.get("MOKA", "DATABASE"), patID, ngsTestID, DNA)
    mv.makeVCFdict(vcfPathLst)
    mv.lookupPrevVars()
    mv.lookupChr() # Retrieve IDs from Moka Chromosome table
    mv.lookupHGNCID() # Retrieve HGNCIDs for genes in VCF from Moka GenesHGNC_current table
    mv.getVars() # Extract variants and annotations from VCF
    mv.insertMoka() # Insert variant records and annotations into Moka NGSVariant and NGSVariantAnnotations tables.
