# location of primer design files to import
primer_design_files = "\\\\gstt.local\\shared\\Genetics\\DNA LAB\\Primers\\MokaPrimerDatabase\\primer_designs"
# location of imported design files 
archived_primer_design_files = primer_design_files + "\\imported_primers\\"


#default strings to import to database (from item table)
genome_build_id = "109" 
scale_of_synth_id = "1978" 
purification_id = "1977" 

#default strings to import to database (from status table)
status_id = "1202218830" # status = designed awaiting ordering

# default tag values (from item table)
f_tag_seq_id = "1974" # sequence of the forward tag
f_tag_name_id = "1972" # name of the forward tag
tag_ordering_name_id = "1973" # key of the tag ordering name
tag_ordering_name_text = "T1_01" # name of the ordering name
r_tag_seq_id = "1976" # sequence of the reverse tag
r_tag_name_id = "1975" # name of the reverse tag
