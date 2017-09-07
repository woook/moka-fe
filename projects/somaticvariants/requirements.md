# Somatic Variants Project Requirements
## Participants
    - Product owner: Andrew Bond
    - Team: Genome Informatics
    - Stakeholders: Oncology, Genome Informatics
## Current Status
Final
## Purpose
Store somatic variants identified by NGS.
## Project Goals & Objectives
* Import variants from Ingenuity into Moka
* Record pathogenicity of variants in Moka
## Requirements
### Functional
The resulting app should:
* Allow user to select a specimen
* Allow user to associate multiple 'EK' and block numbers with a specimen 
* Allow user to import a VCF file, exported from Ingenuity Variant Analysis software, for the patient
* Imported variant annotation must include variant type, genomic coordinates and allele frequency
* Create an NGS test record for the block and attach variants to that test
* Provide a results page to allow user to record variant interpretation 
* Provide a way for the checking of the interpretation to be recorded
* The Oncology tests must be identifiable so that they can be excluded from the WES dashboard, status tracker etc. (use referral type field)
### Technical
- Adapt the existing WES system in Moka:
  - Microsoft Access front end
  - MS SQL server backend
  - Python VCF import script
 - Adhere to minimal viable product
### Usability
- Must be simple for clinical scientists to access and use
- Must be consistent with existing Moka data structures and application logic
