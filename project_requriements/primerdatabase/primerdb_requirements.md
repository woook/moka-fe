
# Primer database
## Participants
- Product owner: Viapath Genome Informatics
- Team: Viapath Genome Informatics
- Stakeholders: Viapath Genome Informatics, Viapath Genetics laboratories
- Other key relationships: 

## Current Status
Draft 

## Purpose
A database which records primers designed by the laboratory using the PHIGIS primer design tool. 
Primers can be displayed in a format compatible with the suppliers ordering template.
Once primers have been ordered they are tested and the results recorded, allowing the library of 'approved/validated' primers to be searched to identify if primers exist for the region of interest or if new primers must be designed.

## Project Goals & Objectives
Save time for Clinical Scientists and technical staff when:
* Designing and ordering primers
* Querying existing primer library
* Tracking status of newly ordered primers
* Track primers which failed testing
* Recording location of primer in freezer

## Requirements
### Functional
* Take the output from the PHIGIS primer design tool and inserts to database
* Create a unique numerical primer name
* Output/display primer information to aid primer ordering.
* Record the result of testing of primer design.
* Query primer catalog to identify primers which overlap a genomic coordinate.

### Technical
* Incorporated into Moka
* Usable by all users on Trust windows computers.

### Usability
* Easy to use
* Must be consistent with existing Moka data structures and application logic
