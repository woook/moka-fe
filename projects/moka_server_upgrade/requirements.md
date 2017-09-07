# Moka server upgrade
## Participants
- Product owner: Joo Wook Ahn
- Team: Genome Informatics
- Stakeholders: Viapath Genetics, Viapath IT, GSTT IT
- Other key relationships: GSTT Clinical Genetics

## Current Status
Draft

## Purpose
Transition Moka (SQL Server instance and Windows file share) from current GSTT Windows server (GSTTV-MOKA) to new GSTT Windows servers (SV-PR-GENAAP01 & SV-TE-GENAAP01)

## Project Goals & Objectives
* Install Windows Server onto SV-PR-GENAAP01 & SV-TE-GENAAP01
* Install SQL Server 2014 onto SV-PR-GENAAP01 & SV-TE-GENAAP01
* Set up Windows fileshare on SV-PR-GENAAP01 & SV-TE-GENAAP01
* Transfer SQL server database from GSTTV-MOKA to SV-PR-GENAPP01
* Transfer shared files from GSTTV-MOKA to SV-PR-GENAPP01
* Create new ODBC links between Moka frontend and new backend
* Package new ODBC links with Moka frontend for deployment to users (PCs)
* Test new Moka backend
* Test new Moka frontend
* Deploy new Moka frontend to users

## Requirements
### Functional
* All data transfered
* No loss of Moka availability for users

### Technical
* GSTTV-MOKA comprises Windows Server 2003 & SQL Server 2005
* SV-PR-GENAAP01 & SV-TE-GENAAP01 should use Windows Server 2012 & SQL Server 2014

### Usability
- SV-PR-GENAAP01 & SV-TE-GENAAP01 should be available via RDP
