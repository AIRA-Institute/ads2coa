# AIRA INSTITUTE NSF COA Generator

This script generates the NSF COA from an author-affiliations csv file.

## Installation:

```bash
pip install git+https://github.com/AIRA-Institute/ads2coa.git#egg=ads2coa
```

## Steps:

 - Get the author affiliations csv file from
   [ADS](https://ui.adsabs.harvard.edu/) and save it locally as
   `authorAffiliations.csv`.

 - Run this script (see `--help` for arguments):

   ```bash
   ads2coa
   ```

 - Edit the generated file and save it as `lastname_coa.xlsx`

   Table 4 is populated; all other tables need to be filled out manually.
