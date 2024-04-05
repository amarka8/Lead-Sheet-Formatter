A Google Sheets Add On that formats lead sheets using Google Apps Script. Coded by Amar Kanakamedala and Mohid Saeed.
Script Contains 4 Main Functions:
  1. removeDups() removes duplicates with the filter, indexOf, and JSON.stringify array methods
  2. emailEditor() moves rows with no email into a separate sheet titled "Linkedin Only" with the filter array method
  3. removeEmptyNames() deletes rows with empty first names
  4. modifyCompanyNames() highlights rows with company names greater than x characters (21 by default) and removes unneccesary company endings (LLC, .co, .ai, etc.) 
