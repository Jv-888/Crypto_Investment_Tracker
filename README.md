Automatically fetches your Coinbase purchase confirmation emails, parses the data, and logs your crypto transactions into a secure, deduplicated .xlsx file.

I built this as an IT automation project to:
  
      Practice scripting and email parsing
      
      Learn to work with .xlsx formats using Python
      
      Automate tedious logging from real financial services
      
Features:

    Parses purchase data from Gmail using IMAP
    
    Extracts asset name, amount, price, and reference code
    
    Converts UTC email timestamps to your local timezone
    
    Prevents duplicate entries using reference codes
    
    Logs everything to Excel with formulas preserved

How To Use:

    1. Copy Repo
    2. Install requirements/Dependencies
    3. Create the .env file or declare Windows Env Variables for username/email password
    4. Run "python CryptoTrackerXLSX.py"


Results:

    1. Refer to "RunningCode.png" to see results of running the code
    2. Refer to "ExcelFile.png" to see how the data is imported 




  


  
