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

  


  
