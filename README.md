This is the Krienen Data Log.  
You can use the 'Enter' key to go to the next field.  
Some fields you must type your response, others are a dropdown menu.  
It will automatically paste what is in your clipboard to the "elab url" column.  
To reset the "barcoded cell sample name" column, right-click DataLogger.app > Show Package Contents > MacOS > open sample_name_counter.json in a text editor, and paste this in while changing "next_counter" to your desired number:  
  {  
    "next_counter": **your_number_here**,  
    "date_info": {},  
    "amp_counter": {}  
}  
It should ask you where you want to save the .xlsx file if you haven't run the program before.  
