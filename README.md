A Python program to scrap films directory from the TMDB database online and create and interact with an Excel file.

Hello User,

This is a Movie database creator for an excel program. When run for the first time it will ask if not already existing the full path of the "My Movie Library.xlsx" where when running the program it will fetch that there and create a modified .xlsm excel macro file. Once open the "My Movie Library.xlsm" run the AddButtons Macro and then 

' In the Tools References dialog in Developer mode,
' scroll down and check the box for
' Microsoft Visual Basic for Applications Extensibility 5.3
' & Microsoft Forms 2.0 Object Library by inserting a form

To fetch the info you put the title in column A and the year in column B and nothing in the other columns in the .xlsx file

N.B. the User tmdbscapper_config.json in the User Documents file looks like this

{
    "excel_file_path": "F:/Richard/My Movie Library.xlsx",
    "api_key": "API_KEY"
}

Enjoy and comment me,
Richard