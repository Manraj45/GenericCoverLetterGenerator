# GenericCoverLetterGenerator

GenericCoverLetterGenerator is a java application that takes a generic cover letter (.docx), the hiring manager's name, the company name and the position applied to.
With this information, this application creates a new cover letter and converts it to pdf.

Additionally, environment variables must be provided for the input file path (INPUT_FILE_PATH), output file path (OUTPUT_FILE_PATH) and the name of the user (NAME).

In the generic cover letter provided by the user, the following must be added in the position where we want to generate them: \<hiring manager> and \<position>. 
These strings will be replaced by the name you provide in the app.
