input:
- directory of word documents with `###-[courses|mission|skills].docx` filenames
- spreadsheet containing information about degree programs

output:
- a list of which degree programs are suitable for upload to dedoose
    - meaning they have all three documents (courses, mission, skills)
    - what else?
- a list of which programs are not suitable and why
    - what is missing?

second level goals:
- find which documents are null (eg. "No skills.") and mark for deletion
- find way to merge documents into one structure



tasks:
- read directory of degree documents
    - make a list of degree ids that have all three
    - note which ids have three degrees but not 'docs made: Y' and vice versa


- read excel file
    - degree id
    - degree name
    - docs made (T/F)


1. load US spreadsheet to pandas
2. load non-US spreadsheet to pandas