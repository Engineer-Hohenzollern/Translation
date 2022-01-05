import docx
from googletrans import Translator # try pip install googletrans==4.0.0-rc1, if there is a Module Not Found Error

d = docx.Document('thai.docx')  # to open the document containing the promotional content

translator = Translator()
i = d.paragraphs  # i return's a list of object where in each of this object contains string of the paragraphs
# elements in the Word document we can access them in str form by .text attribute
# to be more specific  d.paragraphs[list_element].text
i = len(i)  # inorder to get the number of paragraphs in word document, so we can iterate in for loop

newd = docx.Document()  # we're creating a new Word document for the translated text
for r in range(i):
    c = d.paragraphs[r].text  # each paragraphs of the Word document is  returned as a string
    print(c)
    if c is not None:  # we don't want empty lines to be translated otherwise this might cause a type error
        result = translator.translate(c)  # we translate it from thai to eng , note : since source language can be
        # auto-detected we don't need to mention src=th, by default dest= english hence we don't need to mention it
        print(result.text)
        newd.add_paragraph(result.text)  # we add the translated paragraphs to the new Word document

newd.save('newthait.docx')  # it must be saved to unique name within its file path as word would not give permission to
# replace an existing Word docx
