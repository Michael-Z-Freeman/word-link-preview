# word-link-preview
Create link previews for Word (.docx)
![preview-example Large](https://github.com/Michael-Z-Freeman/word-link-preview/assets/951566/a95f091d-54c2-459a-b5c0-ea9083f4aab7)
The preview image shows the docx in Apple Pages where it looks OK. It looks bad in Word but seeing as I don't use word I did not fix it (it's probably easy to fix in Word or in the script).
The project requires a text file "links.txt" with the actual links to create previews from.
The script assumes existence of directories "PREVIEWS/CROPPED" and "HTML".
I have not had time to write an installation guide but you will have to install various modules using pip. Playright requires installing a browser by using "bin/playwright install". Once you have installed via pip you will find "bin/playwright" in the cloned repository (if you are using a virtual env) or in whatever bin directory usually contains Python executables on your system.
