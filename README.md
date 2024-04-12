# AI Generated Alternative Texts for PowerPoint slide-decks using Google-VertexAI


# Dependencies:
```
pip install python-pptx
pip install tqdm
pip install requests
```

# Usage:
Extract all images out of a .PPTX file, save them in the same folder, and names the images by the power-point-slide-number they appear in the deck.

```$ python3 extract_images_from_pptx.py --pptx Autonomic-Module9_v3.pptx```

Obtain the caption predictions from Google's VertexAI for each extracted image in the directory. Writes the captions with the corresponding image name in a csv file. Bearer token should be gathered from Google Cloud Console with $ gcloud auth print-access-token. Also, projectNameID field should be changed inside the code, as well as location (see getCaption() function definition.)

```$ python3 extract_images_from_pptx.py -d . -o record_captions.csv  -b <auth_session_bearer>```

Using the csv file, it updates the Alternative Text of each matching image in-place and saves the updated PPTX.

```$ python3 extract_images_from_pptx.py --update record_captions.csv --pptx Autonomic-Module9_v3.pptx```
