# AI Generated Alternative Texts for PowerPoint slide-decks using Google-VertexAI


# Dependencies:
```
pip install python-pptx
pip install tqdm
pip install requests
```

# Usage:
- Extract all images out of a .PPTX file, save them in the same folder, and names the images by the power-point-slide-number they appear in the deck.

   ```$ python3 auto_alt_text_for_pptx.py --pptx Autonomic-Module9_v3.pptx```

- Obtain the caption predictions from Google's VertexAI for each extracted image in the directory. Writes the captions with the corresponding image name in a csv file. Bearer token should be gathered from Google Cloud Console with $ gcloud auth print-access-token. Also, projectNameID field should be changed inside the code, as well as location (see getCaption() function definition.)

   ```$ python3 auto_alt_text_for_pptx.py --dir . --output record_captions.csv  --bearer <auth_session_bearer>```

- Using the csv file, it updates the Alternative Text of each matching image in-place and saves the updated PPTX.

   ```$ python3 auto_alt_text_for_pptx.py --update record_captions.csv --pptx Autonomic-Module9_v3.pptx```

# Help

```python
$ python3 auto_alt_text_for_pptx.py --help
usage: auto_alt_text_for_pptx.py [-h] [-p PPTX] [-i IMAGE] [-d DIR] [-b BEARER] [-o OUTPUT] [-u UPDATE]

Extract images from PPTX

optional arguments:
  -h, --help            show this help message and exit
  -p PPTX, --pptx PPTX  .pptx filename to extract images
  -i IMAGE, --image IMAGE
                        .png filename to get a caption from GoogleCloudVertex
  -d DIR, --dir DIR     a directory of .png images (extracted from a .pptx file) to get the caption for.
  -b BEARER, --bearer BEARER
                        GoogleCloud VertexAI auth token, can be obtained by running `$ gcloud auth print-access-token` on the cloudshell
  -o OUTPUT, --output OUTPUT
                        When using -dir or -image, we need to provide an output file to write the caption responses from GoogleCloud
  -u UPDATE, --update UPDATE
                        to update the Alternative Text for the images. Need to provide a csv file generated with --output in a prev run from --dir or --image
```

# Automate processing multiple pptx in different subfolders:

If we have multiple slides (one per module), organized as follows;
```bash
➜  auto_alttext_pptx git:(main) ✗ tree .
/Users/tekeoglua/Desktop/Research/autonomic/jhu/course/slides/Spring24
├── Module 1
│   └── Autonomic-Module1_v3.pptx
├── Module 10
│   └── Autonomic-Module10_v3.pptx
├── Module 11
│   └── Autonomic-Module11_v4.pptx
├── Module 2
│   └── Autonomic-Module2_v3.pptx
├── Module 3
│   └── Autonomic-Module3_v2.pptx
├── Module 4
│   └── Autonomic-Module4_v3.pptx
├── Module 5
│   └── Autonomic-Module5a_v3.pptx
├── Module 6
│   └── Autonomic-Module6_v3.pptx
├── Module 7
│   └── Autonomic-Module7_v3.pptx
├── Module 8
│   └── Autonomic-Module8_v2.pptx
└── Module 9
    └── Autonomic-Module9_v3.pptx

```

Make sure, there is only one PPTX file in each subdirectory.

We can run ./automate.sh to process all of them at once. Example run;

```bash
./automate.sh
Default PATH to Root Folder: /Users/tekeoglua/Desktop/Research/autonomic/jhu/course/slides/Spring24
Do you want to change the path to the modules root? (y/n)
n
Using Default PATH to Root Folder: /Users/tekeoglua/Desktop/Research/autonomic/jhu/course/slides/Spring24
/Users/tekeoglua/Desktop/Research/autonomic/jhu/course/slides/Spring24
├── Module 1
│   └── Autonomic-Module1_v3.pptx
├── Module 10
│   └── Autonomic-Module10_v3.pptx
├── Module 11
│   └── Autonomic-Module11_v4.pptx
├── Module 2
│   └── Autonomic-Module2_v3.pptx
├── Module 3
│   └── Autonomic-Module3_v2.pptx
├── Module 4
│   └── Autonomic-Module4_v3.pptx
├── Module 5
│   └── Autonomic-Module5a_v3.pptx
├── Module 6
│   └── Autonomic-Module6_v3.pptx
├── Module 7
│   └── Autonomic-Module7_v3.pptx
├── Module 8
│   └── Autonomic-Module8_v2.pptx
└── Module 9
    └── Autonomic-Module9_v3.pptx

12 directories, 11 files
===================================================
(1) List alt-text for all images in .pptx files
(2) Extract all images from .pptx files.
(3) Get auto-captions from Vertex-AI for images.
(4) Update alt-text for each image in place.
===================================================
Choose 1-4:
```