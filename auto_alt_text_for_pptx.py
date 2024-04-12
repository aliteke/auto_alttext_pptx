"""
    Extracts images from a PPTX file, writes them into the same directory.
    Request GoogleVertexAI Image Captioning Service to generate a image caption for each image extracted from PPTX file in the directory.
    Author: ali.tekeoglu@jhu.edu
"""
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.shapes.base import BaseShape
from tqdm import tqdm
import argparse
import os
import sys
import base64
import requests
import json
import time
import csv

# TODO: Take a look at this repository: https://github.com/waltervanheuven/auto-alt-text
# TODO: Tutorial https://cloud.google.com/vertex-ai/generative-ai/docs/model-reference/image-captioning

def getCaptionsForDir(dirName, bearer, outputFile):
    for image in tqdm([ y for y in os.listdir(".") if "png" in y.lower()]):
        print("[+] Sleeping 2 seconds....")
        time.sleep(2)
        getCaption(image, bearer, outputFile)


def getCaption(imageName, bearer, outputFile):
    """ we got the authorizaion bearer through Google-Cloud-Console command:
        aliteke@cloudshell:~/images (rising-analogy-272214)$ gcloud auth print-access-token
    """
    projectNameID = "rising-analogy-272214"
    location = "us-central1"

    headers = {"Authorization": f"Bearer {bearer}", "Content-Type": "application/json; charset=utf-8"}
    
    response_count = 1
    language_code = "en"
    b64_image = getB64forImage(imageName)
    data = { "instances": [ { "image": { "bytesBase64Encoded": b64_image } } ], "parameters": { "sampleCount": response_count, "language": language_code } }

    url = f"https://{location}-aiplatform.googleapis.com/v1/projects/{projectNameID}/locations/{location}/publishers/google/models/imagetext:predict"
   
    response = requests.post(url, json=data, headers=headers)
    print(f"[+] Request: url: {url}, headers: {headers}")
    print(f"[+] Status: {response.status_code}, Caption for {imageName}: {response.json()}")
    if response.status_code == 429:
        time.sleep(1)
        print("[+] Recursively calling itself again in a 15 seconds...")
        time.sleep(15)
        getCaption(imageName, bearer, outputFile)
    else:
        #result=json.dumps({imageName : response.json()})
        with open(outputFile, 'a') as f:
            f.write(f"{imageName},\"{response.json()['predictions'][0]}\"\n")

def getB64forImage(imageName):
    with open(f"{imageName}", "rb") as image_file:
        data = base64.b64encode(image_file.read()).decode('utf-8')
    return data


def writeImageToFile(pageNum, shapeNum, imageBytes, extension):
    fname = f"image_pg{pageNum}_idx{shapeNum}.{extension}"
    print(f"[+] Writing image to file: {fname}") 
    
    if not os.path.exists(fname):
        with open(fname, "wb") as f:
            f.write(imageBytes)
    else:
        print("[!] File already exists, skipping")


def shape_alt_text(shape: BaseShape) -> str:
    """Alt-text defined in shape's `descr` attribute, or "" if not present."""
    return shape._element._nvXxPr.cNvPr.attrib.get("descr", "")


def shape_alt_text_set(shape: BaseShape, alt_text: str):
    """Alt-text defined in shape's `descr` attribute, or "" if not present."""
    shape._element._nvXxPr.cNvPr.descr = alt_text
    print(f'[+] AFTER INSERT to DICTIONARY: {shape._element._nvXxPr.cNvPr.items()}')


def getAltTextFromPredsFile(genCaptionsFname, imageFname):
    preds=list()
    with open(genCaptionsFname, 'r') as file:
        lines = csv.reader(file)
        for l in lines:
            preds.append(l)

    for item in preds:
        fname =  item[0]
        pred = item[1]
        if imageFname == fname:
            return pred


def updateImageAlttexts(pptxFname, genCaptionsFname):
    print(f"[+] Updating image alt text for ppt file: {pptxFname} with captions from: {genCaptionsFname}")
    prs = Presentation(pptxFname)
    for slideNum, slide in enumerate(prs.slides):
        for shapeNum, shape in enumerate([s for s in slide.shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE]):
            imageFname = f"image_pg{slideNum}_idx{shapeNum}.{shape.image.ext}"
            alt_text = getAltTextFromPredsFile(genCaptionsFname, imageFname)
            print(f"[+] Setting AltText: {alt_text} for {imageFname}")
            shape_alt_text_set(shape, alt_text)

    prs.save(pptxFname)


def extractImagesFromPPTX(fname):
    prs = Presentation(fname)
    for slideNum, slide in enumerate(prs.slides):
        for shapeNum, shape in enumerate([s for s in slide.shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE]):
            #print(f"Slide#: {slideNum}, Shape# {shapeNum}, SlideID: {slide.slide_id}, Shape_type: {shape.shape_type}")
            image_bytes = shape.image.blob    # image's file contents
            #shape_alt_text_set(shape, "Test Alternative Text ALI...")
            writeImageToFile(slideNum, shapeNum, image_bytes, shape.image.ext)

    # TODO: Save the changes to alt-text
    #prs.save(fname)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Extract images from PPTX")
    parser.add_argument('-p', '--pptx', required=False, help=".pptx filename to extract images")
    parser.add_argument('-i', '--image', required=False, help=".png filename to get a caption from GoogleCloudVertex")
    parser.add_argument('-d', '--dir', required=False, help="a directory of .png images (extracted from a .pptx file) to get the caption for.")
    parser.add_argument('-b', '--bearer', required=False, help="GoogleCloud VertexAI auth token, can be obtained by running `$ gcloud auth print-access-token` on the cloudshell")
    parser.add_argument('-o', '--output', required=False, help="When using -dir or -image, we need to provide an output file to write the caption responses from GoogleCloud")
    parser.add_argument('-u', '--update', required=False, help="to update the Alternative Text for the images. Need to provide a csv file generated with --output in a prev run from --dir or --image")
    args = parser.parse_args()

    if args.pptx and args.image and args.dir:
        sys.exit(f"[!] Only provide one flag; either PNG image, PPTX file, or a directory")

    if args.pptx and not args.update:
        if not os.path.exists(f"{args.pptx}"):
            sys.exit(f"[!] {args.pptx} doesn't exist")
        extractImagesFromPPTX(args.pptx)

    elif args.image and args.bearer and args.output:
        if not os.path.exists(f"{args.image}"):
            sys.exit(f"[!] {args.image} doesn't exist")
        getCaption(args.image, args.bearer, args.output)

    elif args.dir and args.bearer and args.output:
        if not os.path.exists(f"{args.dir}"):
            sys.exit(f"[!] {args.dir} doesn't exist")
        getCaptionsForDir(args.dir, args.bearer, args.output)

    elif args.pptx and args.update:
        if not os.path.exists(f"{args.update}") or not os.path.exists(f'{args.pptx}'):
            sys.exit(f"[!] {args.update} or {args.pptx} doesn't exist")
        updateImageAlttexts(args.pptx, args.update)

    else:
        print(f"[!] invalid...")
        parser.print_help()
