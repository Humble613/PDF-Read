# Description:
# This script defines the Document and Page classes to streamline the flow of information through the script.

import os
import numpy as np
import cv2
from io import BytesIO
import re, json, fitz
from post_processing import add_name_city
import datetime
import pytesseract
from paddleocr import PaddleOCR
from dotenv import load_dotenv
import argparse

load_dotenv()

Pocr = PaddleOCR(use_angle_cls=True)

tesseract_Path = os.getenv("TESSERACT_PATH")

# Global variables
cell_height = 160
consignee_width = 435
pro_num_width = 320

def split_pages(digit_doc, past_k, next_k):
    '''
    1. Splits the input pdf into pages
    2. Writes a temporary image for each page to a byte buffer
    3. Loads the image as a numpy array using cv2.imread()
    4. Appends the page image/array to self.pages

    Notes:
    PyMuPDF's get_pixmap() has a default output of 96dpi, while the desired
    resolution is 300dpi, hence the zoom factor of 300/96 = 3.125 ~ 3.
    '''
    print("Splitting PDF into pages")
    pages = []
    try:
        zoom_factor = 3
        for i in range(past_k, next_k):
            # Load page and get pixmap
            page = digit_doc.load_page(i)
            pixmap = page.get_pixmap(matrix=fitz.Matrix(zoom_factor, zoom_factor))

            # Initialize bytes buffer and write PNG image to buffer
            buffer = BytesIO()
            buffer.write(pixmap.tobytes())
            buffer.seek(0)

            # Load image from buffer as array, append to pages, close buffer
            img_array = np.asarray(bytearray(buffer.read()), dtype=np.uint8)
            page_img = cv2.imdecode(img_array, 1)
            pages.append(page_img)
            # cv2.imwrite(f'imgs/{i}.jpg', page_img)
            buffer.close()
    except:
        pass
    if len(pages) == 0:
        val = "01"
    else:
        val = pages
    return val
def subset(set, lim, loc):
        '''
        set: one or multi list or array, lim: size, loc:location(small, medi, large)
        This function reconstructs set according to size of lim in location of loc.
        '''
        cnt, len_set = 0, len(set)        
        v_coor_y1, index_ = [], []
        pop = []
        for i in range(len_set):
            if i < len_set-1:
                try:
                    condition = set[i+1][0] - set[i][0]
                except:
                    condition = set[i+1] - set[i]
                if condition < lim:
                    cnt = cnt + 1
                    pop.append(set[i])
                else:
                    cnt = cnt + 1
                    pop.append(set[i])
                    pop = np.asarray(pop)
                    try:
                        if loc == "small": v_coor_y1.append([min(pop[:, 0]), min(pop[:, 1]), max(pop[:, 2])])
                        elif loc == "medi": v_coor_y1.append([int(np.median(pop[:, 0])), min(pop[:, 1]), max(pop[:, 2])])
                        else: v_coor_y1.append([max(pop[:, 0]), min(pop[:, 1]), max(pop[:, 2])])
                    except:
                        if loc == "small": v_coor_y1.append(min(pop))
                        elif loc == "medi": v_coor_y1.append(int(np.median(pop)))
                        else: v_coor_y1.append(max(pop))  
                    index_.append(cnt)
                    cnt = 0
                    pop = []
            else:
                cnt += 1
                pop.append(set[i])
                pop = np.asarray(pop)
                try:
                    if loc == "small": v_coor_y1.append([min(pop[:, 0]), min(pop[:, 1]), max(pop[:, 2])])
                    elif loc == "medi": v_coor_y1.append([int(np.median(pop[:, 0])), min(pop[:, 1]), max(pop[:, 2])])
                    else: v_coor_y1.append([max(pop[:, 0]), min(pop[:, 1]), max(pop[:, 2])])
                except:
                    if loc == "small": v_coor_y1.append(min(pop))
                    elif loc == "medi": v_coor_y1.append(int(np.median(pop)))
                    else: v_coor_y1.append(max(pop))                    
                index_.append(cnt)

        return v_coor_y1, index_ 

def getTextAndCoorFromPaddle(img, lang='eng'):
    strp_chars = "|^#;$`'-_\/*‘ \n"
    Boxes = Pocr.ocr(img,rec=False)[0]
    image = img.copy()
    newBoxes, Cy_list, Cx_list = [], [], []
    for box in Boxes:
        x0, x1 = int(min(np.array(box)[:, 0])), int(max(np.array(box)[:, 0]))
        y0, y1 = int(min(np.array(box)[:, 1])), int(max(np.array(box)[:, 1]))
        cx, cy = int(x1/2+x0/2), int(y1/2+y0/2)
        # image = cv2.rectangle(image, (x0, y0), (x1, y1), (0, 0, 255), 2)
        newBoxes.append([x0, y0, x1, y1])
        Cy_list.append(cy)
        Cx_list.append(cx)
    
    CyCpy = Cy_list.copy()
    CyUnique, _ = subset(np.sort(CyCpy), 15, 'medi')                

    Cy_list = [CyUnique[np.argmin(abs(np.array(CyUnique)-v))] for v in Cy_list]
    Cy_list, Cx_list, newBoxes = zip(*sorted(zip(Cy_list, Cx_list, newBoxes)))
    all_text = []
    for k, box in enumerate(newBoxes):
        pytesseract.pytesseract.tesseract_cmd = tesseract_Path
        text = pytesseract.image_to_string(img[box[1]:box[3], box[0]:box[2]], lang=lang, config='--psm 6')
        text = text.strip(strp_chars)
        all_text.append(text)
        # image = cv2.rectangle(image, (box[0], box[1]), (box[2], box[3]), (0, 255, 0), 2)
        # color = (0, 0, 255)  # Green color in BGR
        # thickness = 2
        # font = cv2.FONT_HERSHEY_SIMPLEX
        # cv2.putText(image, text, (box[0], box[1]), font, 1, color, thickness)
    return Cy_list, Cx_list, all_text

class Document:
    def __init__(self, fullPath):
        self.full_path = fullPath

    def indexFromFile(self,):
        self.refIndex = [None]*len(self.pages)
        jsonPath = os.path.join(self.doc_dir, self.doc_name+'.pdf.json')
        try:
            with open(jsonPath) as f:
                jsonInfo = json.load(f)[self.doc_name+'.pdf']
            for i, v in enumerate(jsonInfo.values()):
                try: self.refIndex[i]= (np.array(v)[:, 0]).tolist()
                except: pass
        except: pass

    def check_scan_or_digit(self):
        '''
        Check if pdf is digital or scanned.
        '''
        d = self.digit_page.get_text_words()
        if len(d) > 10:# and digit_flag:
            return True
        else:
            return False

    def getCells(self, Cy_list, Cx_list, all_text, consignee_x, consignee_y, props="consignee"):
        minY, maxY = consignee_y, consignee_y+cell_height-10
        if props == "pronumber":
            minX_2 = consignee_x-100+consignee_width
            maxX_2 = minX_2 + pro_num_width
        else:
            minX_2, maxX_2 = consignee_x-100, consignee_x-100+consignee_width
        filtered_elements_index = [k for k, x in enumerate(Cy_list) if x > minY and x < maxY]
        Cy_list_1, Cx_list_1, all_text_1 = [], [], []
        for k in filtered_elements_index:
            Cy_list_1.append(Cy_list[k])
            Cx_list_1.append(Cx_list[k])
            all_text_1.append(all_text[k])
        filtered_elements_index = [k for k, x in enumerate(Cx_list_1) if x > minX_2 and x < maxX_2]
        Cy_list_2, Cx_list_2, all_text_2 = [], [], []
        for k in filtered_elements_index:
            Cy_list_2.append(Cy_list_1[k])
            Cx_list_2.append(Cx_list_1[k])
            all_text_2.append(all_text_1[k])       

        return all_text_2  

    def parse_page(self):
        '''
        main process.
        '''
        # checkDigit = self.check_scan_or_digit()

        if self.img.shape[0]>self.img.shape[1]:
            self.img = cv2.rotate(self.img, cv2.ROTATE_90_COUNTERCLOCKWISE)
        Cy_list, Cx_list, all_text = getTextAndCoorFromPaddle(self.img)
        
        # get index of consignee element
        for i, text in enumerate(all_text):
            if "consignee" in text.lower():
                consignee_index = i
                consignee_y, consignee_x = Cy_list[i], Cx_list[i]
                break


        print()
               
        name_city_list, digit_list = [], []
        try:
            if self.page_num == 1:
                for i in range(4):
                    consignee_list = self.getCells(Cy_list, Cx_list, all_text, consignee_x, consignee_y+cell_height*i)
                    pro_number = self.getCells(Cy_list, Cx_list, all_text, consignee_x, consignee_y+cell_height*i, props="pronumber")
                    name_city_list.append([consignee_list[0], consignee_list[-1].split(',')[0]])
                    digits = ''.join(char for char in pro_number[0] if char.isdigit())
                    digit_list.append(digits)
                    

            else:
                for i in range(8):
                    consignee_list = self.getCells(Cy_list, Cx_list, all_text, consignee_x, consignee_y+cell_height*i)
                    pro_number = self.getCells(Cy_list, Cx_list, all_text, consignee_x, consignee_y+cell_height*i, props="pronumber")         
                    name_city_list.append([consignee_list[0], consignee_list[-1].split(',')[0]])
                    digits = ''.join(char for char in pro_number[0] if char.isdigit())
                    digit_list.append(digits)
        except:
            pass

        return name_city_list, digit_list   

    def parse_doc(self,):
        '''
        In a document, main process is done for all pages 
        '''
        # Split and convert pages to images     
        self.digit_doc = fitz.open(self.full_path)
        past_page_ind, page_batch = 0, 50
        name_city_list, digit_list = [], []
        for i in range(len(self.digit_doc)//page_batch+1):
            next_page_ind = min(page_batch*(i+1), len(self.digit_doc))
            pages = split_pages(self.digit_doc, past_page_ind, next_page_ind)
            past_page_ind = next_page_ind
            self.pages = pages     
            for idx, img in enumerate(self.pages):
                try:
                    # if idx < 2:
                        self.digit_page = self.digit_doc[idx+i*page_batch]
                        self.page_num = idx + 1 + i*page_batch
                        print(f"Reading page {self.page_num} out of {len(self.digit_doc)}")
                        self.img = img
                        self.digit_cen_value = []
                        self.digit_value = []                      
                        v1, v2 = self.parse_page()
                        name_city_list += v1
                        digit_list += v2
                except Exception as e:
                    print(f"    Page {str(idx+1)} of {self.full_path} ran into warning(some errors) in while parsing.")
        print(f"    Completed parsing {self.full_path} with no errors, ...........OK")

        return name_city_list, digit_list
if __name__ == "__main__":
    start = datetime.datetime.now()
    input_folder = "inputs"
    pdf_list = [f for f in os.listdir(input_folder) if (f.split('.')[-1].lower() in ['pdf'])]
    for filename in pdf_list:
        # parser = argparse.ArgumentParser(description='Generate a Word document with tables.')
        # parser.add_argument('file_name', type=str, help='input pdf file name in input folder')
        # args = parser.parse_args()
        # filename = args.file_name
        file_path = os.path.join(input_folder, filename)
        pdf_doc = Document(file_path)
        name_city_list, digit_list = pdf_doc.parse_doc()
        add_name_city(name_city_list, digit_list, filename)
    end = datetime.datetime.now()
    print(f" Time: {end-start}")
