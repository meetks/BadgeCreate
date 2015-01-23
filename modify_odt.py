#!/usr/bin/python -tt 

"""
    'Example: ./modify_odt.py b-arial-black.odt Firstname Lastname  1309  ~/Pictures/computer-security_7447.jpg '
"""

import sys
import argparse
import logging
import shutil
import os
import zipfile
import tempfile
import subprocess

TEMPLATE_NAME = "badge.odt"
pictures = {'image': '2.jpg', 'left':'4.jpg', 'right': '3.jpg', 'barcode':'1.jpg'}

def exec_cmd(cmd):
  logging.info('Execing command: %s', cmd)
  subprocess.check_output(cmd, shell=True)

def replace_txt(first, lastname, ID, new_file):
  try:
    new_odt_file = zipfile.ZipFile(new_file,'a')
    xml_string = new_odt_file.read('content.xml')              
    new_xml = xml_string.replace("Firstname",first);
    new_xml = new_xml.replace("Lastname",lastname);
    new_xml = new_xml.replace("5555",ID);
    new_odt_file.writestr('content.xml', new_xml)          
    new_odt_file.close()
    return 
  except:
    logging.error( 'Unexpected error while replacing text. Bailing!!!')
    raise

def barcode_gen_ins(working_dir, ID):
  try:
    # -n option prevent number output in barcodes
    exec_cmd("barcode -E -b %s  -u in -g 3.9x1.32 -o %s/a.eps -n" % (ID, working_dir))
    exec_cmd("convert -density 600 %s/a.eps -quality 100 %s/barcode.jpg" % (working_dir, working_dir))
    exec_cmd("rm -f %s/a.eps" % working_dir)
  except:
    logging.error('Unexpected error while generating barcode. Bailing!!!')
    raise

def change_pictures(package_dir, working_dir, pic_file, phone, b1,bh,bb, br,fa ):
  try:
    exec_cmd("cp %s %s/Pictures/%s" % (pic_file, working_dir, pictures["image"]))
    exec_cmd("mv %s/barcode.jpg %s/Pictures/%s" % (working_dir, working_dir, pictures["barcode"]))

    # white out all symbols, then insert as per the command line parameters
    exec_cmd("cp %s/white.jpg %s/Pictures/%s" % (package_dir, working_dir, pictures["left"]))
    exec_cmd("cp %s/white.jpg %s/Pictures/%s" % (package_dir, working_dir, pictures["right"]))

    if phone == True:
      exec_cmd("cp %s/Phone-block.jpg %s/Pictures/%s" % (package_dir, working_dir, pictures["right"]))
    if b1 == True:
      exec_cmd("cp %s/1-block.jpg %s/Pictures/%s" % (package_dir, working_dir, pictures["left"]))
    if bh == True:
      exec_cmd("cp %s/H-block.jpg %s/Pictures/%s" % (package_dir, working_dir, pictures["right"]))
    if bb == True:
      exec_cmd("cp %s/B-block.jpg %s/Pictures/%s" % (package_dir, working_dir, pictures["left"]))
    if br == True:
      exec_cmd("cp %s/R-block.jpg %s/Pictures/%s" % (package_dir, working_dir, pictures["left"]))
    if fa  == True:
      exec_cmd("cp %s/Plus-block.jpg %s/Pictures/%s" % (package_dir, working_dir, pictures["right"]))
  except:
    logging.error( 'Unexpected error while chaning pcitures in template file . Bailing!!!')
    raise

def check_files(package_dir):
  for i in ["R-block.jpg", "Plus-block.jpg", "H-block.jpg", "1-block.jpg", "Phone-block.jpg",
            "B-block.jpg", "white.jpg"]:
    if not os.path.exists("%s/%s" % (package_dir, i)):
      logging.error ('Symbol files %s missing. Bailing', i)
      return False
  return True

def create_zip_file_replace_text(template_path, working_dir, new_file, first, lastname, ID):
  try:
    
    exec_cmd("cp %s %s" % (template_path, working_dir))
    working_path_odt = "%s/%s" % (working_dir, TEMPLATE_NAME)
    # replace in zip file without unzipping
    replace_txt(first, lastname, ID, "%s" % working_path_odt)                              
    # -o option to overwrite without prompting
    exec_cmd("unzip -o -d %s %s" % (working_dir, working_path_odt))
  except:
    logging.error( 'Unexpected error while initialization. Bailing!!!')
    raise

def get_all_files(working_dir):
  filenames = []
  for name in os.listdir(working_dir):
    path = os.path.join(working_dir, name)
    if os.path.isfile(path):
      filenames.append(path)
    elif os.path.isdir(path):
      filenames.extend(get_all_files(path))
  return filenames

def create_zip(working_dir):
  files = get_all_files(working_dir)
  _, odt_file = tempfile.mkstemp(suffix=".odt")
  z = zipfile.ZipFile(odt_file, "w")
  for file_name in files:
    arc_name = file_name.replace(working_dir + "/", '', 1)
    z.write(file_name, arcname = arc_name)
  z.close()
  return odt_file
    
      

def convert_to_png(working_dir):
  try:
    exec_cmd("rm -f %s/%s" % (working_dir, TEMPLATE_NAME))
    odt_file = create_zip(working_dir)
    
    exec_cmd("libreoffice --headless --invisible  --convert-to pdf %s --outdir %s" % (odt_file, tempfile.gettempdir()))
    pdf_file = odt_file[:-3] + "pdf"
    exec_cmd("convert -density 600 %s -quality 120 %s" % (pdf_file, pdf_file[:-3]+"png"))
    exec_cmd("rm -rf %s" % working_dir)
    # TODO remove pdf, odt files
  except:
    logging.error( 'Unexpected error while coverting to html. Bailing!!!')
    raise

def main():

    parser = argparse.ArgumentParser()
    parser.add_argument(action="store",dest="firstname", help="Specify the first name" )
    parser.add_argument(action="store",dest="lastname", help="Specify the last name" )
    parser.add_argument(action="store",dest="sevadar_id", help="Specify the sevadar id" )
    parser.add_argument(action="store",dest="picture_filename", help="Specify the picture file" )
    parser.add_argument("-phone", action="store_true", default=False,help="Can use Phone?")
    parser.add_argument("-blockh", action="store_true", default=False,help="H-Block ")
    parser.add_argument("-blockr", action="store_true", default=False,help="R-Block ")
    parser.add_argument("-blockb", action="store_true", default=False,help="B-Block ")
    parser.add_argument("-block1", action="store_true", default=False,help="1-Block ")
    parser.add_argument("-firstaid", action="store_true", default=False,help="First Aid")
    parser.add_argument("-v", action="store_true", default=False,help="Verbose Help")
    args = parser.parse_args()
    
    package_dir = os.path.dirname(sys.argv[0])
    working_dir = tempfile.mkdtemp()
    template_path = "%s/%s" % (package_dir, TEMPLATE_NAME)
    logging.info("Using temporary directory:%s", working_dir)
    if not(os.path.exists(template_path)
           and os.path.exists(args.picture_filename)) or not check_files(package_dir) :
      logging.error ('Unable to find input file or symbols missing. Please check the path. Bailing')
      sys.exit(2)

    first= args.firstname
    lastname= args.lastname
    ID=args.sevadar_id 
    pic_file=args.picture_filename
    phone = args.phone
    blockh = args.blockh
    blockb = args.blockb
    blockr = args.blockr
    block1 = args.block1
    firstaid = args.firstaid
    
    new_file = "badge_" + str(ID) + ".odt"
    logging.info("%s", new_file)
 
    create_zip_file_replace_text(template_path, working_dir, new_file, first, lastname, ID)

    barcode_gen_ins(working_dir, ID)

    change_pictures(package_dir, working_dir, pic_file, phone, block1, blockh, blockb, blockr, firstaid)
    
    convert_to_png(working_dir)

if __name__ == '__main__':
  logging.basicConfig(level=logging.INFO,)
  main()
   
