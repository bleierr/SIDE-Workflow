"""
Script by Roman Bleier

Written for Python 3.5 (on Win)

"""

import os, glob, ntpath
import subprocess
import zipfile
import xml.etree.ElementTree as ET

NSMAP = {"office" : "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "style" : "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    "text" : "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "table" : "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "draw" : "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    "fo" : "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
    "xlink" : "http://www.w3.org/1999/xlink",
      "dc":"http://purl.org/dc/elements/1.1/",
    "meta" : "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
    "number" : "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
    "svg" : "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
    "chart" : "urn:oasis:names:tc:opendocument:xmlns:chart:1.0",
    "dr3d" : "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0",
    "math" : "http://www.w3.org/1998/Math/MathML",
    "form" : "urn:oasis:names:tc:opendocument:xmlns:form:1.0",
    "scriptvurn":"oasis:names:tc:opendocument:xmlns:script:1.0",
    "ooo" : "http://openoffice.org/2004/office",
    "ooow" : "http://openoffice.org/2004/writer",
    "oooc" : "http://openoffice.org/2004/calc",
      "dom" : "http://www.w3.org/2001/xml-events",
    "rpt" : "http://openoffice.org/2005/report",
    "of" : "urn:oasis:names:tc:opendocument:xmlns:of:1.2",
    "xhtml" : "http://www.w3.org/1999/xhtml",
      "grddl" : "http://www.w3.org/2003/g/data-view#",
    "tableooo" : "http://openoffice.org/2009/table",
    "textooo" : "http://openoffice.org/2013/office",
    "loext" : "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0",
         "textooo9": "http://openoffice.org/2009/office"}


for key, val in NSMAP.items():
    ET.register_namespace(key, val)

cwd = os.getcwd()

#Dir of the input files (ODT)
odt_dir = os.sep.join([
        cwd,
        "SIDE-Interfaces-LaTeX",
        "Articles",
        "odt-articles"
    ])

#Dir for the output files
articles_dir = os.sep.join([
        cwd,
        "SIDE-Interfaces-LaTeX",
        "Articles",
        "odt-articles"
    ])

xalan = os.sep.join(["C:",
         "Program Files",
         "Oxygen XML Editor 18",
         "lib",
         "xalan.jar"])

texTools_dir = os.sep.join([
        cwd,
        "TexTools",
    ])

xslt_f = os.sep.join([
        texTools_dir,
        "odt2latex.xsl"
    ])

styles_xml_f = os.sep.join([
        texTools_dir,
        "styles.xml"
    ])

perl_f = os.sep.join([
        cwd,
        "TexTools",
        "cleanXsltTex.pl"
    ])

def get_odt_content_styles(odt_file):
    """get content and styles xml from ODT file"""
    archive = zipfile.ZipFile(odt_file, "r")
    content = archive.read("content.xml")
    styles = archive.read("styles.xml")
    archive.close()
    return content, styles
    
def merge_styles_into_content(content, styles, article_xml_f):
    """merge styles from styles.xml into content.xml"""
    #print("injecting styles in content.xml!")
    rootStyles = ET.fromstring(styles)
    styles = rootStyles.findall("./office:styles/style:style", NSMAP)
    #print("{} styles found".format(len(styles)))
    rootContent = ET.fromstring(content)
    automatic_styles = rootContent.findall("./office:automatic-styles", NSMAP)
    #print("{} automatic-styles found".format(len(automatic_styles)))
    automatic_styles[0].extend(styles)
    
    ET.ElementTree(rootStyles).write(styles_xml_f)
    
    ET.ElementTree(rootContent).write(article_xml_f)
    #print("done injecting styles!")

def path_leaf(path):
    head, tail = ntpath.split(path)
    return tail or ntpath.basename(head)


if __name__ == "__main__":

    #get odt files from dir
    odt_files = glob.glob(odt_dir+os.sep+"*.odt")

    for odt_file in odt_files:
        content, styles = get_odt_content_styles(odt_file)
        article_name = path_leaf(odt_file)[0:-4]
        
        article_xml_f = articles_dir + os.sep + article_name+".xml"

        merge_styles_into_content(content, styles, article_xml_f)

        ttt_f = articles_dir + os.sep + "ttt"

        #Run XSLT mit XALAN
        subprocess.run(["java", "-jar", xalan, "-IN", article_xml_f, "-XSL", xslt_f, "-OUT", ttt_f])

        tex_f = articles_dir + os.sep + article_name+".tex"
        
        #Run Perl cleanXsltTex Script
        outs = subprocess.check_output(["perl", perl_f, ttt_f, article_name])
        with open(tex_f, "wb") as f:
            f.write(outs)

        #clean up
        os.remove(ttt_f)
        os.remove(article_xml_f)
        os.remove(styles_xml_f)

        print("{} Article TeX created".format(article_name))

               
            
