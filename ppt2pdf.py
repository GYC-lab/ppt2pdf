import os
import glob
from win32com.client import gencache
 
 
def get_file_path():
    """
    obtain all the ppt/pptx files in the current directory
    """
    file_path_0 = os.path.split(os.path.abspath(__file__))[0]
    file_path_in = os.path.join(file_path_0, "input")
    pp_files = glob.glob(os.path.join(file_path_in, "*.ppt*"))
    return file_path_0, pp_files
 
def ppt_to_pdf(filename, results):

    name = os.path.basename(filename).split('.')[0] + '.pdf'
    exportfile = os.path.join(results, name)
    if os.path.isfile(exportfile):
        print(name, "already converted")
        return
    p = gencache.EnsureDispatch("PowerPoint.Application")
    try:
        ppt = p.Presentations.Open(filename, False, False, False)
    except Exception as e:
        print(os.path.split(filename)[1], "conversion failed, the failure reason is %s" % e)
    ppt.ExportAsFixedFormat(exportfile, 2, PrintRange=None)
    print('saving PDF file(s): ', exportfile)
    p.Quit()
 
def main():

    file_path_0, pp_files = get_file_path()
    results = os.path.join(file_path_0, "output")
    if not os.path.exists(results):
        os.mkdir(os.path.join(results))

    # clear output
    for root, dirs, files in os.walk(results, topdown=False):
        for name in files:  
            os.unlink(os.path.join(root, name))
        for name in dirs:
            os.rmdir(os.path.join(root, name))

    for pp_file in pp_files:
        ppt_to_pdf(pp_file, results)
 
    print('Done! Enjoy your time!')

if __name__ == "__main__":
    main()