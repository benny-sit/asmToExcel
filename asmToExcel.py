import os
import xlsxwriter as xw
from pathlib import Path
import shutil
import zipfile
patool_availabe = False
try:
    patool_availabe = True
    import patoolib
except ImportError:
    patool_availabe = False
    print("cannot import unrar")

# CONSTANTS
EXCEL_NAME = 'asmToExcel.xlsx'
SHEET_NAME = 'asmFiles'


def excludeFolder(d):
    '''
    The Folders you want to exclude from running on
    :param d: directory path
    :return: True / False -> filter sub folders
    '''
    return '__' in d or d.startswith('.') or d in ['dist', 'venv', 'img', 'build']


def moveFilesToMain(folder_path):
    '''
    removing all sub folders but keeping content in main folder
    :param folder_path: directory path
    :return: None | just organizing the sub folders
    '''
    # MOVE FILES
    # Smartest Way to skip main dir ->
    walker = os.walk(folder_path)
    next(walker)
    for root, dirs, files in walker:

        for file in files:
            shutil.move(os.path.join(root, file), folder_path)

    # CLEAN UP
    [shutil.rmtree(f.path) for f in os.scandir(folder_path) if f.is_dir()]


def unzipOrUnrarFolder(folder, file):
    '''
    Trying to unpack the folder
    :param folder: directory path
    :param file: filename -> endswith .zip/.rar (with patool is possible to use other formats)
    :return: True if file was decompressed else False
    '''
    try:
        patoolib.extract_archive(os.path.join(folder, file), outdir=folder)
        return True
    except:
        if file.endswith('.zip'):
            with zipfile.ZipFile(os.path.join(folder, file)) as zip_ref:
                zip_ref.extractall(folder)
            return True
    return False


def unZipSubfolders():
    '''
    Iterating through all included sub folders and decompressing + organizing
    :return: None | changes sub directories
    '''
    for root, dirs, files in os.walk(os.getcwd()):
        [dirs.remove(d) for d in list(dirs) if excludeFolder(d)]
        if root == os.getcwd(): continue
        if not excludeFolder(root):
            for file in files:
                if file.endswith('.zip') or file.endswith('.rar'):
                    isUnzipped = unzipOrUnrarFolder(root, file)
                    if isUnzipped: os.remove(os.path.join(root, file))
            moveFilesToMain(root)


def correctIdentation(lines):
    '''
    indenting the correct lines
    @param lines: List[List[str]] every line with every word separated
    @return: List[str] every string is a line with correct indentation
    '''
    indent = ' ' * 4
    isBlockIndent = False
    res = []
    for line in lines[:]:
        isFunc = any('proc' == w or 'endp' == w for w in line)
        if line[0] == 'ret' or ':' in line[0] or isFunc:
            if any('proc' == w for w in line) or ':' in line[0]:
                isBlockIndent = True
                res.append("\n" + ' '.join(line))
            else:
                isBlockIndent = False
                res.append(' '.join(line))
        else:
            if isBlockIndent:
                res.append(f'{indent}{line[0].ljust(7)}{" ".join(line[1:])}')
            else:
                res.append(" ".join(line))

    return res


def writeFile(ws, file_name, row, col, format):
    '''
    Writing Excel Date to cell
    @param ws: selected worksheet
    @param file_name: full file name + root
    @param row: insert in to this cell row
    @param col: insert in to this cell col
    @param format: cell format (style of cell)
    @return: None
    '''
    with open(file_name, 'r') as f:
        lines = list(filter(lambda l: l, map(lambda l: l[:l.index(';')] if ';' in l else l, [l.replace(';', ' ; ').split() for l in f.readlines()])))

        # TO ADD -> checks on file

        lines = correctIdentation(lines)
        ws.write(row, col, '\n'.join(lines), format)


def toExcel():
    '''
    Opening all .asm files in subfolders and adding them to a cell in excel
    @return: None ( creating excel file )
    '''
    # Unziping the subfolders in os.getcwd()
    unZipSubfolders()

    # Setting the excel and the format of the cells
    asmFiles = xw.Workbook(os.path.join(os.getcwd(), EXCEL_NAME))
    wsAsm = asmFiles.add_worksheet(SHEET_NAME)
    text_format = asmFiles.add_format({'text_wrap': True, 'valign': 'top'})
    header_format = asmFiles.add_format({'bold': True, 'font_size': 20, 'font_color': '#777777', 'bg_color': '#efefef', 'bottom': 2, 'bottom_color': '#333333'})

    r = 1
    c = 0
    # Iterating Through subfolders
    for root, dirs, files in os.walk(os.getcwd()):
        [dirs.remove(d) for d in list(dirs) if excludeFolder(d)]
        try:
            # Stop iterating if it is subfolder of subfolder
            two_up = os.path.abspath(os.path.join(root, '..', '..'))
            if os.getcwd() == two_up or Path(os.getcwd()) in Path(two_up).parents :
                print("Stopping", root)
                continue
        except IndexError:
            continue

        r = 1
        if root == os.getcwd():
            r = 0
            for d in dirs:
                if '__' not in d and not d.startswith('.') and d not in ['dist', 'venv', 'img', 'build']:
                    dirName = ' '.join(d.split('_')[-2:])
                    if dirName[-1].isdigit():
                        dirName = d.split('-')[0]
                    wsAsm.write(r, c, dirName.capitalize(), header_format)
                    # wsAsm.set_row(r, HEADER_ROW_HEIGHT)
                    c += 1
            c = 0
        else:
            for file in files:
                if file.endswith(".asm"):
                    writeFile(wsAsm, os.path.join(root, file), r, c, text_format)
                    r += 1
            c += 1

    wsAsm.set_column(0, c, 30)

    asmFiles.close()


if __name__ == '__main__':
    toExcel()
    # print(list(os.walk(os.getcwd())))