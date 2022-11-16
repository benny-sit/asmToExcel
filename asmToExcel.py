import os
import xlsxwriter as xw
import zipfile
from pathlib import Path
import shutil


# CONSTANTS
EXCEL_NAME = 'asmToExcel.xlsx'
SHEET_NAME = 'asmFiles'


def upOneDir(path):
    try:
        # from Python 3.6
        parent_dir = Path(path).parents[1]
        # for Python 3.4/3.5, use str to convert the path to string
        # parent_dir = str(Path(path).parents[1])
        shutil.move(path, parent_dir)
    except IndexError:
        # no upper directory
        print("No upper directory / error")


def extractUp(folder_path):
    subfolders = [f.path for f in os.scandir(folder_path) if f.is_dir()]
    for d in subfolders:
        for f in os.listdir(d):
            if os.path.isdir(os.path.join(d, f)):
                extractUp(os.path.join(d, f))
            else:
                upOneDir(os.path.join(d, f))

        try:
            os.rmdir(d)
        except:
            print("Cannot rm dir")


def unzipFolder(folder_path):
    with zipfile.ZipFile(folder_path, 'r') as zip_ref:
        zip_ref.extractall(os.path.dirname(folder_path))


def unZipSubfolders():
    for root, dirs, files in os.walk(os.getcwd()):
        if root == os.getcwd(): continue
        for file in files:
            if file.endswith('.zip'):
                unzipFolder(os.path.join(root, file))
                os.remove(os.path.join(root, file))
        extractUp(root)


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
                if '__' not in d and not d.startswith('.'):
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