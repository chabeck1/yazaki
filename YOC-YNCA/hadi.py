#import pypdf
import pdfplumber
import pandas as pd
import re

def parse_eng_part_name(fullPartName):
    newLinePositions = list()
    for match in re.finditer(r'\n', fullPartName):
        newLinePositions.append(match.start())
        print(match)
    # newline_positions = [match.start() for match in re.finditer(r'\n', text)]
    print(newLinePositions)  # Output: [5, 11, 18, 21]
    numLines = len(newLinePositions)+1
    if (numLines == 2):
        print(fullPartName[newLinePositions[0]+1:])
        return fullPartName[newLinePositions[0]+1:]
    elif (numLines == 3):
        for char in fullPartName[newLinePositions[0]+1:newLinePositions[1]]:
            if ord(char) > 0x7F:
                return fullPartName[newLinePositions[1]+1:]
        return fullPartName[newLinePositions[0]+1:]
    elif (numLines == 4):
        # for char in fullPartName[newLinePositions[0]+1:newLinePositions[3]]:
        #     if ord(char) > 0x7F:
        #         return fullPartName[newLinePositions[1]+1:]
        return fullPartName[newLinePositions[1]+1:]
    else:
        return ''

def extract_strings(row):
    newRow = list()
    for element in row:
        if element != None:
            newRow.append(element)
    return newRow

# Open the PDF file
with pdfplumber.open("66401-070A_002_1_0.pdf") as pdf:
    newTable = list()
    englishList = list()
    for page in pdf.pages:
        pageLines = dict()
        table = page.extract_tables()[0]
        # for row in table:
        #     print(row)
        # print(len(table))
        lines = page.lines
        # print(type(lines[0]))
        for line in lines:
            if (line['width'] > 0 and (line['x0'] < 150 or line['y0'] < 700)):
                pageLines[line['y0']] = line
            # print(line['y0'])
            # print(f"{line}\n")
        sortedPageLines = dict(sorted(pageLines.items(), reverse=True))
        # print(f"\n\n\n\n{sortedPageLines}")
        # print(f"{len(pageLines)}")
        rowNum = 0
        for key,value in sortedPageLines.items():
            # print(key)
            # print(value["stroking_color"])
            # print(value)
            # print(rowNum)
            if (value["stroking_color"] == (1,0,0)):
                print(f"Remove at row {rowNum}, page {value['page_number']}")
            else:
                rowNum = rowNum + 1
        # print(page.page_number)
        # print(table)

        # for row in table[3:]:
        #     noneRemoved = extract_strings(row)
        #     print(noneRemoved)
        #     if (len(noneRemoved) == 12):
        #         englishList.append(parse_eng_part_name(noneRemoved[9]))

        # newTable.append(parse_eng_part_name(row))
        #print(page.extract_text())

    # Testing rectangles
    # rects = page.rects
    # im = page.to_image()
    # imRects = im.draw_rects(rects, fill=(255,0,0), stroke=(0,255,0))
    # imRects.show()
    # pageRects = dict()
    # count = 0
    # for rect in rects:
    #     # pageRects[f"{rect['y0']},{count}"] = rect 
    #     pageRects[rect['y0']] = rect 
    #     count = count + 1
    # print(dict(sorted(pageRects.items(), reverse=True)))
    # print(englishList)