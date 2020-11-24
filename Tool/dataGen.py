import sys
import os
import re

class MethodsName:
    def __init__(self, blockCv, lineCv, blockNCv, lineNCv, blocksCv, blocksNCv, linesCv, linesNCv, PartiallyCover, PartiallyCover_100):
        self.methodNames = {
            "block_coverage" : blockCv,         #per100
            "line_coverage" : lineCv,           #per100
            "block_not_coverage" : blockNCv,    #per100
            "line_not_coverage" : lineNCv,      #per100
            "blocks_covered" : blocksCv,        #number
            "lines_covered" : linesCv,          #number
            "blocks_not_covered" : blocksNCv,   #number
            "lines_not_covered" : linesNCv,     #number
            "PartiallyCover" : PartiallyCover,
            "PartiallyCover_100" : PartiallyCover_100,
        }

# Function: Compare 2 dict 
# Return: 1 - If they are the same
#         2 - If they are diffrence
def compareDict(method, otherDict):
    if cmp(method, otherDict) == 0:
        return 1
    return 0

# Return:
#    0: existed but not the same
#    1: existed and the same
#   -1: not exist and difference
def methodCommon(classDic, classSrc):
    listMethod = list(classDic.keys())
    for mt in listMethod:
        res = compareDict(classDic[mt], classSrc[mt])
        if res == 0:
            return res
    return -1

# Function: Get 2 number after the point
#           Ex. 10.934555
# Return:   => 10.93
def getFormatNumber(value):
    value *= 100
    n = re.search('\d+\.\d{2}', str(value))
    if n:
        value = float(n.group(0))
    value = value*1.0/100
    return value

# Function: Read all classes in all NameSpaces, then get information of classes, include:
#           Lines Coveraged, Blocks Coverage
#           Lines Not Coveraged, Blocks Not Coverage
# Return:   Information of all Classes
#           A Dict contain all Classes
# Paramter: A Dict contain all NameSpaces
def GenerateClassInfo(dicts):
    listNS = list(dicts.keys())
    listCLInfo = dict()
    for ns in listNS:
        listCL = list(dicts[ns].keys())
        for cl in listCL:
            listMT = list(dicts[ns][cl].keys())
            lines_covered = 0
            lines_not_covered = 0
            blocks_covered = 0
            blocks_not_covered = 0
            partiallyCover = 0
            ns_class = ns + "." + cl
            for m in listMT:
                lines_covered += dicts[ns][cl][m].get("lines_covered")
                lines_not_covered += dicts[ns][cl][m].get("lines_not_covered")
                blocks_covered += dicts[ns][cl][m].get("blocks_covered")
                blocks_not_covered += dicts[ns][cl][m].get("blocks_not_covered")
                partiallyCover += dicts[ns][cl][m].get("PartiallyCover")
            #lines_not_covered += partiallyCover
            lines_covered_100 = getFormatNumber((lines_covered*1.0 / (lines_covered + lines_not_covered)))
            lines_not_covered_100 = 1 - lines_covered_100
            blocks_covered_100 = getFormatNumber((blocks_covered*1.0 / (blocks_covered + blocks_not_covered)))
            blocks_not_covered_100 = 1 - blocks_covered_100
            partiallyCover_100 = (partiallyCover*1.0 / (lines_covered + lines_not_covered + partiallyCover))
            clP = MethodsName(blocks_covered_100, lines_covered_100, blocks_not_covered_100, lines_not_covered_100, blocks_covered, blocks_not_covered, lines_covered, lines_not_covered, partiallyCover, partiallyCover_100)
            listCLInfo.update({ns_class: clP.methodNames})
    return listCLInfo

# Function: Contain list svn
# Return:   Contain list svn
# Paramter: 
#           module: Contain name of module
def SvnFile(module):
    listSVN = [
    "U:/internal/Module/generic/06_CD/01_WorkProduct/generator_cs",
    "U:/internal/Module/" + module.lower() + "/06_CD/01_WorkProduct/generator_cs",
    "U:/internal/Module/" + module.lower() + "/04_AD/01_WorkProduct/RH850_X2x_" + module.upper() + "_ParameterDefinition.xlsx",
    "U:/internal/Module/" + module.lower() + "/04_AD/01_WorkProduct/RH850_X2x_" + module.upper() + "_GenTool_ErrorList.xlsx",
    "U:/internal/Module/" + module.lower() + "/04_AD/01_WorkProduct/RH850_X2x_" + module.upper() + "_Configuration.xlsx",
    "U:/internal/Module/" + module.lower() + "/05_UD/01_WorkProduct/RH850_X2x_" + module.upper() + "_GenTool_UD.docx"]

    listTitle = [
        "Project Source Code-Generic",
        "Source Code " + module.upper(),
        "ParameterDefinition (PDF) \nRH850_X2x_" + module.upper() +"_ParameterDefinition.xlsx",
        "GenTool_ErrorList \nRH850_X2x_" + module.upper() +"_GenTool_ErrorList.xlsx",
        "Configuration (CDF) \nRH850_X2x_" + module.upper() +"_Configuration.xls",
        "TUD \nRH850_X2x_" + module.upper() +"_GenTool_UD.docx"
    ]
    return listSVN, listTitle

# Function: Contain list svn
# Return:   Contain list svn
# Paramter: 
#           module: Contain name of module
def SvnFileGeneric():
    listSVN = [
    "U:/internal/Module/generic/06_CD/01_WorkProduct/generator_cs",
    "U:/internal/Module/generic/04_AD/01_WorkProduct/RH850_X2x_Generic_GenTool_ErrorList.xlsx",
    "U:/internal/Module/generic/05_UD/01_WorkProduct/RH850_X2x_Generic_GenTool_UD.docx"]

    listTitle = [
        "Project Source Code-Generic",
        "GenTool_ErrorList \nRH850_X2x_Generic_GenTool_ErrorList.xlsx",
        "TUD \nRH850_X2x_Generic_GenTool_UD.docx"
    ]
    return listSVN, listTitle