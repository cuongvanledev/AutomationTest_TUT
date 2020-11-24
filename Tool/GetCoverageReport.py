import openpyxl 
from xml.dom import minidom
import sys
import os
import re
from dataGen import MethodsName
import dataGen
import TUT_GenReport
import GenResult
import shutil

# argv[1]: Msn (Adc, Spi...)
# argv[2]: Copy and commit or not
Msn = sys.argv[1]
Commit = sys.argv[2]

# Check Msn: Generic or other module
# If Msn is Generic then it not have device
if Msn.lower() == "generic":
    pathCoveraged = [TUT_GenReport.workSpace(Msn) + Msn + "\\" + Msn + "_DynamicCodeCoverage.coveragexml"]
    pathResult = [TUT_GenReport.workSpace(Msn) + Msn + "\\" + "Test_Result_" + Msn + ".trx"]
    devices = ["None"]
    numDevice = 1
else:
    devices = ["E2x", "U2x"]
    pathCoveraged = [TUT_GenReport.workSpace(Msn) + devices[0] + "\\E2x_DynamicCodeCoverage.coveragexml", TUT_GenReport.workSpace(Msn) + devices[1] + "\\U2x_DynamicCodeCoverage.coveragexml"]
    pathResult = [TUT_GenReport.workSpace(Msn) + devices[0] + "\\Test_Result_E2x.trx", TUT_GenReport.workSpace(Msn) + devices[1] + "\\Test_Result_U2x.trx"]
    numDevice = 2

# Store data into dictionary
NameSpaceDict_Ux = dict()
NameSpaceDict_Ex = dict()
nsList = [NameSpaceDict_Ex, NameSpaceDict_Ux]
ClassDict = dict()
rowCurrent = 0

# Read .coveragexml of each devices
for n in range(numDevice):
    xmldoc = minidom.parse(pathCoveraged[n])
    modules = xmldoc.getElementsByTagName('CoverageDSPriv')
    
    for module in modules:
        allModule = module.getElementsByTagName('Module')
        for m in allModule:
            flagModule = False
            moduleName = m.getElementsByTagName("ModuleName")[0].firstChild.nodeValue
            if Msn.lower() == "generic":
                if moduleName.find('tut')!=0:
                    flagModule = True
            else:
                if moduleName.find('tut')!=0  and moduleName.find('mcalconfgen')!=0 and moduleName.find('gut')!=0:
                    flagModule = True
            #Only get module MCAL such as Dio, Adc. Ignore module tut_xxx and mcalconfgen
            if flagModule == True:

                # Get all Namespace
                NamespaceTables = m.getElementsByTagName('NamespaceTable')
                for nameSpace_ in NamespaceTables:

                    #Get all class
                    classes = nameSpace_.getElementsByTagName('Class')
                    for c in classes:

                        #Get all method in class
                        functions = c.getElementsByTagName('Method')
                        for ns in functions:
                            NamespaceName = str(nameSpace_.getElementsByTagName('NamespaceName')[0].firstChild.nodeValue)
                            Class = str(c.getElementsByTagName('ClassName')[0].firstChild.nodeValue)
                            method = str(ns.getElementsByTagName('MethodName')[0].firstChild.nodeValue)
                            if re.search(r"\w+\.<>", Class) is not None:
                                continue

                            # Get line  coveraged of method
                            LinesCovered = float(ns.getElementsByTagName('LinesCovered')[0].firstChild.nodeValue) + float(ns.getElementsByTagName('LinesPartiallyCovered')[0].firstChild.nodeValue)
                            lines_not_covered = float(ns.getElementsByTagName('LinesNotCovered')[0].firstChild.nodeValue)

                            # Get 4 numbers after dot: Eg. 0.4324245 => result = 0.4323
                            LinesCovered_100 = float(LinesCovered/(LinesCovered + lines_not_covered))
                            n_ = re.search('\d+\.\d{4}', str(LinesCovered_100))
                            if n_:
                                LinesCovered_100 = float(n_.group(0))
                            LinesNotCovered_100 = 1 - LinesCovered_100
                            
                            # Get block coveraged of method
                            BlocksCovered = float(ns.getElementsByTagName('BlocksCovered')[0].firstChild.nodeValue)
                            blocks_not_covered = float(ns.getElementsByTagName('BlocksNotCovered')[0].firstChild.nodeValue)
                            BlocksCovered_100 = float(BlocksCovered/(BlocksCovered + blocks_not_covered))
                            
                            # Get 4 numbers after dot: Eg. 0.4324245 => result = 0.4323
                            m = re.search('\d+\.\d{4}', str(BlocksCovered_100))
                            if m:
                                BlocksCovered_100 = float(m.group(0))
                            BlocksNotCovered_100 = 1 - BlocksCovered_100

                            PartiallyCover = float(ns.getElementsByTagName('LinesPartiallyCovered')[0].firstChild.nodeValue)
                            PartiallyCover_100 = 0

                            #Store data
                            p = MethodsName(BlocksCovered_100, LinesCovered_100, BlocksNotCovered_100, LinesNotCovered_100, BlocksCovered, blocks_not_covered, LinesCovered, lines_not_covered, PartiallyCover, PartiallyCover_100)
                            cl = dict()
                            nsp = dict()
                            if NamespaceName in nsList[n]:
                                if Class in nsList[n][NamespaceName]:
                                    a = 1
                                else:
                                    nsList[n][NamespaceName][Class] = dict()
                            else:
                                nsList[n][NamespaceName] = dict()
                                nsList[n][NamespaceName][Class] = dict()

                            cl = nsList[n][NamespaceName][Class]
                            if method in cl:
                                method += "_temp"
                            cl.update({method: p.methodNames})
    TUT_GenReport.genLogFile(pathResult[n], Msn, devices[n])
    GenResult.createResult(Msn)

listNameSpace = list(nsList[0].keys())

# Append data of E2x into U2x 
for ns in listNameSpace:
    if ns in nsList[1]:
        listClass = list(nsList[0][ns].keys())        
        for cl in listClass:
            res = dataGen.methodCommon(nsList[0][ns][cl], nsList[1][ns][cl])
            if res == 0:
                nsList[1][ns].update({cl + '_temp' : nsList[0][ns][cl]})
    else:
        nsList[1].update({ns : nsList[0][ns]})

# Call GenerateClassInfo to calculator information of class include:
# Lines number, Block, % of both
Dict_Result = dataGen.GenerateClassInfo(nsList[1])

# Call summarySheet to create Result Summary and store information from Dict_Result to sheet
TUT_GenReport.summarySheet(Dict_Result, Msn)

# Call classSheet to create each class sheet
TUT_GenReport.classSheet(nsList[1], Msn, Dict_Result, Commit)

#rename Coverage_result.coverage file to RH850_X2x_Msn"_TUT_CoverageResult_U2A8_Beta.coverage file
for dev in devices:
    if dev == "None":
        pathDir = TUT_GenReport.workSpace(Msn) + Msn + "\\Coverage_result.coverage"
        pathSrc = TUT_GenReport.workSpace(Msn) + Msn + "\\RH850_X2x_" + Msn + "_TUT_CoverageResult_U2A8_Beta.coverage"
    else:
        pathDir = TUT_GenReport.workSpace(Msn) + dev + "\\Coverage_result.coverage"
        pathSrc = TUT_GenReport.workSpace(Msn) + dev + "\\RH850_" + dev + "_" + Msn.upper() + "_TUT_CoverageResult_U2A8_Beta.coverage"
    TUT_GenReport.renameFile(pathDir, pathSrc)
    if os.path.exists(pathDir):
        os.remove(pathDir)
if Commit == "Yes":
    fileName = "RH850_X2x_" + Msn.upper() + "_TUT_CR_U2A8_Beta.xlsx"
    Output = TUT_GenReport.workSpace(Msn)
    shutil.copy2(Output + fileName, "U:/internal/Module/" + Msn.lower() + "/07_UT/01_WorkProduct_T/result/U2A8/Beta/test_report/")

