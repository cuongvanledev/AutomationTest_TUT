import shutil
import TUT_GenReport
import os
import sys


Msn = sys.argv[1]

Output = TUT_GenReport.workSpace(Msn)

if Msn.lower() == "Generic":
    shutil.copy2(Output + "RH850_X2x_Generic_TUT_CoverageResult_U2A8_Beta.zip", "U:/internal/Module/" + Msn.lower() + "/07_UT/01_WorkProduct_T/result/U2A8/Beta/test_report/")
else:
    fileNameE2x = "RH850_E2x_" + Msn.upper() + "_TUT_U2A8_Beta.zip"
    fileNameU2x = "RH850_U2x_" + Msn.upper() + "_TUT_U2A8_Beta.zip"
    TUT_GenReport.renameFile(Output + "RH850_E2x_" + Msn + "_TUT_U2A8_Beta.zip", Output + fileNameE2x)
    TUT_GenReport.renameFile(Output + "RH850_U2x_" + Msn + "_TUT_U2A8_Beta.zip", Output + fileNameU2x)
    shutil.copy2(Output + fileNameE2x, "U:/internal/Module/" + Msn.lower() + "/07_UT/01_WorkProduct_T/result/U2A8/Beta/test_report/")
    shutil.copy2(Output + fileNameU2x, "U:/internal/Module/" + Msn.lower() + "/07_UT/01_WorkProduct_T/result/U2A8/Beta/test_report/")