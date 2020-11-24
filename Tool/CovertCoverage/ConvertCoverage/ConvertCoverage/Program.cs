using Microsoft.VisualStudio.Coverage.Analysis;

namespace CoverageConverter
{
    class Program
    {
        static void Main(string[] args)
        {

            /*
             * pathCoverageFile: Coverage file need to convert to xml
             * pathOutputXMLFile: Coveragexml file coverted from .coverage file
             */
            string pathCoverageFile = args[0];
            string pathOutputXMLFile = args[1];
            using (CoverageInfo info = CoverageInfo.CreateFromFile(
                pathCoverageFile,
                new string[] { @"DIRECTORY_OF_YOUR_DLL_OR_EXE" },
                new string[] { }))
            {
                CoverageDS data = info.BuildDataSet();
                data.WriteXml(pathOutputXMLFile);
            }
        }
    }
}