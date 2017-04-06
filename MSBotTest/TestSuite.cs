using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text.RegularExpressions;

namespace MSBotTest
{
    /// <summary>
    /// Test suite class
    /// </summary>
    public class TestSuite
    {
        /// <summary>
        /// List of test cases
        /// </summary>
        public List<TestCase> TestCases { get; set; }


        /// <summary>
        /// constructor to load test cases from excel file
        /// </summary>
        /// <param name="fileName"></param>
        public TestSuite(string fileName)
        {
            TestCases = new List<TestCase>();

            // code to fix the broken uri in excel files
            // from: http://ericwhite.com/blog/handling-invalid-hyperlinks-openxmlpackageexception-in-the-open-xml-sdk/
            using (FileStream fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                UriFixer.FixInvalidUri(fs, brokenUri => { return new Uri("http://broken-link/"); });
            }

            // open the file
            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(fileName, true))
            {
                // read all sheets
                IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();

                WorkbookPart workBookPart = spreadSheetDocument.WorkbookPart;
                var indexSheet = workBookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == "Index").First();
                WorksheetPart wsPart = workBookPart.GetPartById(indexSheet.Id) as WorksheetPart;

                //var indexSheet = sheets.Where(s => s.Name == "Index").First();
                var rows = wsPart.Worksheet.Descendants<Row>(); // indexSheet.GetFirstChild<SheetData>().Descendants<Row>();


                foreach (var row in rows)
                {
                    if (row.RowIndex == 1) continue; // skip headers
                    var cells = row.Descendants<Cell>();
                    TestCase testCase = new TestCase();
                    foreach (var cell in cells)
                    {
                        switch (ColumnIndex(cell.CellReference))
                        {
                            case 0: // test case sheet name
                                testCase.SheetName = indexSheet.GetCellValue(cell.CellReference);
                                break;
                            case 1: // test case description
                                testCase.Description = indexSheet.GetCellValue(cell.CellReference);
                                break;
                        }
                    }

                    if (!string.IsNullOrEmpty(testCase.SheetName) && !string.IsNullOrEmpty(testCase.Description))
                    {
                        // valid test case
                        // check for the existance of sheet as well
                        if (sheets.Where(s => s.Name == testCase.SheetName).Any())
                        {
                            // run through the steps
                            testCase.Steps = GetTestCaseSteps(testCase, sheets, workBookPart);
                            TestCases.Add(testCase);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Get the steps for test case
        /// </summary>
        /// <param name="testCase"></param>
        /// <param name="sheets"></param>
        /// <param name="workBookPart"></param>
        /// <returns></returns>
        private List<TestCaseStep> GetTestCaseSteps(TestCase testCase, IEnumerable<Sheet> sheets, WorkbookPart workBookPart)
        {
            //var testCaseSheet = sheets.Where(s => s.Name == testCase.SheetName).First();

            var testCaseSheet = workBookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == testCase.SheetName).First();
            WorksheetPart wsPart = workBookPart.GetPartById(testCaseSheet.Id) as WorksheetPart;

            var rows = wsPart.Worksheet.Descendants<Row>(); // testCaseSheet.Descendants<Row>();
            int number = 1;
            List<TestCaseStep> TestCaseSteps = new List<TestCaseStep>();
            TestCaseStep steps = new TestCaseStep() { Number = number, Actions = new List<StepAction>() };
            number++;
            TestCaseSteps.Add(steps);
            var currentStepNumber = 0;
            foreach (var row in rows)
            {
                if (row.RowIndex == 1) continue; // skip headers
                var cells = row.Descendants<Cell>();

                var stepNumberCellReference = string.Format("A{0}", row.RowIndex);
                var userInputCellReference = string.Format("B{0}", row.RowIndex);
                var botResponseCellReference = string.Format("C{0}", row.RowIndex);
                var entitiesCellReference = string.Format("D{0}", row.RowIndex);

                var stepNumber = testCaseSheet.GetCellValue(stepNumberCellReference);
                var userInput = testCaseSheet.GetCellValue(userInputCellReference);
                var botResponse = testCaseSheet.GetCellValue(botResponseCellReference);
                var entities = testCaseSheet.GetCellValue(entitiesCellReference);

                List<Entity> entitiesList = GetEntities(userInput, entities);

                if (!string.IsNullOrEmpty(stepNumber))
                {
                    // note the current step number
                    currentStepNumber = Convert.ToInt32(stepNumber);
                    // new step action
                    StepAction action = new StepAction()
                    {
                        StepNumber = currentStepNumber,
                        Input = userInput,
                        ExpectedResponse = botResponse,
                        FirstInOrder = true,
                        Entities = entitiesList
                    };

                    foreach (var step in TestCaseSteps)
                    {
                        step.Actions.Add(action);
                    }
                }
                else
                {
                    StepAction action = new StepAction()
                    {
                        StepNumber = currentStepNumber,
                        Input = userInput,
                        ExpectedResponse = botResponse,
                        Entities = entitiesList
                    };

                    List<TestCaseStep> TempTestCaseSteps = new List<TestCaseStep>();
                    foreach (var step in TestCaseSteps)
                    {
                        if (step.Actions.Where(a => a.FirstInOrder && a.StepNumber == currentStepNumber).Any())
                        {
                            var newStep = step.Copy();
                            newStep.Actions.Remove(newStep.Actions.Where(a => a.FirstInOrder && a.StepNumber == currentStepNumber).First());
                            newStep.Number = number;
                            number++;
                            newStep.Actions.Add(action);
                            TempTestCaseSteps.Add(newStep);
                        }
                    }
                    TestCaseSteps.AddRange(TempTestCaseSteps);
                }
            }
            return TestCaseSteps;
        }

        /// <summary>
        /// get all the entities from the user input and add to the step action
        /// </summary>
        /// <param name="userInput"></param>
        /// <param name="entities"></param>
        /// <returns></returns>
        private List<Entity> GetEntities(string userInput, string entities)
        {
            var entitiesList = new List<Entity>();
            Regex regexEntities = new Regex(@"\{(\d)+\,(*)+\}");
            var m = regexEntities.Matches(entities);
            for (int c = 0; c < m.Count; c++)
            {
                var entityDefinition = m[c].Value;
                entityDefinition = entityDefinition.Substring(1, entityDefinition.Length - 2);

                int entityIndex = Convert.ToInt32(entityDefinition.Split(new char[] { ',' })[0]);
                string entityName = Convert.ToString(entityDefinition.Split(new char[] { ',' })[1]);
                if (!string.IsNullOrEmpty(entityName))
                {
                    entitiesList.Add(new Entity()
                    {
                        Index = entityIndex,
                        Name = entityName
                    });
                }
            }
            Console.ReadLine();
            return entitiesList;
        }

        public static T DeepClone<T>(T obj)
        {
            using (var ms = new MemoryStream())
            {
                var formatter = new BinaryFormatter();
                formatter.Serialize(ms, obj);
                ms.Position = 0;

                return (T)formatter.Deserialize(ms);
            }
        }

        private int ColumnIndex(string reference)
        {
            int ci = 0;
            reference = reference.ToUpper();
            for (int ix = 0; ix < reference.Length && reference[ix] >= 'A'; ix++)
                ci = (ci * 26) + ((int)reference[ix] - 64);
            return ci - 1;
        }
    }

    public class TestCase
    {
        public string SheetName { get; set; }
        public string Description { get; set; }
        public List<TestCaseStep> Steps { get; set; }
    }

    [Serializable]
    public class TestCaseStep
    {
        public int Number { get; set; }
        public List<StepAction> Actions { get; set; }
    }

    [Serializable]
    public class StepAction
    {
        public int StepNumber { get; set; }
        public bool FirstInOrder { get; set; }
        public string Input { get; set; }
        public string ExpectedResponse { get; set; }

        public List<Entity> Entities { get; set; }
    }

    [Serializable]
    public class Entity
    {
        public int Index { get; set; }
        public string Name { get; set; }
    }
}
