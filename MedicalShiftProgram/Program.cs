using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using Google.OrTools.ConstraintSolver;
using Google.OrTools.Sat;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;

class ShiftScheduler
{
    // Define constants
    static string WORKBOOK = "C:\\Users\\NoLee\\source\\repos\\MedicalShiftProgram\\nosokomeio.xlsx";
    static List<string> DAY_NAMES = new List<string> { "Δ","Τ","Τ","Π","Π","Σ","Κ" };
    static List<int> shiftsPerDayList = new List<int> { };
    static int _weekendDays;
    static int _numPeople;
    static int _numDays;
    static int _totalShifts;
    static bool _allowB2BShifts;
    static bool _balanceShiftsAutomatically;
    static string _firstDayName;

    static List<int> _weekendDaysIndices = new List<int> { };
    static List<int> _isGeneralDaysIndices = new List<int> { };
    static List<string> _people = new List<string> { };
    static List<int> _juniorIndices = new List<int> { };
    static List<int> _seniorIndices = new List<int> { };
    static List<int> _desiredShiftCounts = new List<int> { };
    static Dictionary<int, List<int>> _unavailableDays = new Dictionary<int, List<int>> { };

    static void Main()
    {
        Console.OutputEncoding = System.Text.Encoding.Unicode;
        try
        {
            Initialize();
        }
        catch (Exception e) 
        {
            Console.WriteLine($"ERROR: {e.Message.ToString()}");
            Console.WriteLine($"Stacktrace: {e.StackTrace.ToString()}");
            return;
        }

        CpModel model = new CpModel();

        // Decision variables: shift[i][j] = 1 if person i works on day j, otherwise 0
        BoolVar[,] shifts = new BoolVar[_numPeople, _numDays];
        for (int i = 0; i < _numPeople; i++)
        {
            for (int j = 0; j < _numDays; j++)
            {
                shifts[i, j] = model.NewBoolVar($"shift_{i}_{j}");
            }
        }

        // Constraint 1: Each day must have exactly i people working
        for (int j = 0; j < _numDays; j++)
        {
            var dailyWorkers = new List<ILiteral>();
            for (int i = 0; i < _numPeople; i++)
            {
                dailyWorkers.Add(shifts[i, j]);
            }
            model.Add(LinearExpr.Sum(dailyWorkers) == shiftsPerDayList[j]); // Adjust based on the number of shifts that day
        }

        // Constraint 2: No one should work two consecutive days
        for (int i = 0; i < _numPeople; i++)
        {
            for (int j = 0; j < _numDays - 2; j++)
            {
                // If person `i` works on day `j`, then they must not work on `j+1`
                model.AddBoolOr(new ILiteral[] { shifts[i, j].Not(), shifts[i, j + 1].Not() });
            }
            //Edge case for last day
            model.AddBoolOr(new ILiteral[] { shifts[i, _numDays -1].Not(), shifts[i, _numDays - 2].Not() });
        }
        // Constraint 2b: Minimize of j+2 working days for each person
        BoolVar[,] violations = new BoolVar[_numPeople, _numDays - 2]; // Only for days where j+2 is valid
        for (int i = 0; i < _numPeople; i++)
        {
            for (int j = 0; j < _numDays - 2; j++) // Ensure j+2 is within range
            {
                violations[i, j] = model.NewBoolVar($"violation_{i}_{j}");

                // Add the logic for a violation
                // violation[i, j] = shifts[i, j] AND shifts[i, j+2]
                model.AddBoolAnd(new ILiteral[] { shifts[i, j], shifts[i, j + 2] }).OnlyEnforceIf(violations[i, j]);
                model.AddBoolOr(new ILiteral[] { shifts[i, j].Not(), shifts[i, j + 2].Not() }).OnlyEnforceIf(violations[i, j].Not());
            }
        }

        // Define the objective function
        Google.OrTools.Sat.IntVar totalViolations = model.NewIntVar(0, _numPeople * (_numDays - 2), "total_violations");
        model.Add(totalViolations == LinearExpr.Sum(from i in Enumerable.Range(0, _numPeople)
                                                    from j in Enumerable.Range(0, _numDays - 2)
                                                    select violations[i, j]));

        // Constraint 3: Respect each person's unavailable days
        foreach (var person in _unavailableDays.Keys)
        {
            foreach (var day in _unavailableDays[person])
            {
                model.Add(shifts[person, day] == 0); // Ensure person is not scheduled on unavailable days
            }
        }

        // Constraint 4: Balance weekends equally among people
        //TODO this should change accordingly
        //var weekendWorkload = new List<ILiteral>[_numPeople];
        //for (int i = 0; i < _numPeople; i++)
        //{
        //    weekendWorkload[i] = new List<ILiteral>();
        //    foreach (var day in _weekendDaysIndices)
        //    {
        //        weekendWorkload[i].Add(shifts[i, day]);
        //    }
        //}
        //var weekendShifts = shiftsPerDayList.Where((_, i) => _weekendDaysIndices.Contains(i)).Sum();
        //int minWeekendShifts = weekendShifts / _numPeople;
        //int maxWeekendShifts = minWeekendShifts + 1;

        //for (int i = 0; i < _numPeople; i++)
        //{
        //    model.AddLinearConstraint(LinearExpr.Sum(weekendWorkload[i]), minWeekendShifts, maxWeekendShifts);
        //}

        if (_balanceShiftsAutomatically)
        {
            // Constraint 5: Balance total shifts equally among people with ±1 difference
            int totalMinShifts = _totalShifts / _numPeople;
            int totalMaxShifts = totalMinShifts + 1;

            for (int i = 0; i < _numPeople; i++)
            {
                var totalShiftsPerPerson = new List<ILiteral>();
                for (int j = 0; j < _numDays; j++)
                {
                    totalShiftsPerPerson.Add(shifts[i, j]);
                }
                model.AddLinearConstraint(LinearExpr.Sum(totalShiftsPerPerson), totalMinShifts, totalMaxShifts);
            }
        }
        else
        {
            // Constraint 5: Ensure each person gets the specified number of shifts
            for (int i = 0; i < _numPeople; i++)
            {
                // Create an expression that sums up all the shifts this person has
                List<ILiteral> personShifts = new List<ILiteral>();
                for (int j = 0; j < _numDays; j++)
                {
                    personShifts.Add(shifts[i, j]);
                }

                // Add a constraint to ensure the exact number of shifts for the person
                model.Add(LinearExpr.Sum(personShifts) == _desiredShiftCounts[i]);
            }
        }

        // Constraint 6: JUNIOR members require at least one SENIOR member on the same shift
        for(int j = 0; j < _numDays; j++)
{
            foreach (int junior in _juniorIndices)
            {
                // Collect all senior presence indicators for the given day
                List<ILiteral> seniorPresence = new List<ILiteral>();
                foreach (int senior in _seniorIndices)
                {
                    seniorPresence.Add(shifts[senior, j]);
                }

                // Create a condition that at least one SENIOR is present
                var atLeastOneSenior = new List<ILiteral>();
                atLeastOneSenior.AddRange(seniorPresence);
                atLeastOneSenior.Add(shifts[junior, j].Not()); // If JUNIOR is working, then add OR condition for at least one SENIOR

                // Ensure that if a JUNIOR is working, at least one SENIOR must be working
                model.AddBoolOr(atLeastOneSenior);
            }
        }

        // Constraint 7: Max 1 JUNIOR per day
        for (int j = 0; j < _numDays; j++)
        {
            // Collect all juniors working on day j
            List<ILiteral> juniorShiftsOnDay = new List<ILiteral>();
            foreach (int junior in _juniorIndices)
            {
                juniorShiftsOnDay.Add(shifts[junior, j]);
            }

            // Ensure that at most one junior works per day
            model.Add(LinearExpr.Sum(juniorShiftsOnDay) <= 1);
        }

        // Constraint 8: Create auxiliary variables for juniors working on general days
        List<BoolVar> juniorOnGeneralDay = new List<BoolVar>();

        foreach (int day in _isGeneralDaysIndices)
        {
            foreach (int junior in _juniorIndices)
            {
                // Track if the junior is assigned on a general day
                juniorOnGeneralDay.Add(shifts[junior, day]);
            }
        }

        //// Minimize the total violations for j+2 day
        //model.Minimize(totalViolations);
        //// Maximize junior presence on general days
        //model.Maximize(LinearExpr.Sum(juniorOnGeneralDay));

        // Step 1: Solve for minimum violations
        model.Minimize(totalViolations);
        CpSolver solver = new CpSolver();
        CpSolverStatus status = solver.Solve(model);
        int minViolations = (int)solver.ObjectiveValue;  // Store the best possible violations count

        //TODO MPEN-NEXT: add or remove the following validation in the future
        //----------------------------------------------------------------------------------------------------------------------------------
        // Step 2: Re-run with a constraint on totalViolations and maximize junior presence
        //model.Add(totalViolations == minViolations);  // Ensure we maintain optimal violations
        //model.Maximize(LinearExpr.Sum(juniorOnGeneralDay));
        //solver.Solve(model);
        //----------------------------------------------------------------------------------------------------------------------------------



        SolutionCollector solutionCollector = new SolutionCollector(shifts, _numPeople, _numDays);

        // Set solver parameters (optional)
        solver.StringParameters = "num_search_workers:8"; // Use multiple threads for faster search
        solver.StringParameters = "max_time_in_seconds:60"; // Stop after 60 seconds

        // Solve the model and collect solutions
        CpSolverStatus status2 = solver.Solve(model, solutionCollector);

        // Check results
        if (status2 == CpSolverStatus.Feasible || status2 == CpSolverStatus.Optimal)
        {
            List<int[,]> solutions = solutionCollector.GetSolutions();
            Console.WriteLine($"Found {solutions.Count} solutions.");

            // Display or save the solutions
            for (int s = 0; s < solutions.Count; s++)
            {
                Console.WriteLine($"Solution {s + 1}:");
                for (int i = 0; i < _numPeople; i++)
                {
                    for (int j = 0; j < _numDays; j++)
                    {
                        Console.Write(solutions[s][i, j] + " ");
                    }
                    Console.WriteLine();
                }
                var x = solutions[s];
                FinalizeSolutionNew(model, x, "solution_new_"+s);
                Console.WriteLine();
            }
        }
        else
        {
            Console.WriteLine("No feasible solution found.");
        }


        // Check results
        FinalizeSolution(model, shifts);
    }

    static void Initialize()
    {
        // Open the Excel workbook
        using (var workbook = new XLWorkbook(WORKBOOK))
        {
            // Access a specific worksheet by name
            var programWorksheet = workbook.Worksheet("program");
            var negativesWorksheet = workbook.Worksheet("negatives");

            _numDays = programWorksheet.Cell(35, 5).GetValue<int>();
            _allowB2BShifts = programWorksheet.Cell(35, 2).GetValue<bool>();
            _balanceShiftsAutomatically = programWorksheet.Cell(36, 2).GetValue<bool>();
            _totalShifts = programWorksheet.Cell(36, 5).GetValue<int>();
            var totalShiftPerPerson = programWorksheet.Cell(37, 5).GetValue<int>();
            if (totalShiftPerPerson != _totalShifts)
            {
                throw new Exception("'Total shifts' are not equal to 'Total shifts alocated per person'");
            }
            _numPeople = negativesWorksheet.ColumnsUsed().Count()-1;
            var rowsCount = _numDays + 1;
            //Setup shifts per day
            for (int row = 2; row <= rowsCount; row++)
            {
                var dayNumber = row - 2;
                var shiftsPerDay = programWorksheet.Cell(row, 2).GetValue<int>();
                shiftsPerDayList.Add(shiftsPerDay);
            }
            _totalShifts = shiftsPerDayList.Sum();

            //Setup weekends list
            for (int row = 2; row <= rowsCount; row++)
            {
                var dayNumber = row - 2;
                var isWeekend = programWorksheet.Cell(row, 3).GetValue<int>();
                if (isWeekend == 1)
                {
                    _weekendDaysIndices.Add(dayNumber);
                }
            }
            _weekendDays = _weekendDaysIndices.Count();

            //Setup isGeneral list
            for (int row = 2; row <= rowsCount; row++)
            {
                var dayNumber = row - 2;
                var isGeneral = programWorksheet.Cell(row, 4).GetValue<int>();
                if (isGeneral == 1)
                {
                    _isGeneralDaysIndices.Add(dayNumber);
                }
            }

            _firstDayName = programWorksheet.Cell(2, 5).GetValue<string>();

            //Setup people
            for (int col = 2; col <= _numPeople+1; col++)
            {
                _people.Add(negativesWorksheet.Cell(1, col).GetValue<string>());
            }

            //Setup juniors/seniors
            for (int col = 2; col <= _numPeople + 1; col++)
            {
                var personIndice = col - 2;
                var isJunior = negativesWorksheet.Cell(35, col).GetValue<bool>();
                if (isJunior)
                {
                    _juniorIndices.Add(personIndice);
                }
                else
                {
                    _seniorIndices.Add(personIndice);
                }
            }


            //Setup desired Shift Counts
            for (int col = 2; col <= _numPeople + 1; col++)
            {
                var shiftCount = negativesWorksheet.Cell(36, col).GetValue<int>();
                _desiredShiftCounts.Add(shiftCount);
            }

            //Setup negatives for each person
            for (int col = 2; col <= _numPeople+1; col++)
            {
                var negativeDaysList = new List<int> { };
                for (int row = 2; row <= rowsCount; row++)
                {
                    var value = negativesWorksheet.Cell(row, col).GetValue<int>();
                    if (value == 1)
                    {
                        negativeDaysList.Add(row - 2);
                    }
                }
                _unavailableDays.Add(col - 2, negativeDaysList);
            }
        }
    }

    static void FinalizeSolution(CpModel model, BoolVar[,] shifts)
    {
        // Solve the model
        CpSolver solver = new CpSolver();
        var status = solver.Solve(model);
        var weekendCount = new List<int>();
        var totalCount = new List<int>();
        for (var i = 0; i <= _people.Count(); i++)
        {
            weekendCount.Add(0);
            totalCount.Add(0);
        }

        using (var workbook = new XLWorkbook(WORKBOOK))
        {
            // Access a specific worksheet by name
            var worksheet = workbook.Worksheet("solution");

            if (status == CpSolverStatus.Feasible || status == CpSolverStatus.Optimal)
            {
                for (int j = 2; j < 33; j++)
                {
                    for (int i = 2; i < _numPeople+2; i++)
                    {
                        worksheet.Cell(i, j).Value = ""; //Clear previous value
                        worksheet.Cell(i, j).Style.Fill.BackgroundColor = XLColor.NoColor; //clear previous color
                    }
                }
                var firstDayIndice = FindFirstDayIndice(_firstDayName);
                for (int j = 0; j < _numDays; j++)
                {
                    var col = j + 2;
                    var weekendText = _weekendDaysIndices.Contains(j) ? "(weekend)" : "";
                    Console.Write($"Day {j + 1}{weekendText}: ");

                    var currentDayIndice = (j + firstDayIndice) % (DAY_NAMES.Count);
                    //Add label for each day
                    worksheet.Cell(1, col).Value = DAY_NAMES[currentDayIndice];
                    //Mark day as when isGeneral
                    if (_isGeneralDaysIndices.Contains(j))
                    {
                        worksheet.Cell(1, col).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }
                    //Mark days for each person
                    for (int i = 0; i < _numPeople; i++)
                    {
                        var row = i + 2;
                        var hasShift = solver.BooleanValue(shifts[i, j]);
                        var hasDayOff = _unavailableDays[i].Contains(j);
                        if (hasShift && hasDayOff)
                        {
                            throw new Exception($"Shift schduled on day off for {_people[i]} on day {j + 1}");
                        } 
                        else if (hasShift)
                        {
                            worksheet.Cell(row, col).Value = 1;
                            if (_weekendDaysIndices.Contains(j))
                            {
                                weekendCount[i]++;
                            }
                            totalCount[i]++;
                            Console.Write($"{_people[i]} ");
                        }
                        else if (hasDayOff)
                        {
                            worksheet.Cell(row, col).Value = "X";
                        }
                    }
                    Console.WriteLine();
                }
                Console.WriteLine();
                var totalCol = 34;
                var totalWeekendCol = 35;
                worksheet.Cell(1, totalCol).Value = "Total";
                worksheet.Cell(1, totalWeekendCol).Value = "Total Weekends";
                for (var i = 0; i < _numPeople; i++)
                {
                    worksheet.Cell(i + 2, totalWeekendCol).Value = weekendCount[i];
                    Console.WriteLine($"Weekend count for {_people[i]}: {weekendCount[i]}");
                }
                Console.WriteLine();
                for (var i = 0; i < _numPeople; i++)
                {
                    worksheet.Cell(i + 2, totalCol).Value = totalCount[i];
                    Console.WriteLine($"Total count for {_people[i]}: {totalCount[i]}");
                }
                workbook.Save();
            }
            else
            {
                Console.WriteLine("No feasible solution found.");
            }
        }
    }


    static void FinalizeSolutionNew(CpModel model, int[,] solution, string sheet)
    {
        // Solve the model
        CpSolver solver = new CpSolver();
        var status = solver.Solve(model);
        var weekendCount = new List<int>();
        var totalCount = new List<int>();
        for (var i = 0; i <= _people.Count(); i++)
        {
            weekendCount.Add(0);
            totalCount.Add(0);
        }

        using (var workbook = new XLWorkbook(WORKBOOK))
        {
            // Access a specific worksheet by name
            var worksheet = workbook.Worksheet(sheet);

            if (status == CpSolverStatus.Feasible || status == CpSolverStatus.Optimal)
            {
                for (int j = 2; j < 33; j++)
                {
                    for (int i = 2; i < _numPeople + 2; i++)
                    {
                        worksheet.Cell(i, j).Value = ""; //Clear previous value
                        worksheet.Cell(i, j).Style.Fill.BackgroundColor = XLColor.NoColor; //clear previous color
                    }
                }
                var firstDayIndice = FindFirstDayIndice(_firstDayName);
                for (int j = 0; j < _numDays; j++)
                {
                    var col = j + 2;
                    var weekendText = _weekendDaysIndices.Contains(j) ? "(weekend)" : "";
                    Console.Write($"Day {j + 1}{weekendText}: ");

                    var currentDayIndice = (j + firstDayIndice) % (DAY_NAMES.Count);
                    //Add label for each day
                    worksheet.Cell(1, col).Value = DAY_NAMES[currentDayIndice];
                    //Mark day as when isGeneral
                    if (_isGeneralDaysIndices.Contains(j))
                    {
                        worksheet.Cell(1, col).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }
                    //Mark days for each person
                    for (int i = 0; i < _numPeople; i++)
                    {
                        var row = i + 2;
                        var hasShift = solution[i, j] == 1;
                        var hasDayOff = _unavailableDays[i].Contains(j);
                        if (hasShift && hasDayOff)
                        {
                            throw new Exception($"Shift schduled on day off for {_people[i]} on day {j + 1}");
                        }
                        else if (hasShift)
                        {
                            worksheet.Cell(row, col).Value = 1;
                            if (_weekendDaysIndices.Contains(j))
                            {
                                weekendCount[i]++;
                            }
                            totalCount[i]++;
                            Console.Write($"{_people[i]} ");
                        }
                        else if (hasDayOff)
                        {
                            worksheet.Cell(row, col).Value = "X";
                        }
                    }
                    Console.WriteLine();
                }
                Console.WriteLine();
                var totalCol = 34;
                var totalWeekendCol = 35;
                worksheet.Cell(1, totalCol).Value = "Total";
                worksheet.Cell(1, totalWeekendCol).Value = "Total Weekends";
                for (var i = 0; i < _numPeople; i++)
                {
                    worksheet.Cell(i + 2, totalWeekendCol).Value = weekendCount[i];
                    Console.WriteLine($"Weekend count for {_people[i]}: {weekendCount[i]}");
                }
                Console.WriteLine();
                for (var i = 0; i < _numPeople; i++)
                {
                    worksheet.Cell(i + 2, totalCol).Value = totalCount[i];
                    Console.WriteLine($"Total count for {_people[i]}: {totalCount[i]}");
                }
                workbook.Save();
            }
            else
            {
                Console.WriteLine("No feasible solution found.");
            }
        }
    }



    private static int FindFirstDayIndice(string dayName)
    {
        // Define an array with the days of the week starting from Monday
        string[] daysOfWeek = { "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday" };

        // Convert the day to lowercase for case-insensitive comparison
        dayName = dayName.ToLower();

        // Loop through the array to find the index of the given day
        for (int i = 0; i < daysOfWeek.Length; i++)
        {
            if (daysOfWeek[i].ToLower() == dayName)
            {
                return i; // Return the index
            }
        }

        // Return -1 if the day is not found
        return -1;
    }
}
