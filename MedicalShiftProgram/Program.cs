using ClosedXML.Excel;
using Google.OrTools.Sat;
using System;
using System.Collections.Generic;
using System.Linq;

class ShiftScheduler
{
    // Define constants
    static string WORKBOOK = "C:\\Users\\NoLee\\source\\repos\\MedicalShiftProgram\\nosokomeio.xlsx";
    static List<int> shiftsPerDayList = new List<int> { };
    static int _weekendDays;
    static int _numPeople;
    static int _numDays;
    static int _totalShifts;
    static bool _allowB2BShifts;
    static bool _balanceShiftsAutomatically;

    static List<int> _weekendDaysIndices = new List<int> { };
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
                // If person `i` works on day `j`, then they must not work on `j+1` or `j+2`
                model.AddBoolOr(new ILiteral[] { shifts[i, j].Not(), shifts[i, j + 1].Not() });
                if (_allowB2BShifts == false)
                {
                    model.AddBoolOr(new ILiteral[] { shifts[i, j].Not(), shifts[i, j + 2].Not() });
                }
            }
            //Edge case for last day
            model.AddBoolOr(new ILiteral[] { shifts[i, _numDays -1].Not(), shifts[i, _numDays - 2].Not() });
        }

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
        var weekendWorkload = new List<ILiteral>[_numPeople];
        for (int i = 0; i < _numPeople; i++)
        {
            weekendWorkload[i] = new List<ILiteral>();
            foreach (var day in _weekendDaysIndices)
            {
                weekendWorkload[i].Add(shifts[i, day]);
            }
        }
        var weekendShifts = shiftsPerDayList.Where((_, i) => _weekendDaysIndices.Contains(i)).Sum();
        int minWeekendShifts = weekendShifts / _numPeople;
        int maxWeekendShifts = minWeekendShifts + 1;

        for (int i = 0; i < _numPeople; i++)
        {
            model.AddLinearConstraint(LinearExpr.Sum(weekendWorkload[i]), minWeekendShifts, maxWeekendShifts);
        }

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
                    }
                }

                for (int j = 0; j < _numDays; j++)
                {
                    var col = j + 2;
                    var weekendText = _weekendDaysIndices.Contains(j) ? "(weekend)" : "";
                    Console.Write($"Day {j + 1}{weekendText}: ");
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
}
