using ClosedXML.Excel;
using Google.OrTools.Sat;
using System;
using System.Collections.Generic;
using System.Linq;

class ShiftScheduler
{
    // Define constants
    static int SHIFTS_PER_DAY = 3;
    static int _weekendDays; // Let's assume Saturday and Sunday are the last two days
    static int _numPeople;
    static int _numDays;
    static int _totalShifts;
    static List<int> _weekendDaysIndices = new List<int> { };
    static List<string> _people = new List<string> { };
    static Dictionary<int, List<int>> _unavailableDays = new Dictionary<int, List<int>> { };

    static void Main()
    {
        Initialize();
        // Create a CP-SAT model
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

        // Constraint 1: Each day must have exactly 3 people working
        for (int j = 0; j < _numDays; j++)
        {
            var dailyWorkers = new List<ILiteral>();
            for (int i = 0; i < _numPeople; i++)
            {
                dailyWorkers.Add(shifts[i, j]);
            }
            model.Add(LinearExpr.Sum(dailyWorkers) == SHIFTS_PER_DAY);
        }

        // Constraint 2: No one should work two consecutive days
        for (int i = 0; i < _numPeople; i++)
        {
            for (int j = 0; j < _numDays - 1; j++)
            {
                model.AddBoolOr(new ILiteral[] { shifts[i, j].Not(), shifts[i, j + 1].Not() });
            }
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

        var weekendWorkload = new List<ILiteral>[_numPeople];
        for (int i = 0; i < _numPeople; i++)
        {
            weekendWorkload[i] = new List<ILiteral>();
            foreach (var day in _weekendDaysIndices)
            {
                weekendWorkload[i].Add(shifts[i, day]);
            }
        }
        int minWeekendShifts = (_weekendDays * SHIFTS_PER_DAY) / _numPeople;
        int maxWeekendShifts = minWeekendShifts + 1;

        for (int i = 0; i < _numPeople; i++)
        {
            model.AddLinearConstraint(LinearExpr.Sum(weekendWorkload[i]), minWeekendShifts, maxWeekendShifts);
        }

        //// New Constraint 5: Everyone works approximately the same number of days
        //// Calculate the minimum and maximum shifts each person can work
        int minShifts = _totalShifts / _numPeople;       // Minimum shifts a person should work
        int maxShifts = (_totalShifts + _numPeople - 1) / _numPeople; // Maximum shifts (ceil)
        var workload = new List<ILiteral>[_numPeople];
        for (int i = 0; i < _numPeople; i++)
        {
            workload[i] = new List<ILiteral>();
            for (int j = 0; j < _numDays - 1; j++)
            {
                workload[i].Add(shifts[i, j]);
            }
        }

        for (int i = 0; i < _numPeople; i++)
        {
            model.AddLinearConstraint(LinearExpr.Sum(workload[i]), minShifts, maxShifts);
        }

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
        // Check results
        if (status == CpSolverStatus.Feasible || status == CpSolverStatus.Optimal)
        {
            for (int j = 0; j < _numDays; j++)
            {
                var x = _weekendDaysIndices.Contains(j) ? "(weekend)" : "";
                Console.Write($"Nov {j + 1}{x}: ");
                for (int i = 0; i < _numPeople; i++)
                {
                    if (solver.BooleanValue(shifts[i, j]))
                    {
                        if (_weekendDaysIndices.Contains(j))
                        {
                            weekendCount[i]++;
                        }
                        totalCount[i]++;
                        Console.Write($"{_people[i]} ");
                    }
                }
                Console.WriteLine();
            }
            Console.WriteLine();
            for (var i=0; i<_people.Count();i++)
            {
                Console.WriteLine($"Weekend count for {_people[i]}: {weekendCount[i]}");
            }
            Console.WriteLine();
            for (var i = 0; i < _people.Count(); i++)
            {
                Console.WriteLine($"Total count for {_people[i]}: {totalCount[i]}");
            }
        }
        else
        {
            Console.WriteLine("No feasible solution found.");
        }
    }

    static void Initialize()
    {
        // Open the Excel workbook
        using (var workbook = new XLWorkbook("C:\\Users\\NoLee\\source\\repos\\MedicalShiftProgram\\nosokomeio.xlsx"))
        {
            // Access a specific worksheet by name
            var programWorksheet = workbook.Worksheet("program");
            var negativesWorksheet = workbook.Worksheet("negatives");

            _numDays = programWorksheet.Cell(1, 1).GetValue<int>();
            _totalShifts = _numDays * SHIFTS_PER_DAY;
            _numPeople = negativesWorksheet.ColumnsUsed().Count()-1;
            var rowsCount = _numDays + 1;
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

            //Setup workersPerDay array
            //for (int row = 2; row <= rowsCount; row++)
            //{
            //    workersPerDay.Add(programWorksheet.Cell(row, 2).GetValue<int>());
            //}

            //Setup people
            for (int col = 2; col <= _numPeople+1; col++)
            {
                _people.Add(negativesWorksheet.Cell(1, col).GetValue<string>());
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

            //weekendCount = people.ToDictionary(person => person, person => 0);
        }
    }

    //var unavailableDays = new Dictionary<int, List<int>>
    //{
    //    { 0, new List<int> {  } },  // ΧΡΙΣΤΟΔΟΥΛΟΥ
    //    { 1, new List<int> {  } },  // ΦΙΛΙΠΠΟΥΣΗ
    //    { 2, new List<int> { } },  // ΜΠΙΝΙΣΚΟΥ
    //    { 3, new List<int> { 2 } },  // ΜΠΟΤΟΥ
    //    { 4, new List<int> {  } },  // ΣΟΥΛΙΜΑ
    //    { 5, new List<int> {  } },  // ΜΠΑΚΟΠΟΥΛΟΣ
    //    { 6, new List<int> { } },  // ΚΑΡΑΜΟΛΕΓΚΟΥ
    //    { 7, new List<int> { } },  // ΤΣΟΥΜΑ
    //    { 8, new List<int> {  } },  // ΜΟΝΟΠΑΤΗΣ
    //    { 9, new List<int> { } },  // ΠΡΟΔΡΟΜΑΚΗΣ
    //    { 10, new List<int> {  } },  // ΑΥΓΕΡΙΝΟΥ
    //    { 11, new List<int> { 22, 23 } },  // ΜΠΡΑΙΜΑΚΗΣ
    //    { 12, new List<int> { } },  // ΠΑΠΑΚΩΝΣΤΑΝΤΙΝΟΥ
    //    { 13, new List<int> {  } }  // ΠΑΝΑΓΟΥΛΑ
    //};
}
