using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Encodings.Web;
using System.Text.Json;

class ShiftScheduler3
{
    static List<string> people = new List<string> { };

    // Days off mapping for each person
    static Dictionary<string, List<int>> daysOff = new Dictionary<string, List<int>>() {};

    // Define weekends (Saturday and Sunday) for this example
    static List<int> weekends = new List<int> { };

    // Track weekend shifts per person
    static Dictionary<string, int> weekendCount = people.ToDictionary(person => person, person => 0);

    // Specify the number of workers needed per day (e.g., most days need 2, some need 3)
    static List<int> workersPerDay = new List<int> { };

    static int daysInMonth = 0;

    static int countTotalTest = 0;

    // Initialize schedule dictionary (1-30 days)
    static Dictionary<int, List<string>> schedule = new Dictionary<int, List<string>>();

    static void Main1()
    {
        Initialize();
        var options = new JsonSerializerOptions
        {
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            WriteIndented = true
        };

        string jsonString = JsonSerializer.Serialize(weekends, options);

        for (int day = 1; day <= 30; day++)
        {
            schedule[day] = new List<string>();
        }
        int minCount = 0;
        int maxCount = 100;
        int minCountTotal = 0;
        int maxCountTotal = 100;
        var daysTotalForPerson = new Dictionary<string, int> { };
        //Try to have 1 diff for weekends
        while ((maxCount > minCount +2) || (maxCountTotal > minCountTotal + 2))
        {
            schedule.Clear();
            daysTotalForPerson.Clear();
            foreach (string person in people)
            {
                daysTotalForPerson.Add(person, 0);
            }
            weekendCount = people.ToDictionary(person => person, person => 0);
            var isValid = AssignShifts(1);
            if (isValid == false)
            {
                Console.WriteLine("No valid schedule could be found.");
                return;
            }
            for (int day = 1; day <= 30; day++)
            {
                foreach(string person in schedule[day])
                {
                    daysTotalForPerson[person]++;
                }
            }
            minCount = weekendCount.Values.Min();
            maxCount = weekendCount.Values.Max();

            minCountTotal = daysTotalForPerson.Values.Min();
            maxCountTotal = daysTotalForPerson.Values.Max();
        }

        // Start the shift assignment using backtracking
        //if (AssignShifts(1))
        //{
            // Print the successful schedule
       

        for (int day = 1; day <= 30; day++)
        {
            Console.WriteLine($"Day {day}: {string.Join(", ", schedule[day])}");
        }

        Console.WriteLine("");
        foreach (string person in people)
        {
            Console.WriteLine("Total days for " + person + ": " + daysTotalForPerson[person]);
        }
        Console.WriteLine("");
        foreach (string person in people)
        {
            Console.WriteLine("Weekend days for " + person + ": " + weekendCount[person]);
        }

        Console.WriteLine("Total tries " + countTotalTest);
    }

    // Method to check if a person can work on a given day
    static bool CanWork(string person, int day)
    {
        // Check if the person has the day off
        if (daysOff.ContainsKey(person) && daysOff[person].Contains(day))
            return false;

        // Check if they worked the previous day
        if (day > 1 && schedule[day - 1].Contains(person))
            return false;

        // Check if they worked the other previous day
        //only sometimes
/*        Random random = new Random();
        int randomNumber = random.Next(3);
        if (randomNumber == 1 && day > 2 && schedule[day - 2].Contains(person))
            return false;*/

        return true;
    }

    // Backtracking function to assign shifts
    static bool AssignShifts(int day)
    {
        // If we have reached past the last day, return true (successful assignment)
        if (day > 30)
        {
            return true;
        }

        // Try assigning the required number of people for the current day
        //people.Where(x => CanWork(x, day))
        var candidates = people.OrderBy(x => Guid.NewGuid()).ToList(); // Randomize to add variation

        return TryAssignPeople(day, 0, new List<string>(), candidates);
    }

    // Helper function to recursively assign required workers for the day
    static bool TryAssignPeople(int day, int assignedCount, List<string> assignedPeople, List<string> candidates)
    {
        // Base case: if we have assigned the required number of people for the day
        if (assignedCount == workersPerDay[day - 1])
        {
            schedule[day] = new List<string>(assignedPeople);

            // Adjust weekend counts
            foreach (var person in assignedPeople)
            {
                if (weekends.Contains(day))
                {
                    weekendCount[person]++;
                }
            }

            // Recur for the next day
            if (AssignShifts(day + 1))
            {
                return true; // Successful assignment
            }

            // Backtrack: Remove the assignments
            foreach (var person in assignedPeople)
            {
                if (weekends.Contains(day))
                {
                    weekendCount[person]--;
                }
            }

            return false; // Backtracking
        }

        // Try to assign more people from the candidate list
        foreach (var person in candidates)
        {
            if (!CanWork(person, day) || assignedPeople.Contains(person))
                continue;

            // Check for weekend balancing
/*            if (weekends.Contains(day) )
            {
                int minCount = weekendCount.Values.Min() + 1;
                if (weekendCount[person] > minCount)
                {
                    continue;
                }
            }*/

            // Add the person to the assigned list and recurse
            assignedPeople.Add(person);
            countTotalTest++;
            if (TryAssignPeople(day, assignedCount + 1, assignedPeople, candidates))
            {
                return true;
            }

            // Backtrack: Remove the person if it didn't work out
            assignedPeople.Remove(person);
        }

        // No valid assignment found for this day, backtrack
        return false;
    }

    static void Initialize()
    {
        // Open the Excel workbook
        using (var workbook = new XLWorkbook("C:\\Users\\NoLee\\source\\repos\\MedicalShiftProgram\\nosokomeio.xlsx"))
        {
            // Access a specific worksheet by name
            var programWorksheet = workbook.Worksheet("program");
            var negativesWorksheet = workbook.Worksheet("negatives");

            daysInMonth = programWorksheet.Cell(1,1).GetValue<int>();
            var rowsCount = daysInMonth + 1;
            var columnCountNegatives = negativesWorksheet.ColumnsUsed().Count();

            //Setup weekends list
            for (int row=2; row<= rowsCount; row++)
            {
                var dayNumber = row - 1;
                var isWeekend = programWorksheet.Cell(row, 3).GetValue<int>();
                if (isWeekend == 1)
                {
                    weekends.Add(dayNumber);
                }
            }

            //Setup workersPerDay array
            for (int row = 2; row <= rowsCount; row++)
            {
                workersPerDay.Add(programWorksheet.Cell(row, 2).GetValue<int>());
            }

            //Setup people
            for (int col = 2; col <= columnCountNegatives; col++)
            {
                people.Add(negativesWorksheet.Cell(1, col).GetValue<string>());
            }

            //Setup negatives for each person
            for (int col = 2; col <= columnCountNegatives; col++)
            {
                var negativeDaysList = new List<int> { };
                for (int row = 2; row <= rowsCount; row++)
                {
                    var value = negativesWorksheet.Cell(row, col).GetValue<int>();
                    if (value == 1)
                    {
                        negativeDaysList.Add(row - 1);
                    }
                }
                daysOff.Add(people[col-2], negativeDaysList);
            }

            weekendCount = people.ToDictionary(person => person, person => 0);
        }
    }
}
