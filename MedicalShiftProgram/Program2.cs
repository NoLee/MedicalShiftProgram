using System;
using System.Collections.Generic;
using System.Linq;

class ShiftScheduler
{
    static List<string> people = new List<string> { "Alice", "Bob", "Charlie", "David", "Eva", "Frank", "Grace", "Helen" };

    // Days off mapping for each person
    static Dictionary<string, List<int>> daysOff = new Dictionary<string, List<int>>()
    {
        { "Alice", new List<int> { 3, 7, 14,15,23 } },
        { "Bob", new List<int> { 2, 8, 15,17,29 } },
        { "Charlie", new List<int> { 1,2,3, 9, 17 } },
        { "David", new List<int> { 5,8,9, 10, 13 } },
        { "Eva", new List<int> { 3, 6, 8,15,16, 20 } },
        { "Frank", new List<int> { 4, 12,13,14,15,16, 18 } },
        { "Grace", new List<int> { 7, 16, 21,22,23,24,25,26 } },
        { "Helen", new List<int> { 1,2,12, 14, 22 } }
    };

    // Define weekends (Saturday and Sunday) for this example
    static List<int> weekends = new List<int> { 6, 7, 13, 14, 20, 21, 27, 28 };

    // Track weekend shifts per person
    static Dictionary<string, int> weekendCount = people.ToDictionary(person => person, person => 0);

    // Initialize schedule dictionary (1-30 days)
    static Dictionary<int, List<string>> schedule = new Dictionary<int, List<string>>();

    static void Main2()
    {
        for (int day = 1; day <= 30; day++)
        {
            schedule[day] = new List<string>();
        }

        // Start the shift assignment using backtracking
        if (AssignShifts(1))
        {
            // Print the successful schedule
            for (int day = 1; day <= 30; day++)
            {
                Console.WriteLine($"Day {day}: {string.Join(", ", schedule[day])}");
            }

            
            foreach (string person in people)
            {
                Console.WriteLine(person + ": "+ weekendCount[person]);
            }
        }
        else
        {
            Console.WriteLine("No valid schedule could be found.");
        }
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
        if (day > 2 && schedule[day - 2].Contains(person))
            return false;

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

        // Try assigning two different people for the current day
        var candidates = people.OrderBy(x => Guid.NewGuid()).ToList(); // Randomize to add variation

        foreach (var firstPerson in candidates)
        {
            if (!CanWork(firstPerson, day))
                continue;

            foreach (var secondPerson in candidates)
            {
                if (firstPerson == secondPerson || !CanWork(secondPerson, day))
                    continue;

                // Check for weekend balancing
              if (weekends.Contains(day))
                {
                    //find min weekend person
                    // ta dika mou na einai TOULAXISTON <= min+1
                    int minCount = weekendCount.Values.Min() +1;
                    if (weekendCount[firstPerson] > minCount ||
                        weekendCount[secondPerson] > minCount)
                    {
                        continue;
                    }
/*                    if (weekendCount[firstPerson] >= (weekends.Count / people.Count) ||
                        weekendCount[secondPerson] >= (weekends.Count / people.Count))
                    {
                        continue;
                    }*/
                }

                // Assign the two people to this day
                schedule[day].Add(firstPerson);
                schedule[day].Add(secondPerson);

                // Adjust weekend count
                if (weekends.Contains(day))
                {
                    weekendCount[firstPerson]++;
                    weekendCount[secondPerson]++;
                }

                // Recur for the next day
                if (AssignShifts(day + 1))
                {
                    return true; // Successful assignment
                }

                // Backtrack: Remove the assignment and reset counts if needed
                schedule[day].Remove(firstPerson);
                schedule[day].Remove(secondPerson);
                if (weekends.Contains(day))
                {
                    weekendCount[firstPerson]--;
                    weekendCount[secondPerson]--;
                }
            }
        }

        // No valid assignment found for this day, backtrack
        return false;
    }
}
