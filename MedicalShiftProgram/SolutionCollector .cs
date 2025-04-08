using Google.OrTools.Sat;
using System.Collections.Generic;

public class SolutionCollector : CpSolverSolutionCallback
{
    private List<int[,]> _solutions;
    private BoolVar[,] _shifts;
    private int _numPeople;
    private int _numDays;

    public SolutionCollector(BoolVar[,] shifts, int numPeople, int numDays)
    {
        _solutions = new List<int[,]>();
        _shifts = shifts;
        _numPeople = numPeople;
        _numDays = numDays;
    }

    public override void OnSolutionCallback()
    {
        // Store the current solution
        int[,] solution = new int[_numPeople, _numDays];
        for (int i = 0; i < _numPeople; i++)
        {
            for (int j = 0; j < _numDays; j++)
            {
                solution[i, j] = Value(_shifts[i, j]) == 1 ? 1 : 0;
            }
        }
        _solutions.Add(solution);

        // Stop searching after a certain number of solutions (e.g., 10)
        if (_solutions.Count >= 10)
        {
            StopSearch();
        }
    }

    public List<int[,]> GetSolutions()
    {
        return _solutions;
    }
}