using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace CorrectContraAccountLogicDLL
{
    public class Solver
    {
        private readonly int _roundAccuracy;
        private readonly Company _company;

        public Solver(Company company)
        {
            _company = company;
            _roundAccuracy = company.GetCompanyService().GetAdminInfo().TotalsAccuracy;
        }

        Stopwatch _stopwatch;
        public List<JournalEntryLineModel> SolveCombinations(JournalEntryLineModel targetLine, List<JournalEntryLineModel> searchLines, int milliseconds)
        {
            _stopwatch = new Stopwatch();
            _stopwatch.Start();
            _sourceLines = new List<JournalEntryLineModel>();
            RecursiveSolveCombinations(targetLine, 0, new List<JournalEntryLineModel>(), searchLines, 0, milliseconds);
            return _sourceLines;
        }

        public List<JournalEntryLineModel> SolveCombinationsNegative(JournalEntryLineModel targetLine, List<JournalEntryLineModel> searchLines)
        {
            _sourceLines = new List<JournalEntryLineModel>();
            RecursiveSolveCombinationsNegative(targetLine, 0, new List<JournalEntryLineModel>(), searchLines, 0);
            return _sourceLines;
        }

        private List<JournalEntryLineModel> _sourceLines;

        private void RecursiveSolveCombinations(JournalEntryLineModel targetLine, double currentSum, List<JournalEntryLineModel> included, List<JournalEntryLineModel> notIncluded, int startIndex, int milliseconds)
        {
            for (int index = startIndex; index < notIncluded.Count; index++)
            {
                if (_stopwatch.ElapsedMilliseconds > milliseconds)
                {
                    return;
                }
                double goal = targetLine.Debit == 0 ? targetLine.Credit : targetLine.Debit;
                JournalEntryLineModel nextLine = notIncluded[index];
                double nextAmount = nextLine.Debit == 0 ? nextLine.Credit : nextLine.Debit;
                double amountToCompare = Math.Round(currentSum + nextAmount, _roundAccuracy);

                if (amountToCompare == goal)
                {
                    List<JournalEntryLineModel> newResult = new List<JournalEntryLineModel>(included) { nextLine };
                    _sourceLines = newResult;
                    return;
                }
                else if (Math.Abs(amountToCompare) < Math.Abs(goal))
                {
                    List<JournalEntryLineModel> nextIncuded = new List<JournalEntryLineModel>(included) { nextLine };
                    List<JournalEntryLineModel> nextNonIncluded = new List<JournalEntryLineModel>(notIncluded);
                    nextNonIncluded.Remove(nextLine);
                    RecursiveSolveCombinations(targetLine, amountToCompare, nextIncuded, nextNonIncluded, startIndex++, milliseconds);
                }
            }
        }


        private void RecursiveSolveCombinationsNegative(JournalEntryLineModel targetLine, double currentSum, List<JournalEntryLineModel> included, List<JournalEntryLineModel> notIncluded, int startIndex)
        {
            var roundTotalsAccuracy = _company.GetCompanyService().GetAdminInfo().TotalsAccuracy;
            for (int index = startIndex; index < notIncluded.Count; index++)
            {
                double goal = targetLine.Debit == 0 ? targetLine.Credit : targetLine.Debit;
                JournalEntryLineModel nextLine = notIncluded[index];
                double nextAmount = nextLine.Debit == 0 ? nextLine.Credit : nextLine.Debit;
                double amountToCompare = Math.Round(currentSum + nextAmount, roundTotalsAccuracy);

                if (amountToCompare + goal == 0)
                {
                    List<JournalEntryLineModel> newResult = new List<JournalEntryLineModel>(included) { nextLine };
                    _sourceLines = newResult;
                }
                else if (Math.Abs(amountToCompare) < Math.Abs(goal))
                {
                    List<JournalEntryLineModel> nextIncuded = new List<JournalEntryLineModel>(included) { nextLine };
                    List<JournalEntryLineModel> nextNonIncluded = new List<JournalEntryLineModel>(notIncluded);
                    nextNonIncluded.Remove(nextLine);
                    RecursiveSolveCombinationsNegative(targetLine, amountToCompare, nextIncuded, nextNonIncluded, startIndex++);
                }
            }
        }


        private List<Dictionary<int, double>> _mResults;
        public List<Dictionary<int, double>> Solve(double goal, Dictionary<int, double> elements)
        {
            _mResults = new List<Dictionary<int, double>>();
            RecursiveSolve(goal, 0, new Dictionary<int, double>(), new Dictionary<int, double>(elements), 0);
            return _mResults;
        }

        private void RecursiveSolve(double goal, double currentSum,
            Dictionary<int, double> included, Dictionary<int, double> notIncluded, int startIndex)
        {
            var roundTotalsAccuracy = _company.GetCompanyService().GetAdminInfo().TotalsAccuracy;

            for (int index = startIndex; index < notIncluded.Count; index++)
            {
                double nextValue = notIncluded.Values.ElementAt(index);
                double coparerSum = Math.Round(currentSum + nextValue, roundTotalsAccuracy);
                if (coparerSum == goal)
                {
                    Dictionary<int, double> newResult = new Dictionary<int, double>(included) { { notIncluded.First().Key, nextValue } };
                    _mResults.Add(newResult);
                }
                else if (currentSum + nextValue < goal)
                {
                    Dictionary<int, double> nextIncluded = new Dictionary<int, double>(included)
                    {
                        {notIncluded.First().Key, nextValue}
                    };
                    Dictionary<int, double> nextNotIncluded = new Dictionary<int, double>(notIncluded);
                    nextNotIncluded.Remove(notIncluded.First().Key);
                    RecursiveSolve(goal, currentSum + nextValue,
                        nextIncluded, nextNotIncluded, startIndex++);
                }
            }
        }
    }
}
