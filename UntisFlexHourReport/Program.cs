using ClosedXML.Excel;
using UntisFlexHourReport;
using System.Diagnostics;

Console.WriteLine("== BAUEs Untis-Auswertung ==");

try
{
    Console.Write("Pfad zum Untis-Bericht: ");
    string inputPath = Console.ReadLine()?.Trim('"') ?? "";
    string outputPath = $"{Path.GetFileNameWithoutExtension(inputPath)}_mit_Auswertung{Path.GetExtension(inputPath)}";
    string fullOutputPath = Path.Combine(Path.GetDirectoryName(inputPath)!, outputPath);
    File.Copy(inputPath, fullOutputPath, overwrite: true);
    List<Teacher> teacherList = ReadUntisReport(fullOutputPath);
    AddUntisReportSummary(fullOutputPath, teacherList);
    Process.Start(new ProcessStartInfo(fullOutputPath) { UseShellExecute = true });
}
catch (FileNotFoundException e)
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine($"FEHLER: \"{e.FileName}\" konnte nicht gefunden werden.");
    Console.ResetColor();
    Console.ReadLine();
}
catch (Exception e)
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine($"FEHLER: {e.Message}");
    Console.ResetColor();
    Console.ReadLine();
}

static List<Teacher> ReadUntisReport(string path)
{
    using var wb = new XLWorkbook(path);
    var ws = wb.Worksheet(1);

    List<Teacher> teacherList = new();
    var startRow = ws.FirstRow();
    while (startRow.RowNumber() < ws.LastRowUsed().RowNumber())
    {
        var teacherNameRow = FindNextTeacherRow(startRow);
        var shortName = teacherNameRow.Cell(1).GetString();
        var firstName = teacherNameRow.Cell(4).GetString();
        var lastName = teacherNameRow.Cell(2).GetString();

        var actualHours = 0m;
        var tableHeaderRow = teacherNameRow.RowBelow(3);
        var actualHourCell = tableHeaderRow.Cells(c => c.GetString().Equals("Realstunden", StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
        if (actualHourCell == null)
        {
            throw new Exception($"Can't find cell \"Realstunden\" for teacher {shortName}");
        }

        var fUpisCell = tableHeaderRow.Cells(c => c.GetString().Equals("F-Upis", StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
        if (fUpisCell == null)
        {
            throw new Exception($"Can't find cell \"F-Upis\" for teacher {shortName}");
        }

        actualHourCell = actualHourCell.CellBelow();
        fUpisCell = fUpisCell.CellBelow();
        while (!actualHourCell.IsEmpty())
        {
            if (IsLessonIdentifier(fUpisCell.GetString()))
            {
                actualHours += decimal.Parse(actualHourCell.GetString());
            }
            actualHourCell = actualHourCell.CellBelow();
            fUpisCell = fUpisCell.CellBelow();
        }
        var teacher = new Teacher(shortName, firstName, lastName, actualHours);
        teacherList.Add(teacher);

        startRow = actualHourCell.WorksheetRow().RowBelow(4);
    }

    return teacherList;

    static IXLRow FindNextTeacherRow(IXLRow row)
    {
        while (!IsTeacherShortName(row.Cell(1).GetString()))
        {
            row = row.RowBelow();
        }
        return row;
    }

    static bool IsTeacherShortName(string value)
    {
        return value.Length == 4 && value.All(c => char.IsUpper(c));
    }

    static bool IsLessonIdentifier(string text)
    {
        return !text.Equals("R", StringComparison.InvariantCultureIgnoreCase);
    }
}

static void AddUntisReportSummary(string path, List<Teacher> teacherList)
{
    using var wb = new XLWorkbook(path);
    var ws = wb.AddWorksheet("Auswertung");

    var headerRow = ws.Row(1);
    headerRow.Cell(1).Value = "Kürzel";
    headerRow.Cell(2).Value = "Vorname";
    headerRow.Cell(3).Value = "Nachname";
    headerRow.Cell(4).Value = "Realstunden";
    headerRow.Cell(5).Value = "Flexminuten/Woche";
    headerRow.Cell(6).Value = "Flexstunden/Woche";
    headerRow.Cell(7).Value = "Flexstunden/Jahr";
    headerRow.Cell(8).Value = "Anzahl Wochen Zeitraum 1";
    headerRow.Cell(9).Value = "Flexstunden pro Woche im Zeitraum 1";
    headerRow.Cell(10).Value = "Anzahl Wochen Zeitraum 2";
    headerRow.Cell(11).Value = "Flexstunden pro Woche im Zeitraum 2";
    headerRow.Cell(12).Value = "Soll-/Istvergleich der Flexstunden pro Jahr";

    var minuteReduction = ws.Row(1).Cell(15);
    minuteReduction.CellLeft().Value = "Minutenreduktion";
    var minuteReductionAddress = minuteReduction.Address.ToStringFixed();
    minuteReduction.Value = 7;

    var hourDuration = ws.Row(2).Cell(15);
    hourDuration.CellLeft().Value = "Stundendauer";
    var hourDurationAddress = hourDuration.Address.ToStringFixed();
    hourDuration.Value = 43;

    var weekCount = ws.Row(3).Cell(15);
    weekCount.CellLeft().Value = "Wochenanzahl";
    var weekCountAddress = weekCount.Address.ToStringFixed();
    weekCount.Value = 43;

    for (int i = 0; i < teacherList.Count; i++)
    {
        var row = ws.Row(i + 2);

        row.Cell(1).Value = teacherList[i].NameCode;
        row.Cell(2).Value = teacherList[i].FirstName;
        row.Cell(3).Value = teacherList[i].LastName;

        var actualHours = row.Cell(4);
        var actualHoursAddress = actualHours.Address.ToStringRelative();
        actualHours.Value = teacherList[i].ActualHours;

        var flexMinutesPerWeek = row.Cell(5);
        var flexMinutesPerWeekAddress = flexMinutesPerWeek.Address.ToStringRelative();
        flexMinutesPerWeek.FormulaA1 = $"={actualHoursAddress}*{minuteReductionAddress}";

        var flexHoursPerWeek = row.Cell(6);
        var flexHoursPerWeekAddress = flexHoursPerWeek.Address.ToStringRelative();
        flexHoursPerWeek.FormulaA1 = $"={flexMinutesPerWeekAddress}/{hourDurationAddress}";

        var flexHoursPerYear = row.Cell(7);
        var flexHoursPerYearAddress = flexHoursPerYear.Address.ToStringRelative();
        flexHoursPerYear.FormulaA1 = $"{flexHoursPerWeekAddress}*{weekCountAddress}";

        var flexHoursPerWeekInTimespan1 = row.Cell(9);
        var flexHoursPerWeekInTimespan1Address = flexHoursPerWeekInTimespan1.Address.ToStringRelative();
        flexHoursPerWeekInTimespan1.FormulaA1 = $"=ROUNDUP({flexHoursPerWeekAddress},0)";

        var flexHoursPerWeekInTimespan2 = row.Cell(11);
        var flexHoursPerWeekInTimespan2Address = flexHoursPerWeekInTimespan2.Address.ToStringRelative();
        flexHoursPerWeekInTimespan2.FormulaA1 = $"=ROUNDDOWN({flexHoursPerWeekAddress},0)";

        var timespan1WeekCount = row.Cell(8);
        var timespan1WeekCountAddress = timespan1WeekCount.Address.ToStringRelative();
        timespan1WeekCount.FormulaA1 = $"=ROUNDDOWN({flexHoursPerYearAddress}-{flexHoursPerWeekInTimespan2Address}*{weekCountAddress},0)";

        var timespan2WeekCount = row.Cell(10);
        var timespan2WeekCountAddress = timespan2WeekCount.Address.ToStringRelative();
        timespan2WeekCount.FormulaA1 = $"={weekCountAddress}-{timespan1WeekCountAddress}";

        var variance = row.Cell(12);
        variance.FormulaA1 = $"={flexHoursPerYearAddress}-({timespan1WeekCountAddress}*{flexHoursPerWeekInTimespan1Address}+{timespan2WeekCountAddress}*{flexHoursPerWeekInTimespan2Address})";
    }

    var table = ws.Range($"A1:L{teacherList.Count + 1}").CreateTable();
    var firstTableRow = table.DataRange.FirstCell().Address.RowNumber;
    var lastTableRow = table.DataRange.LastCell().Address.RowNumber;
            
    ws.Range($"D{firstTableRow}:G{lastTableRow}").Style.NumberFormat.Format = "0.00";
    ws.Range($"H{firstTableRow}:K{lastTableRow}").Style.NumberFormat.Format = "0";
    ws.Range($"L{firstTableRow}:L{lastTableRow}").Style.NumberFormat.Format = "0.00";

    ws.Columns().AdjustToContents();

    wb.Save();
}
