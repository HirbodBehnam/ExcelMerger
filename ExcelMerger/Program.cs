using OfficeOpenXml;
using System.Text;

namespace ExcelMerger;

internal class Program
{
	static void Main(string[] args)
	{
		ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
		var names = GetStudentMap("std.csv");
		// Check std merge mode
		if (args.Length > 0 && args[0] == "-m")
		{
			MergeStudentMap(names);
			return;
		}

		File.Delete("destination.xlsx");
		using var source = new ExcelPackage(new FileInfo("raw.xlsx"));
		using var destination = new ExcelPackage(new FileInfo("destination.xlsx"));
		foreach (ExcelWorksheet? sheet in source.Workbook.Worksheets)
		{
			// Add the sheet
			ExcelWorksheet? destinationWorksheet = destination.Workbook.Worksheets.Add(sheet.Name);
			// Reverse columns if needed
			if (args.Length > 0 && args[0] == "-r")
				ReverseColumns(sheet);
			// Copy first row
			destinationWorksheet.Cells[1, 1].Value = "Name";
			for (var column = 1; column <= sheet.Dimension.End.Column; column++)
				destinationWorksheet.Cells[1, column + 1].Value = Reverse(sheet.Cells[1, column].Text);
			// Add names
			for (var row = 2; row <= sheet.Dimension.End.Row; row++)
			{
				string stdNumber = ToEnglishNumber(sheet.Cells[row, 1].Text);
				destinationWorksheet.Cells[row, 1].Value = names.ContainsKey(stdNumber) ? names[stdNumber] : "ناشناس";
				for (var column = 1; column <= sheet.Dimension.End.Column; column++)
				{
					string data = ToEnglishNumber(sheet.Cells[row, column].Text);
					if (double.TryParse(data, out double dataInt))
						destinationWorksheet.Cells[row, column + 1].Value = dataInt;
					else
						destinationWorksheet.Cells[row, column + 1].Value = Reverse(data);
				}
			}
		}

		destination.Save();
	}

	private static Dictionary<string, string> GetStudentMap(string path)
	{
		var names = new Dictionary<string, string>();
		using var reader = new StreamReader(path);
		reader.ReadLine(); // skip headers
		while (reader.ReadLine() is { } line)
		{
			string[] splitted = line.Split(',');
			names.Add(splitted[0].Trim(), splitted[1].Trim());
		}

		return names;
	}

	private static void MergeStudentMap(IDictionary<string, string> original)
	{
		var newStudents = GetStudentMap("new-std.csv");
		foreach ((string stdId, string stdName) in newStudents)
			if (!original.ContainsKey(stdId))
			{
				original[stdId] = stdName;
				Console.WriteLine($"Added {stdId}");
			}

		// Write back
		using var writer = new StreamWriter("std.csv");
		writer.WriteLine("stdId,stdName");
		foreach ((string stdId, string stdName) in original)
			writer.WriteLine($"{stdId},{stdName}");
	}

	private static string ToEnglishNumber(string input)
	{
		StringBuilder englishNumbers = new(input.Length);
		for (var i = 0; i < input.Length; i++)
		{
			if (char.IsDigit(input[i]))
				englishNumbers.Append(char.GetNumericValue(input, i));
			else
				englishNumbers.Append(input[i]);
		}

		return englishNumbers.ToString();
	}

	private static void ReverseColumns(ExcelWorksheet sheet)
	{
		// Note that this function will throw an exception if there are merged cells
		int start = 1, end = sheet.Dimension.End.Column;
		int tempColumn = end + 1; // we move the temp data here
		int rows = sheet.Dimension.End.Row; // total number of rows to copy
		while (start < end)
		{
			sheet.Cells[1, start, rows, start].Copy(sheet.Cells[1, tempColumn, rows, tempColumn]);
			sheet.Cells[1, end, rows, end].Copy(sheet.Cells[1, start, rows, start]);
			sheet.Cells[1, tempColumn, rows, tempColumn].Copy(sheet.Cells[1, end, rows, end]);
			sheet.DeleteColumn(tempColumn);
			start++;
			end--;
		}
	}

	private static string Reverse(string s)
	{
		//return s;
		char[] charArray = s.ToCharArray();
		Array.Reverse(charArray);
		return new string(charArray);
	}
}