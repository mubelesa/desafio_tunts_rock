using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using static Constants;

class Program
{
    // Define constants for the spreadsheet tab name, range, and ID
    private const string TAB_NAME = "engenharia_de_software";
    private const string RANGE = $"{TAB_NAME}!A4:H27";
    private const string SPREADSHEET_ID = "1TBZV2gld0izZ_mTyckMrIfVKmoHTb7U1OYXRlpGLVW0";

    static void Main()
    {
         // Path to the credentials file
        string credentialsPath = "C:\\Users\\muril\\source\\repos\\desafio_engenharia_software\\desafio_engenharia_software\\credentials.json";
        
        // Get credentials and build the Google Sheets service
        var credentials = GetCredentials(credentialsPath);
        var service = BuildSheetsService(credentials);

         // Read and edit the spreadsheet
        ReadAndEditSpreadsheet(service);
    }

    // Build the Google Sheets service with the provided credentials
    private static SheetsService BuildSheetsService(GoogleCredential credential)
    {
        return new SheetsService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = "Desafio Tunts",
        });
    }

    // Create scoped credentials for the Google Sheets API
    private static GoogleCredential GetCredentials(string credentialsPath)
    {
        using (var stream = new FileStream(credentialsPath, FileMode.Open, FileAccess.Read))
        {
            // Create scoped credentials for the Google Sheets API
            return GoogleCredential.FromStream(stream)
                .CreateScoped(new[] { SheetsService.Scope.Spreadsheets });
        }
    }

    // Read data from and edit the specified spreadsheet
    private static void ReadAndEditSpreadsheet(SheetsService service)
    {
        // Get values from the spreadsheet
        var request = service.Spreadsheets.Values.Get(SPREADSHEET_ID, RANGE);
        var response = request.Execute();
        var values = response.Values;
        var line = 4;//skip header

         // Check if data was found and process each row
        if (values != null && values.Count > 0)
        {
            foreach (var row in values)
            {
                // Default situation
                var situation = Situation.Approved;

                // Parse grades and absences
                var firstTestGrade = decimal.Parse(row[SheetColumns.TEST1].ToString());
                var secondTestGrade = decimal.Parse(row[SheetColumns.TEST2].ToString());
                var thirdTestGrade = decimal.Parse(row[SheetColumns.TEST3].ToString());
                var faltas = decimal.Parse(row[SheetColumns.ABSCENSES].ToString());

                 // Calculate required grade for final exam
                decimal requiredGradeOnFinalExam = 0;
                decimal media = (firstTestGrade + secondTestGrade + thirdTestGrade) / 3;

                // Determine student situation based on grades and absences
                if (absences > 15)
                    situation = Situation.FailedByAbscence;
                if (faltas > 15)
                    situation = Situation.FailedByAbscence;

                if (media < 50.0M)
                    situation = Situation.FailedByAbscence;

                else if (media >= 50.0M && media < 70.0M)
                {
                    situation = Situation.FinalExam;

                    requiredGradeOnFinalExam = Math.Ceiling(100 - media);
                }
                // Update the student's situation in the spreadsheet
                UpdateStudentSituation(line, situation, requiredGradeOnFinalExam, service);

                line++;
            }
        }
        else
        {
            Console.WriteLine("No data found.");
        }
    }

    // Translate the situation enum to a string for the spreadsheet
    public static string TranslateToSheet(Situation situation)
    {
        return situation switch
        {
            Situation.Approved => "Aprovado",
            Situation.Failed => "Reprovado por Nota",
            Situation.FailedByAbscence => "Reprovado por Falta",
            Situation.FinalExam => "Exame Final",
            _ => "Situação não encontrada",
        };
    }

    // Update the student's situation and required grade for the final exam in the spreadsheet
    private static void UpdateStudentSituation(int studentLine, Situation situation, decimal requiredGradeOnFinalExam, SheetsService service)
    {
        // Define the range to update
        var range = $"{TAB_NAME}!G{studentLine}:H{studentLine}";
        var valueRange = new ValueRange();

        // Prepare the values to update
        var objectList = new List<object>() { TranslateToSheet(situation), requiredGradeOnFinalExam };
        valueRange.Values = new List<IList<object>> { objectList };

        var updateRequest = service.Spreadsheets.Values.Update(valueRange, SPREADSHEET_ID, range);
        updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
        updateRequest.Execute();
    }
}
