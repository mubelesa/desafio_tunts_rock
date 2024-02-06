using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using static Constants;

class Program
{

    private const string TAB_NAME = "engenharia_de_software";
    private const string RANGE = $"{TAB_NAME}!A4:H27";
    private const string SPREADSHEET_ID = "1TBZV2gld0izZ_mTyckMrIfVKmoHTb7U1OYXRlpGLVW0";

    static void Main()
    {
        string credentialsPath = "C:\\Users\\muril\\source\\repos\\desafio_engenharia_software\\desafio_engenharia_software\\credentials.json";

        var credentials = GetCredentials(credentialsPath);
        var service = BuildSheetsService(credentials);

        ReadAndEditSpreadsheet(service);
    }

    private static SheetsService BuildSheetsService(GoogleCredential credential)
    {
        return new SheetsService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = "Desafio Tunts",
        });
    }

    private static GoogleCredential GetCredentials(string credentialsPath)
    {
        using (var stream = new FileStream(credentialsPath, FileMode.Open, FileAccess.Read))
        {
            return GoogleCredential.FromStream(stream)
                .CreateScoped(new[] { SheetsService.Scope.Spreadsheets });
        }
    }

    private static void ReadAndEditSpreadsheet(SheetsService service)
    {
        var request = service.Spreadsheets.Values.Get(SPREADSHEET_ID, RANGE);
        var response = request.Execute();
        var values = response.Values;
        var line = 4;//skip header

        if (values != null && values.Count > 0)
        {
            foreach (var row in values)
            {
                var situation = Situation.Approved;

                var firstTestGrade = decimal.Parse(row[SheetColumns.TEST1].ToString());
                var secondTestGrade = decimal.Parse(row[SheetColumns.TEST2].ToString());
                var thirdTestGrade = decimal.Parse(row[SheetColumns.TEST3].ToString());
                var faltas = decimal.Parse(row[SheetColumns.ABSCENSES].ToString());

                decimal requiredGradeOnFinalExam = 0;
                decimal media = (firstTestGrade + secondTestGrade + thirdTestGrade) / 3;

                if (faltas > 15)
                    situation = Situation.FailedByAbscence;

                if (media < 50.0M)
                    situation = Situation.FailedByAbscence;

                else if (media >= 50.0M && media < 70.0M)
                {
                    situation = Situation.FinalExam;

                    requiredGradeOnFinalExam = Math.Ceiling(100 - media);
                }

                UpdateStudentSituation(line, situation, requiredGradeOnFinalExam, service);

                line++;
            }
        }
        else
        {
            Console.WriteLine("No data found.");
        }
    }

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

    private static void UpdateStudentSituation(int studentLine, Situation situation, decimal requiredGradeOnFinalExam, SheetsService service)
    {
        var range = $"{TAB_NAME}!G{studentLine}:H{studentLine}";
        var valueRange = new ValueRange();

        var objectList = new List<object>() { TranslateToSheet(situation), requiredGradeOnFinalExam };
        valueRange.Values = new List<IList<object>> { objectList };

        var updateRequest = service.Spreadsheets.Values.Update(valueRange, SPREADSHEET_ID, range);
        updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
        updateRequest.Execute();
    }
}