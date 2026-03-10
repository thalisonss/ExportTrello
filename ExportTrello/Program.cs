using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using ClosedXML.Excel;
using Microsoft.Extensions.Configuration;

class CardResultado
{
    public string Card { get; set; }
    public string ListaAtual { get; set; }
    public DateTime Criado { get; set; }
    public DateTime? EmAndamento { get; set; }
    public DateTime? Concluido { get; set; }
    public string Etiquetas { get; set; }
}

class TrelloConfig
{
    public string Key { get; set; }
    public string Token { get; set; }
    public string BoardId { get; set; }
}

class ExportConfig
{
    public string OutputDirectory { get; set; } = Environment.CurrentDirectory;
    public string FileNamePattern { get; set; } = "trello_export_{date}.xlsx";
}

static class Logger
{
    private static readonly object Sync = new();
    private static readonly string LogFilePath = Path.Combine(Environment.CurrentDirectory, "logs", "export_trello.log");

    public static void Info(string message) => Write("INFO", message);

    public static void Error(string message, Exception? exception = null)
    {
        var fullMessage = exception == null
            ? message
            : $"{message}{Environment.NewLine}{exception}";

        Write("ERROR", fullMessage);
    }

    private static void Write(string level, string message)
    {
        var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] {message}";

        Console.WriteLine(line);

        lock (Sync)
        {
            var logDirectory = Path.GetDirectoryName(LogFilePath);
            if (!string.IsNullOrWhiteSpace(logDirectory))
                Directory.CreateDirectory(logDirectory);

            File.AppendAllText(LogFilePath, line + Environment.NewLine);
        }
    }
}

class Program
{
    static async Task Main()
    {
        try
        {
            var configuration = new ConfigurationBuilder()
                .SetBasePath(AppContext.BaseDirectory)
                .AddJsonFile("appsettings.json", optional: false)
                .Build();

            var trelloConfig = configuration
                .GetSection("Trello")
                .Get<TrelloConfig>();

            var exportConfig = configuration
                .GetSection("Export")
                .Get<ExportConfig>() ?? new ExportConfig();

            string key = trelloConfig.Key;
            string token = trelloConfig.Token;
            string boardId = trelloConfig.BoardId;

            using HttpClient client = new HttpClient();

            Logger.Info("Baixando dados do board...");

            string boardUrl =
                $"https://api.trello.com/1/boards/{boardId}?cards=all&lists=all&key={key}&token={token}";

            string boardJson = await client.GetStringAsync(boardUrl);

            JObject board = JObject.Parse(boardJson);

            var lists = board["lists"]
                .ToDictionary(
                    l => l["id"].ToString(),
                    l => l["name"].ToString()
                );

            var cards = board["cards"]
                .Select(c => new
                {
                    Id = c["id"].ToString(),
                    Nome = c["name"].ToString(),
                    ListaId = c["idList"].ToString(),
                    Labels = string.Join(", ",
                        c["labels"].Select(l => l["name"]?.ToString()).Where(x => !string.IsNullOrWhiteSpace(x)))
                })
                .ToList();

            Logger.Info("Baixando histórico de movimentações...");

            string actionsUrl =
                $"https://api.trello.com/1/boards/{boardId}/actions?filter=updateCard:idList&limit=1000&key={key}&token={token}";

            string actionsJson = await client.GetStringAsync(actionsUrl);

            JArray actions = JArray.Parse(actionsJson);

            var andamentoPorCard = new Dictionary<string, DateTime>();
            var concluidoPorCard = new Dictionary<string, DateTime>();

            foreach (var action in actions)
            {
                var data = action["data"];

                if (data["listAfter"] == null)
                    continue;

                string listName = data["listAfter"]["name"].ToString().ToLower();
                string cardId = data["card"]["id"].ToString();
                DateTime date = DateTime.Parse(action["date"].ToString());

                if (listName.Contains("andamento"))
                {
                    if (!andamentoPorCard.ContainsKey(cardId))
                        andamentoPorCard.Add(cardId, date);
                }

                if (listName.Contains("concluido"))
                {
                    if (!concluidoPorCard.ContainsKey(cardId))
                        concluidoPorCard.Add(cardId, date);
                }
            }

            Logger.Info("Processando cards...");

            var resultado = new List<CardResultado>();

            foreach (var card in cards)
            {
                DateTime criado = GetCreationDate(card.Id);

                DateTime? andamento = null;
                DateTime? concluido = null;

                if (andamentoPorCard.ContainsKey(card.Id))
                    andamento = andamentoPorCard[card.Id];

                if (concluidoPorCard.ContainsKey(card.Id))
                    concluido = concluidoPorCard[card.Id];

                resultado.Add(new CardResultado
                {
                    Card = card.Nome,
                    ListaAtual = lists[card.ListaId],
                    Criado = criado,
                    EmAndamento = andamento,
                    Concluido = concluido,
                    Etiquetas = card.Labels
                });
            }

            Logger.Info("Gerando Excel...");

            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Cards");

            ws.Cell(1, 1).Value = "Card";
            ws.Cell(1, 2).Value = "Lista Atual";
            ws.Cell(1, 3).Value = "Criado";
            ws.Cell(1, 4).Value = "Entrou Em Andamento";
            ws.Cell(1, 5).Value = "Entrou Concluido";
            ws.Cell(1, 6).Value = "Etiquetas";

            int row = 2;

            foreach (var r in resultado)
            {
                ws.Cell(row, 1).Value = r.Card;
                ws.Cell(row, 2).Value = r.ListaAtual;
                ws.Cell(row, 3).Value = r.Criado;

                if (r.EmAndamento.HasValue)
                    ws.Cell(row, 4).Value = r.EmAndamento.Value;

                if (r.Concluido.HasValue)
                    ws.Cell(row, 5).Value = r.Concluido.Value;

                ws.Cell(row, 6).Value = r.Etiquetas;

                row++;
            }

            ws.Column(3).Style.DateFormat.Format = "dd/MM/yyyy HH:mm";
            ws.Column(4).Style.DateFormat.Format = "dd/MM/yyyy HH:mm";
            ws.Column(5).Style.DateFormat.Format = "dd/MM/yyyy HH:mm";

            ws.Columns().AdjustToContents();

            string fileName = exportConfig.FileNamePattern
                .Replace("{date}", DateTime.Now.ToString("yyyy-MM-dd"))
                .Replace("{datetime}", DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss"));

            string outputDirectory = string.IsNullOrWhiteSpace(exportConfig.OutputDirectory)
                ? Environment.CurrentDirectory
                : exportConfig.OutputDirectory;

            Directory.CreateDirectory(outputDirectory);

            string file = Path.Combine(outputDirectory, fileName);

            workbook.SaveAs(file);

            Logger.Info($"Planilha criada: {file}");
            Logger.Info($"Logs gravados em: {Path.Combine(Environment.CurrentDirectory, "logs", "export_trello.log")}");
        }
        catch (Exception ex)
        {
            Logger.Error("Erro durante a exportação do Trello.", ex);
            throw;
        }
    }

    static DateTime GetCreationDate(string cardId)
    {
        string hex = cardId.Substring(0, 8);
        long seconds = Convert.ToInt64(hex, 16);

        return DateTimeOffset
            .FromUnixTimeSeconds(seconds)
            .LocalDateTime;
    }
}
