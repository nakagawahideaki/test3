using ClosedXML.Excel;
using GraphQL;
using GraphQL.Client.Http;
using GraphQL.Client.Serializer.Newtonsoft;
using System.Net.Http.Headers;
using System.IO;

public class GitHubProjectUpdater
{
    private readonly GraphQLHttpClient _client;
    private readonly string _owner;
    private readonly string _repo;

    public GitHubProjectUpdater(string githubToken, string owner, string repo)
    {
        _owner = owner;
        _repo = repo;

        _client = new GraphQLHttpClient(new GraphQLHttpClientOptions
        {
            EndPoint = new Uri("https://api.github.com/graphql")
        }, new NewtonsoftJsonSerializer());

        _client.HttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", githubToken);
    }

    private async Task<string?> GetRepositoryId()
    {
        var request = new GraphQLRequest
        {
            Query = @"
query($owner: String!, $repo: String!) {
    repository(owner: $owner, name: $repo) {
    id
    }
}",
            Variables = new { owner = _owner, repo = _repo }
        };

        var response = await _client.SendQueryAsync<dynamic>(request);

        if (response.Errors != null && response.Errors.Length > 0)
        {
            foreach (var error in response.Errors)
            {
                Console.WriteLine($"GraphQL Error: {error.Message}");
            }
            throw new Exception("GraphQL request failed.");
        }

        try
        {
            return response.Data.repository.id.ToString();
        }
        catch
        {
            Console.WriteLine($"Repository '{_owner}/{_repo}' not found.");
            return null;
        }
    }

    public async Task<string?> GetProjectId(string projectName)
    {
        var request = new GraphQLRequest
        {
            Query = @"
query ($owner: String!, $repo: String!, $projectName: String!) {
    repository(owner: $owner, name: $repo) {
    projectsV2(query: $projectName, first: 1) {
        nodes {
        id
        }
    }
    }
}",
            Variables = new
            {
                owner = _owner,
                repo = _repo,
                projectName = projectName
            }
        };

        var response = await _client.SendQueryAsync<dynamic>(request);

        Console.WriteLine(Newtonsoft.Json.JsonConvert.SerializeObject(response, Newtonsoft.Json.Formatting.Indented));

        if (response.Errors != null && response.Errors.Length > 0)
        {
            foreach (var error in response.Errors)
            {
                Console.WriteLine($"GraphQL Error: {error.Message}");
            }
            return null;
        }

        try
        {
            return response.Data.repository.projectsV2.nodes[0].id.ToString();
        }
        catch
        {
            Console.WriteLine($"Project '{projectName}' not found.");
            return null;
        }
    }

    private async Task<string> CreateIssue(string title, string body)
    {
        string? repositoryId = await GetRepositoryId();

        if (repositoryId == null)
        {
            throw new Exception("Failed to get repository ID.");
        }

        var createIssueRequest = new GraphQLRequest
        {
            Query = @"
mutation ($repositoryId: ID!, $title: String!, $body: String!) {
    createIssue(input: {repositoryId: $repositoryId, title: $title, body: $body}) {
    issue {
        id
    }
    }
}",
            Variables = new
            {
                repositoryId = repositoryId,
                title = title,
                body = body
            }
        };

        var createIssueResponse = await _client.SendMutationAsync<dynamic>(createIssueRequest);

        if (createIssueResponse.Errors != null && createIssueResponse.Errors.Length > 0)
        {
            foreach (var error in createIssueResponse.Errors)
            {
                Console.WriteLine($"GraphQL Error: {error.Message}");
            }
            throw new Exception("GraphQL request failed.");
        }

        return createIssueResponse.Data.createIssue.issue.id.ToString();
    }

    public async Task AddProjectV2ItemById(string projectId, string contentId)
    {
        var request = new GraphQLRequest
        {
            Query = @"
mutation ($projectId: ID!, $contentId: ID!) {
    addProjectV2ItemById(input: {projectId: $projectId, contentId: $contentId}) {
    item {
        id
    }
    }
}",
            Variables = new
            {
                projectId = projectId,
                contentId = contentId
            }
        };

        var response = await _client.SendMutationAsync<dynamic>(request);

        if (response.Errors != null && response.Errors.Length > 0)
        {
            foreach (var error in response.Errors)
            {
                Console.WriteLine($"GraphQL Error: {error.Message}");
            }
            throw new Exception("GraphQL request failed.");
        }

        Console.WriteLine($"Item added to project. Item ID: {contentId}");
    }

    public async Task UpdateExcelDataToProject(string excelFilePath, string projectId)
    {
        using (var workbook = new XLWorkbook(excelFilePath))
        {
            var worksheet = workbook.Worksheet(1);
            var lastRowUsed = worksheet.LastRowUsed();

            for (int row = 2; row <= lastRowUsed?.RowNumber(); row++)
            {
                string title = worksheet.Cell(row, 1).Value.ToString();
                string body = worksheet.Cell(row, 2).Value.ToString();

                try
                {
                    string issueId = await CreateIssue(title, body);
                    await AddProjectV2ItemById(projectId, issueId);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing row {row}: {ex.Message}");
                }
            }
        }
    }
}

// 実行例
public class Example
{
    public static async Task Main(string[] args)
    {
        if (args.Length < 5)
        {
            Console.WriteLine("Error: Insufficient arguments.");
            return;
        }

        // GitHubトークン、Excelファイルパス、プロジェクト名、所有者、リポジトリを設定
        //var githubToken = "aaa";
        //var excelFilePath = @"C:\Users\nakagawa\Desktop\GitHubExport.xlsx";
        //var projectName = "KanbanTest";
        //var owner = "nakagawahideaki";
        //var repo = "test3";
        var githubToken = args[0];
        var excelFilePath = args[1];
        var projectName = args[2];
        var owner = args[3];
        var repo = args[4];

        // 一時ファイルパスを仮想的に生成
        var tempFilePath = Path.GetTempFileName() + ".xlsx"; // 一時ファイルとして.xlsxを付け加える
        Console.WriteLine(excelFilePath);
        Console.WriteLine(tempFilePath);

        // Excelファイルを仮パスにコピー
        File.Copy(excelFilePath, tempFilePath, true);
        Console.WriteLine("Copied original Excel file to temporary location.");

        // GitHubProjectUpdaterインスタンスを作成
        var updater = new GitHubProjectUpdater(githubToken, owner, repo);
        var projectId = await updater.GetProjectId(projectName);

        if (projectId != null)
        {
            // Excelデータに基づいてプロジェクトを更新
            await updater.UpdateExcelDataToProject(tempFilePath, projectId.ToString());
            Console.WriteLine("Finished.");
        }

        // プログラム終了後に一時ファイルを削除
        File.Delete(tempFilePath);
    }
}