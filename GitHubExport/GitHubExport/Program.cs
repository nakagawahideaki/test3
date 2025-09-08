using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Transactions;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

using ClosedXML.Excel;
using GraphQL;
using GraphQL.Client.Http;
using GraphQL.Client.Serializer.Newtonsoft;
using Newtonsoft.Json.Linq;
using System.Net.Http.Headers;
using System.Threading.Tasks;

public class GitHubProjectUpdater
{
    // GraphQLクライアント
    private readonly GraphQLHttpClient _client;
    // リポジトリの所有者名
    private readonly string _owner;
    // リポジトリ名
    private readonly string _repo;

    // コンストラクタ：GitHubトークン、所有者、リポジトリ名を受け取る
    public GitHubProjectUpdater(string githubToken, string owner, string repo)
    {
        _owner = owner;
        _repo = repo;

        // GraphQLクライアントの初期化
        _client = new GraphQLHttpClient(new GraphQLHttpClientOptions
        {
            EndPoint = new Uri("https://api.github.com/graphql")
        }, new NewtonsoftJsonSerializer());

        // 認証ヘッダーの設定
        _client.HttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", githubToken);
    }

    // リポジトリIDを取得する
    private async Task<string?> GetRepositoryId()
    {
        // GraphQLリクエストの作成
        var request = new GraphQLRequest
        {
            // リポジトリIDを取得するためのクエリ
            Query = @"
query($owner: String!, $repo: String!) {
  repository(owner: $owner, name: $repo) {
    id
  }
}",
            // クエリ変数
            Variables = new { owner = _owner, repo = _repo }
        };

        // GraphQLクエリの実行
        var response = await _client.SendQueryAsync<dynamic>(request);

        // エラー処理
        if (response.Errors != null && response.Errors.Length > 0)
        {
            foreach (var error in response.Errors)
            {
                Console.WriteLine($"GraphQL Error: {error.Message}");
            }

            throw new Exception("GraphQL request failed.");
        }

        // リポジトリIDを返す
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

    // プロジェクトIDを取得する
    public async Task<string?> GetProjectId(string projectName)
    {
        // GraphQLリクエストの作成
        var request = new GraphQLRequest
        {
            // プロジェクトIDを取得するためのクエリ
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
            // クエリ変数
            Variables = new
            {
                owner = _owner,
                repo = _repo,
                projectName = projectName
            }
        };

        // GraphQLクエリの実行
        var response = await _client.SendQueryAsync<dynamic>(request);

        // エラー処理
        if (response.Errors != null && response.Errors.Length > 0)
        {
            foreach (var error in response.Errors)
            {
                Console.WriteLine($"GraphQL Error: {error.Message}");
            }
            throw new Exception("GraphQL request failed.");
        }

        // プロジェクトIDを返す
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

    // Issueを作成する
    private async Task<string> CreateIssue(string title, string body)
    {
        // リポジトリIDの取得
        string? repositoryId = await GetRepositoryId();

        // リポジトリID取得失敗時のエラー処理
        if (repositoryId == null)
        {
            throw new Exception("Failed to get repository ID.");
        }

        // GraphQLリクエストの作成
        var createIssueRequest = new GraphQLRequest
        {
            // Issueを作成するためのクエリ
            Query = @"
mutation ($repositoryId: ID!, $title: String!, $body: String!) {
  createIssue(input: {repositoryId: $repositoryId, title: $title, body: $body}) {
    issue {
      id
    }
  }
}",
            // クエリ変数
            Variables = new
            {
                repositoryId = repositoryId,
                title = title,
                body = body
            }
        };

        // GraphQLクエリの実行
        var createIssueResponse = await _client.SendMutationAsync<dynamic>(createIssueRequest);

        // エラー処理
        if (createIssueResponse.Errors != null && createIssueResponse.Errors.Length > 0)
        {
            foreach (var error in createIssueResponse.Errors)
            {
                Console.WriteLine($"GraphQL Error: {error.Message}");
            }

            throw new Exception("GraphQL request failed.");
        }

        // IssueのIDを返す
        return createIssueResponse.Data.createIssue.issue.id.ToString();
    }

    // プロジェクトにアイテムを追加する
    public async Task AddProjectV2ItemById(string projectId, string contentId)
    {
        // GraphQLリクエストの作成
        var request = new GraphQLRequest
        {
            // プロジェクトにアイテムを追加するためのクエリ
            Query = @"
mutation ($projectId: ID!, $contentId: ID!) {
  addProjectV2ItemById(input: {projectId: $projectId, contentId: $contentId}) {
    item {
      id
    }
  }
}",
            // クエリ変数
            Variables = new
            {
                projectId = projectId,
                contentId = contentId
            }
        };

        // GraphQLクエリの実行
        var response = await _client.SendMutationAsync<dynamic>(request);

        // エラー処理
        if (response.Errors != null && response.Errors.Length > 0)
        {
            foreach (var error in response.Errors)
            {
                Console.WriteLine($"GraphQL Error: {error.Message}");
            }

            throw new Exception("GraphQL request failed.");
        }

        // 成功メッセージを出力
        Console.WriteLine($"Item added to project. Item ID: {contentId}");
    }

    // Excelデータに基づいてプロジェクトを更新する
    public async Task UpdateExcelDataToProject(string excelFilePath, string projectId)
    {
        // Excelファイルを開く
        using (var workbook = new XLWorkbook(excelFilePath))
        {
            // 1番目のワークシートを取得
            var worksheet = workbook.Worksheet(1);
            // 使用されている最後の行を取得
            var lastRowUsed = worksheet.LastRowUsed();

            // 各行を処理
            for (int row = 2; row <= lastRowUsed?.RowNumber(); row++)
            {
                // タイトルと本文を取得
                string title = worksheet.Cell(row, 1).Value.ToString();
                string body = worksheet.Cell(row, 2).Value.ToString();

                try
                {
                    // Issueを作成
                    string issueId = await CreateIssue(title, body);
                    // プロジェクトにIssueを追加
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
    // メイン関数
    public static async Task Main(string[] args)
    {
        // GitHubトークン、Excelファイルパス、プロジェクト名、所有者、リポジトリを設定
        //var githubToken = "***";
        //var excelFilePath = @"C:\Users\nakagawa\Desktop\GitHubExport2.xlsx";
        //var projectName = "KanbanTest";
        //var owner = "nakagawahideaki";
        //var repo = "test3";
        var githubToken = args.FirstOrDefault(a => a.StartsWith("--github-token"))?.Split('=')[1];
        var excelFilePath = args.FirstOrDefault(a => a.StartsWith("--excel-file-path"))?.Split('=')[1];
        var projectName = args.FirstOrDefault(a => a.StartsWith("--project-name"))?.Split('=')[1];
        var owner = args.FirstOrDefault(a => a.StartsWith("--owner"))?.Split('=')[1];
        var repo = args.FirstOrDefault(a => a.StartsWith("--repo"))?.Split('=')[1];

        // 引数のチェック
        if (string.IsNullOrEmpty(githubToken) || string.IsNullOrEmpty(excelFilePath) ||
            string.IsNullOrEmpty(projectName) || string.IsNullOrEmpty(owner) || string.IsNullOrEmpty(repo))
        {
            Console.WriteLine("Missing required arguments");
            return;
        }

        // GitHubProjectUpdaterインスタンスを作成
        var updater = new GitHubProjectUpdater(githubToken, owner, repo);
        // プロジェクトIDを取得
        var projectId = await updater.GetProjectId(projectName);

        // プロジェクトIDが取得できた場合
        if (projectId != null)
        {
            // Excelデータに基づいてプロジェクトを更新
            await updater.UpdateExcelDataToProject(excelFilePath, projectId);
            Console.WriteLine("Finished.");
        }
    }
}