using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;

// 各種 ID などの定義
var clientId = "XXXXXXXX";      // クライアント ID
var tenantId = "XXXXXXXX";      // テナント ID
var clientSecret = "XXXXXXXX";  // クライアント シークレット
var userId = "XXXXXXXX";        // ユーザー ID

var testFolderName = "TestFolder";
var isCreateTestFile = true;     // テスト用に200件以上のファイルを作成するかどうか。そこそこ時間が掛かるので注意


// 使いまわすので最初に定義しておく
HttpClient httpClient = new();

// クライアント シークレットによる認証情報を定義
ClientSecretCredential clientSecretCredential = new(tenantId, clientId, clientSecret);

// HttpClient と認証情報で Microsoft Graph サービスのクライアントを生成
using GraphServiceClient graphClient = new(httpClient, clientSecretCredential);

// 対象ユーザーに紐づく OneDrive を取得 (紐づいているドライブが OneDrive １つという前提)
var drive = await graphClient.Users[userId].Drive.GetAsync();
if (drive == null || drive.Id == null)
{
	Console.WriteLine("ドライブを取得できませんでした。");
	return;
}

// OneDrive のルートを取得
var root = await graphClient.Drives[drive.Id].Root.GetAsync();
if (root == null || root.Id == null)
{
	Console.WriteLine("OneDrive のルートを取得できませんでした。");
	return;
}

// ルートにテスト用のフォルダ作成
var testFolderItem = await OneDriveUtil.CreateFolder(graphClient, drive.Id, root.Id, testFolderName);
if (testFolderItem == null || testFolderItem.Id == null)
{
	Console.WriteLine("テストフォルダを作成できませんでした。");
	return;
}

// テスト用ファイル作成
if (isCreateTestFile)
{
	for (var i = 0;	i < 250; i++)
	{
		using MemoryStream ms = new MemoryStream();
		ms.Write("テストです"u8);
		ms.Seek(0, SeekOrigin.Begin);
		var fileName = $"{DateTime.Now:yyyyMMddHHmmss}_{i + 1}.txt";
		Console.WriteLine($"{fileName} をアップロードしています。");
		await OneDriveUtil.UploadFile(graphClient, drive.Id, testFolderItem.Id, fileName, ms);
	}
}

// テストフォルダの中身一覧取得
var folderChildren = await graphClient.Drives[drive.Id].Items[testFolderItem.Id].Children.GetAsync();
if (folderChildren == null || folderChildren.Value == null)
{
	Console.WriteLine("フォルダの一覧を取得できませんでした。");
	return;
}

Console.WriteLine("Value プロパティを使用した一覧取得です。");
var count = 0;
foreach (var item in folderChildren.Value)
{
	count++;
	Console.WriteLine($"{count:d4}: Type={(item.File != null ? "File" : "Folder")}, Id={item.Id}, Name={item.Name}, Size={item.Size}");
}

Console.WriteLine("");
Console.WriteLine("PageIterator を使用した一覧取得です。");
count = 0;
var pageIterator = PageIterator<DriveItem, DriveItemCollectionResponse>
		.CreatePageIterator(
				graphClient,
				folderChildren,
				(item) =>
				{
					count++;
					Console.WriteLine($"{count:d4}: Type={(item.File != null ? "File" : "Folder")}, Id={item.Id}, Name={item.Name}, Size={item.Size}");
					return true;  // false を返すまで次のアイテムを列挙します
				});
await pageIterator.IterateAsync();

