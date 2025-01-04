using Microsoft.Graph;
using Microsoft.Graph.Models;

public static class OneDriveUtil
{
	/// <summary>
	/// フォルダを作成します。
	/// すでにフォルダがある場合は何もしません。
	/// </summary>
	/// <param name="graphClient">GraphServiceClient。</param>
	/// <param name="oneDriveId">OneDrive の ID。フォルダの ID ではありません。</param>
	/// <param name="folderId">フォルダを作成する親フォルダの ID。</param>
	/// <param name="folderName">作成するフォルダ名。</param>
	/// <returns>作成したフォルダの情報。</returns>
	public async static Task<DriveItem?> CreateFolder(GraphServiceClient graphClient, string oneDriveId, string folderId, string folderName)
	{
		var requestBody = new DriveItem
		{
			Name = folderName,
			Folder = new(),
			AdditionalData = new Dictionary<string, object>
			{
				{"@microsoft.graph.conflictBehavior" , "replace" },
      },
		};
		return await graphClient.Drives[oneDriveId].Items[folderId].Children.PostAsync(requestBody);
	}

	/// <summary>
	/// ファイルをアップロードします。上限はおおよそ 50MB 付近です。
	/// </summary>
	/// <param name="graphClient">GraphServiceClient。</param>
	/// <param name="oneDriveId">OneDrive の ID。フォルダの ID ではありません。</param>
	/// <param name="uploadFolderId">アップロード先フォルダ ID。</param>
	/// <param name="fileName">アップロード先のファイル名。</param>
	/// <param name="uploadData">アップロードするデータ。</param>
	/// <returns></returns>
	public async static Task<DriveItem?> UploadFile(GraphServiceClient graphClient, string oneDriveId, string uploadFolderId, string fileName, Stream uploadData)
		=> await graphClient.Drives[oneDriveId].Items[uploadFolderId].ItemWithPath(fileName).Content.PutAsync(uploadData);
}