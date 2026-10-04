using Microsoft.SharePoint.Client;
using PnP.Core.Model.SharePoint;
using PnP.Framework.Utilities;
using PnP.PowerShell.Commands.Model.Graph;
using System;
using System.Collections;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace PnP.PowerShell.Commands.Files
{
    [Cmdlet(VerbsCommon.Get, "PnPFile", DefaultParameterSetName = URLASFILEOBJECT)]
    public class GetFile : PnPWebCmdlet
    {
        private const string URLTOPATH = "Save to local path";
        private const string URLASSTRING = "Return as string";
        private const string URLASLISTITEM = "Return as list item";
        private const string URLASFILEOBJECT = "Return as file object";
        private const string URLASMEMORYSTREAM = "Return as memorystream";

        [Parameter(Mandatory = true, ParameterSetName = URLASFILEOBJECT, Position = 0, ValueFromPipeline = true)]
        [Parameter(Mandatory = true, ParameterSetName = URLASLISTITEM, Position = 0, ValueFromPipeline = true)]
        [Parameter(Mandatory = true, ParameterSetName = URLTOPATH, Position = 0, ValueFromPipeline = true)]
        [Parameter(Mandatory = true, ParameterSetName = URLASSTRING, Position = 0, ValueFromPipeline = true)]
        [Parameter(Mandatory = true, ParameterSetName = URLASMEMORYSTREAM, Position = 0, ValueFromPipeline = true)]
        [Alias("ServerRelativeUrl", "SiteRelativeUrl")]
        public string Url;

        [Parameter(Mandatory = true, ParameterSetName = URLTOPATH)]
        public string Path = string.Empty;

        [Parameter(Mandatory = true, ParameterSetName = URLTOPATH)]
        public string Filename = string.Empty;

        [Parameter(Mandatory = true, ParameterSetName = URLTOPATH)]
        public SwitchParameter AsFile;

        [Parameter(Mandatory = true, ParameterSetName = URLASLISTITEM)]
        public SwitchParameter AsListItem;

        [Parameter(Mandatory = false, ParameterSetName = URLASLISTITEM)]
        public SwitchParameter ThrowExceptionIfFileNotFound;

        [Parameter(Mandatory = false, ParameterSetName = URLASSTRING)]
        public SwitchParameter AsString;

        [Parameter(Mandatory = false, ParameterSetName = URLTOPATH)]
        public SwitchParameter Force;

        [Parameter(Mandatory = false, ParameterSetName = URLASFILEOBJECT)]
        public SwitchParameter AsFileObject;

        [Parameter(Mandatory = false, ParameterSetName = URLASMEMORYSTREAM)]
        public SwitchParameter AsMemoryStream;

        [Parameter(Mandatory = false, ParameterSetName = URLASFILEOBJECT)]
        [Parameter(Mandatory = false, ParameterSetName = URLASLISTITEM)]
        [Parameter(Mandatory = false, ParameterSetName = URLTOPATH)]
        [Parameter(Mandatory = false, ParameterSetName = URLASSTRING)]
        [Parameter(Mandatory = false, ParameterSetName = URLASMEMORYSTREAM)]
        public SwitchParameter UseGraph;

        protected override void ExecuteCmdlet()
        {
            var serverRelativeUrl = string.Empty;
            if (string.IsNullOrEmpty(Path))
            {
                Path = SessionState.Path.CurrentFileSystemLocation.Path;
            }
            else
            {
                if (!System.IO.Path.IsPathRooted(Path))
                {
                    Path = System.IO.Path.Combine(SessionState.Path.CurrentFileSystemLocation.Path, Path);
                }
            }

            if (UseGraph)
            {
                // Must not touch the site, web or list: an app holding only a Lists, ListItems or Files selected
                // permission cannot read any of them, and lower selected scopes never grant access upwards.
                ExecuteThroughGraph();
                return;
            }

            if (Uri.IsWellFormedUriString(Url, UriKind.Absolute))
            {
                // We can't deal with absolute URLs
                Url = UrlUtility.MakeRelativeUrl(Url);
            }

            var webUrl = CurrentWeb.EnsureProperty(w => w.ServerRelativeUrl);

            // Every branch below uses the Url as provided when a file exists there and only falls back to its decoded
            // form when it does not. Resolving inside the branch keeps that to a single lookup per parameter set, and
            // the PnP Core branches have to address the file by its unique id rather than by its URL.
            switch (ParameterSetName)
            {
                case URLTOPATH:

                    // Get a reference to the file to download
                    IFile fileToDownload = Utilities.FileUrlResolver.ResolveFile(Url, webUrl, ClientContext, CurrentWeb, Connection.PnPContext);
                    string fileToDownloadName = !string.IsNullOrEmpty(Filename) ? Filename : fileToDownload.Name;
                    string fileOut = System.IO.Path.Combine(Path, fileToDownloadName);

                    if (System.IO.File.Exists(fileOut) && !Force)
                    {
                        LogWarning($"File '{fileToDownloadName}' exists already. Use the -Force parameter to overwrite the file.");
                    }
                    else
                    {
                        SaveFileToLocal(fileToDownload, fileOut).GetAwaiter().GetResult();
                    }

                    break;
                case URLASFILEOBJECT:
                    serverRelativeUrl = Utilities.FileUrlResolver.Resolve(Url, webUrl, ClientContext, CurrentWeb);
                    var fileObject = CurrentWeb.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeUrl));
                    try
                    {
                        ClientContext.Load(fileObject, f => f.Author, f => f.Length, f => f.ModifiedBy, f => f.Name, f => f.TimeCreated, f => f.TimeLastModified, f => f.Title);
                        ClientContext.ExecuteQueryRetry();
                    }
                    catch (ServerException)
                    {
                        // Assume the cause of the exception is that a principal cannot be found and try again without:
                        // Fallback in case the creator or person having last modified the file no longer exists in the environment such that the file can still be downloaded
                        ClientContext.Load(fileObject, f => f.Length, f => f.Name, f => f.TimeCreated, f => f.TimeLastModified, f => f.Title);
                        ClientContext.ExecuteQueryRetry();
                    }
                    WriteObject(fileObject);
                    break;
                case URLASLISTITEM:
                    serverRelativeUrl = Utilities.FileUrlResolver.Resolve(Url, webUrl, ClientContext, CurrentWeb);
                    var fileListItem = CurrentWeb.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeUrl));

                    ClientContext.Load(fileListItem, f => f.Exists, f => f.ListItemAllFields);

                    ClientContext.ExecuteQueryRetry();
                    if (fileListItem.Exists)
                    {
                        WriteObject(fileListItem.ListItemAllFields);
                    }
                    else
                    {
                        if (ThrowExceptionIfFileNotFound)
                        {
                            throw new PSArgumentException($"No file found with the provided Url {serverRelativeUrl}", "Url");
                        }
                    }
                    break;
                case URLASSTRING:
                    serverRelativeUrl = Utilities.FileUrlResolver.Resolve(Url, webUrl, ClientContext, CurrentWeb);
                    WriteObject(CurrentWeb.GetFileAsString(serverRelativeUrl));
                    break;
                case URLASMEMORYSTREAM:
                    IFile fileMemoryStream;

                    try
                    {
                        fileMemoryStream = Utilities.FileUrlResolver.ResolveFile(Url, webUrl, ClientContext, CurrentWeb, Connection.PnPContext, f => f.Author, f => f.Length, f => f.ModifiedBy, f => f.Name, f => f.TimeCreated, f => f.TimeLastModified, f => f.Title);
                    }
                    catch (ServerException)
                    {
                        // Assume the cause of the exception is that a principal cannot be found and try again without:
                        // Fallback in case the creator or person having last modified the file no longer exists in the environment such that the file can still be downloaded
                        fileMemoryStream = Utilities.FileUrlResolver.ResolveFile(Url, webUrl, ClientContext, CurrentWeb, Connection.PnPContext, f => f.Length, f => f.Name, f => f.TimeCreated, f => f.TimeLastModified, f => f.Title);
                    }

                    var stream = new System.IO.MemoryStream(fileMemoryStream.GetContentBytes());
                    WriteObject(stream);
                    break;
            }
        }

        private void ExecuteThroughGraph()
        {
            // The shares endpoint resolves the file from its URL alone, so no site, drive or list id has to be looked up first
            var shareId = EncodeSharingUrl(GetAbsoluteFileUrl());
            LogDebug($"Addressing the file through Microsoft Graph as share {shareId}");

            switch (ParameterSetName)
            {
                case URLTOPATH:
                    var driveItemToDownload = GetDriveItemJson(shareId);
                    var name = driveItemToDownload.TryGetProperty("name", out JsonElement nameElement) ? nameElement.GetString() : null;
                    string fileOut = System.IO.Path.Combine(Path, !string.IsNullOrEmpty(Filename) ? Filename : name);

                    if (System.IO.File.Exists(fileOut) && !Force)
                    {
                        LogWarning($"File '{System.IO.Path.GetFileName(fileOut)}' exists already. Use the -Force parameter to overwrite the file.");
                    }
                    else
                    {
                        using var response = DownloadContent(driveItemToDownload);
                        using var downloadedContentStream = response.Content.ReadAsStream();
                        using var content = System.IO.File.Create(fileOut);
                        downloadedContentStream.CopyTo(content, 2 * 1024 * 1024);
                    }
                    break;

                case URLASFILEOBJECT:
                    WriteObject(GraphRequestHelper.Get<DriveItem>($"v1.0/shares/{shareId}/driveItem"));
                    break;

                case URLASLISTITEM:
                    JsonElement listItem;
                    try
                    {
                        listItem = JsonSerializer.Deserialize<JsonElement>(GraphRequestHelper.Get($"v1.0/shares/{shareId}/listItem?$expand=fields"));
                    }
                    catch (GraphException ex) when (ex.HttpResponse?.StatusCode == HttpStatusCode.NotFound && !ThrowExceptionIfFileNotFound)
                    {
                        return;
                    }
                    WriteObject(ConvertToGraphListItem(listItem));
                    break;

                case URLASSTRING:
                    using (var response = DownloadContent(GetDriveItemJson(shareId)))
                    {
                        WriteObject(response.Content.ReadAsStringAsync().GetAwaiter().GetResult());
                    }
                    break;

                case URLASMEMORYSTREAM:
                    using (var response = DownloadContent(GetDriveItemJson(shareId)))
                    {
                        WriteObject(new System.IO.MemoryStream(response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult()));
                    }
                    break;
            }
        }

        /// <summary>
        /// Builds the absolute URL of the file from the URL as provided, without querying SharePoint for the web URL
        /// </summary>
        private string GetAbsoluteFileUrl()
        {
            Uri absoluteUri;
            if (Uri.TryCreate(Url, UriKind.Absolute, out Uri providedUri) && (providedUri.Scheme == Uri.UriSchemeHttps || providedUri.Scheme == Uri.UriSchemeHttp))
            {
                absoluteUri = providedUri;
            }
            else
            {
                var connectionUri = new Uri(Connection.Url);
                absoluteUri = Url.StartsWith('/')
                    ? new Uri($"{connectionUri.Scheme}://{connectionUri.Authority}{Url}")
                    : new Uri($"{connectionUri.GetLeftPart(UriPartial.Path).TrimEnd('/')}/{Url}");
            }

            // Re-encode every segment so the URL is encoded exactly once, whether it was passed encoded or not
            var path = string.Join("/", absoluteUri.AbsolutePath.Split('/').Select(segment => Uri.EscapeDataString(Uri.UnescapeDataString(segment))));
            return $"{absoluteUri.Scheme}://{absoluteUri.Authority}{path}";
        }

        /// <summary>
        /// Encodes a URL into a Microsoft Graph sharing token, see https://learn.microsoft.com/graph/api/shares-get#encoding-sharing-urls
        /// </summary>
        private static string EncodeSharingUrl(string url)
        {
            var base64Value = Convert.ToBase64String(Encoding.UTF8.GetBytes(url));
            return "u!" + base64Value.TrimEnd('=').Replace('/', '_').Replace('+', '-');
        }

        private JsonElement GetDriveItemJson(string shareId)
        {
            return JsonSerializer.Deserialize<JsonElement>(GraphRequestHelper.Get($"v1.0/shares/{shareId}/driveItem"));
        }

        /// <summary>
        /// Downloads the content of a drive item through its pre-authenticated download URL, which needs no Authorization header
        /// </summary>
        private HttpResponseMessage DownloadContent(JsonElement driveItem)
        {
            if (!driveItem.TryGetProperty("@microsoft.graph.downloadUrl", out JsonElement downloadUrlElement) || string.IsNullOrEmpty(downloadUrlElement.GetString()))
            {
                throw new PSArgumentException($"The item at {Url} has no content to download. Ensure it is a file and not a folder.", nameof(Url));
            }

            var response = Connection.HttpClient.GetAsync(downloadUrlElement.GetString(), HttpCompletionOption.ResponseHeadersRead).GetAwaiter().GetResult();
            if (!response.IsSuccessStatusCode)
            {
                response.Dispose();
                throw new PSInvalidOperationException($"Downloading the content of {Url} failed with status code {(int)response.StatusCode} {response.StatusCode}");
            }
            return response;
        }

        private static GraphListItem ConvertToGraphListItem(JsonElement listItem)
        {
            var result = new GraphListItem
            {
                Id = GetString(listItem, "id"),
                WebUrl = GetString(listItem, "webUrl"),
                ETag = GetString(listItem, "eTag"),
                CreatedDateTime = listItem.TryGetProperty("createdDateTime", out JsonElement created) && created.TryGetDateTime(out DateTime createdValue) ? createdValue : null,
                LastModifiedDateTime = listItem.TryGetProperty("lastModifiedDateTime", out JsonElement modified) && modified.TryGetDateTime(out DateTime modifiedValue) ? modifiedValue : null,
                Fields = new Hashtable(StringComparer.OrdinalIgnoreCase)
            };

            if (listItem.TryGetProperty("fields", out JsonElement fields))
            {
                foreach (var field in fields.EnumerateObject())
                {
                    if (field.Name.StartsWith('@'))
                    {
                        continue;
                    }
                    result.Fields[field.Name] = field.Value.ValueKind switch
                    {
                        JsonValueKind.String => field.Value.GetString(),
                        JsonValueKind.Number => field.Value.TryGetInt64(out long longValue) ? longValue : field.Value.GetDouble(),
                        JsonValueKind.True => true,
                        JsonValueKind.False => false,
                        JsonValueKind.Null => null,
                        _ => field.Value.GetRawText()
                    };
                }
            }

            return result;
        }

        private static string GetString(JsonElement element, string propertyName)
        {
            return element.TryGetProperty(propertyName, out JsonElement value) && value.ValueKind == JsonValueKind.String ? value.GetString() : null;
        }

        private static async Task SaveFileToLocal(IFile fileToDownload, string filePath)
        {
            // Start the download
            using (Stream downloadedContentStream = await fileToDownload.GetContentAsync(true))
            {
                // Download the file bytes in 2MB chunks and immediately write them to a file on disk 
                // This approach avoids the file being fully loaded in the process memory
                var bufferSize = 2 * 1024 * 1024;  // 2 MB buffer

                using (FileStream content = System.IO.File.Create(filePath))
                {
                    byte[] buffer = new byte[bufferSize];
                    int read;
                    while ((read = await downloadedContentStream.ReadAsync(buffer, 0, buffer.Length)) != 0)
                    {
                        content.Write(buffer, 0, read);
                    }
                }
            }
        }
    }
}
