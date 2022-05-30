// SPDX-License-Identifier: GPL-3.0-or-later

using System.Text.Json;
using System.Text.RegularExpressions;
using System.Web;

namespace Sharepoint2Aria {
    public class Sharepoint {
        public string initial_url_;
        public Uri real_url_;
        public string relpath_base_;  // "/personal/.../Documents/...";
        public string relpath_base_uid_;
        public bool relpath_base_isfile_;
        public string api_; // "https://...sharepoint.com/personal/.../_api";
        public string fedauth_;  // "77u..."

        private JsonElement HttpGet(string url) {
            var client_handler = new HttpClientHandler();
            client_handler.AllowAutoRedirect = true;

            var req = new HttpRequestMessage(HttpMethod.Get, url);
            req.Headers.Add("Accept", "application/json");
            req.Headers.Add("Cookie", $"FedAuth={fedauth_}");

            HttpClient client = new HttpClient(client_handler);
            var reqTask = client.SendAsync(req);
            if (!reqTask.Wait(new TimeSpan(0, 0, 5))) {
                throw new Exception($"HTTP request timed out: ${url}");
            }
            HttpResponseMessage rsp = reqTask.Result;
            if (rsp.StatusCode != System.Net.HttpStatusCode.OK) {
                Console.WriteLine(rsp);
                Console.WriteLine(rsp.Content.ReadAsStringAsync().Result);
                throw new Exception($"Unexpected HTTP status code: ${rsp.StatusCode}");
            }

            string rsp_body = rsp.Content.ReadAsStringAsync().Result;
            var jd = JsonDocument.Parse(rsp_body) ?? throw new Exception("Failed to parse reply: " + rsp_body);
            return jd.RootElement;
        }

        public Sharepoint(string url) {
            // Check URL format
            // f for folder, i for item?
            Regex re = new Regex(@"^https://[a-z-]+\.sharepoint\.com/:(i|f):/g/personal/.*$");
            MatchCollection matches = re.Matches(url);
            if (matches.Count == 0) {
                throw new ArgumentException($"Invalid url: ${url}");
            }
            initial_url_ = url;

            // HTTP request
            var client_handler = new HttpClientHandler();
            client_handler.AllowAutoRedirect = false;
            HttpClient client = new HttpClient(client_handler);
            Console.WriteLine("Waiting for the FedAuth cookie...");
            var httpGetTask = client.GetAsync(url);
            if (!httpGetTask.Wait(new TimeSpan(0, 0, 5))) {
                throw new Exception($"HTTP request timed out: ${url}");
            }
            HttpResponseMessage rsp = httpGetTask.Result;
            if (rsp.StatusCode != System.Net.HttpStatusCode.Found) {
                throw new Exception($"Unexpected HTTP status code: {rsp.StatusCode} url={url}");
            }
            real_url_ = rsp.Headers.Location ?? throw new Exception("HTTP 302 not contain new location");

            // Extract server relative path
            var relpath = System.Web.HttpUtility.ParseQueryString(real_url_.Query)["id"];
            relpath_base_ = relpath ?? throw new Exception("New location isn't expected " + real_url_);

            // Construct API path
            string s = real_url_.ToString();
            int pos = s.IndexOf("/_layouts/");
            if (pos < 0) throw new Exception("Unexpected url: " + s);
            api_ = s.Substring(0, pos) + "/_api";

            // Extract FedAuth
            fedauth_ = "";
            re = new Regex(@"FedAuth=([0-9a-zA-Z/+]+)");
            foreach (var cookie in rsp.Headers) {
                if (cookie.Key != "Set-Cookie") continue;
                foreach (string val in cookie.Value) {
                    MatchCollection ms = re.Matches(val);
                    if (ms.Count == 0) continue;
                    Match m = ms[0];
                    fedauth_ = m.Groups[1].ToString();
                }
            }
            if (fedauth_ == "") {
                throw new Exception("Cannot find FedAuth in response");
            }

            // Print cookie valid date
            // string decodedfedauth = System.Text.Encoding.UTF8.GetString(System.Convert.FromBase64String(fedauth_));
            // DateTime startTime = DateTime.FromFileTime(Int64.Parse(decodedfedauth.Split("|")[4].Split(",")[1]));
            // DateTime endTime = DateTime.FromFileTime(Int64.Parse(decodedfedauth.Split("|")[4].Split(",")[3]));
            // Console.WriteLine($"Valid since: {startTime.ToString()}");
            // Console.WriteLine($"Now time:    {DateTime.Now.ToString()}");
            // Console.WriteLine($"Valid until: {endTime.ToString()}");

            // Get base uid
            url = $"{api_}/web/GetSharingLinkData(@Link)?@Link='{HttpUtility.UrlEncode(initial_url_)}'";
            var ret = HttpGet(url);
            string debugjson = JsonSerializer.Serialize(ret, new JsonSerializerOptions() { WriteIndented = true });
            relpath_base_uid_ = ret.GetProperty("ObjectUniqueId").GetString() ?? throw new Exception("bad uid: " + debugjson);
            int type = ret.GetProperty("ObjectType").GetInt32();
            if (type == 1) {
                relpath_base_isfile_ = true;
            } else if (type == 2) {
                relpath_base_isfile_ = false;
            } else {
                throw new Exception("invalid type: " + debugjson);
            }
        }

        public List<FileInfo> ListFiles() {
            if (relpath_base_isfile_) {
                // Special case: Top level is a file.
                string url = $"{api_}/web/GetFileById('{relpath_base_uid_}')";
                var f = HttpGet(url);
                String debugjson = JsonSerializer.Serialize(f, new JsonSerializerOptions { WriteIndented = true });
                string file_name = f.GetProperty("Name").GetString() ?? throw new Exception($"bad file: {debugjson}");
                string file_relpath = f.GetProperty("ServerRelativeUrl").GetString() ?? throw new Exception($"bad file: {debugjson}");
                Console.WriteLine($"File UID={relpath_base_uid_} Name={file_name} Rel={file_relpath}");
                return new List<FileInfo> { new FileInfo(file_relpath, api_, relpath_base_uid_) };
            }

            // Walk through dirs using BFS
            Queue<string> q = new Queue<string>(); // UID of folders
            List<FileInfo> ret = new List<FileInfo>();
            q.Enqueue(relpath_base_uid_);

            while (q.Count > 0) {
                string this_dir_uid = q.Dequeue();

                {   // Folder info
                    string dirinfo_url = $"{api_}/web/GetFolderById('{this_dir_uid}')";
                    var dirinfo_obj = HttpGet(dirinfo_url);
                    string debugjson = JsonSerializer.Serialize(dirinfo_obj, new JsonSerializerOptions { WriteIndented = true });
                    string dir_name = dirinfo_obj.GetProperty("Name").GetString() ?? throw new Exception($"bad folder: {debugjson}");
                    string dir_relpath = dirinfo_obj.GetProperty("ServerRelativeUrl").GetString() ?? throw new Exception($"bad Folder: {debugjson}");
                    Console.WriteLine($"Scanning {dir_relpath}");
                }

                {   // Enqueue subdirs
                    var subdirs_obj = HttpGet($"{api_}/web/GetFolderById('{this_dir_uid}')/Folders");
                    var debugjson = JsonSerializer.Serialize(subdirs_obj, new JsonSerializerOptions { WriteIndented = true });
                    foreach (JsonElement subdir_obj in subdirs_obj.GetProperty("value").EnumerateArray()) {
                        string subdir_uid = subdir_obj.GetProperty("UniqueId").GetString() ?? throw new Exception($"bad subdir: {debugjson}");
                        q.Enqueue(subdir_uid);
                    }
                }

                {   // Enqueue subfiles
                    var subfiles_obj = HttpGet($"{api_}/web/GetFolderById('{this_dir_uid}')/Files");
                    var debugjson = JsonSerializer.Serialize(subfiles_obj, new JsonSerializerOptions { WriteIndented = true });
                    foreach (JsonElement subfile_obj in subfiles_obj.GetProperty("value").EnumerateArray()) {
                        string subfile_uid = subfile_obj.GetProperty("UniqueId").GetString() ?? throw new Exception($"bad subfile: {debugjson}");
                        string subfile_name = subfile_obj.GetProperty("Name").GetString() ?? throw new Exception($"bad subfile: {debugjson}");
                        string subfile_relpath = subfile_obj.GetProperty("ServerRelativeUrl").GetString() ?? throw new Exception($"bad subfile: {debugjson}");
                        Console.WriteLine($"  Found file {subfile_relpath} UID={subfile_uid}");
                        ret.Add(new FileInfo(subfile_relpath, api_, subfile_uid));
                    }
                }
            }
            return ret;
        }
    }
}