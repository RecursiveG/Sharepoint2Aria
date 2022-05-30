// SPDX-License-Identifier: GPL-3.0-or-later

using System.Text.Json;
using System.Text.RegularExpressions;
using System.Web;
using System.Net;

namespace Sharepoint2Aria {
    public class Sharepoint {
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
            // TODO handle 503, too fast?
            if (rsp.StatusCode != System.Net.HttpStatusCode.OK) {
                Console.WriteLine(rsp);
                Console.WriteLine(rsp.Content.ReadAsStringAsync().Result);
                throw new Exception($"Unexpected HTTP status code: ${rsp.StatusCode}");
            }

            string rsp_body = rsp.Content.ReadAsStringAsync().Result;
            var jd = JsonDocument.Parse(rsp_body) ?? throw new Exception("Failed to parse reply: " + rsp_body);
            return jd.RootElement;
        }

        private static string GetFedAuthCookieFromResponse(HttpResponseMessage rsp) {
            Regex re = new Regex(@"FedAuth=([0-9a-zA-Z/+]+)");
            foreach (var cookie in rsp.Headers) {
                if (cookie.Key != "Set-Cookie") continue;
                foreach (string val in cookie.Value) {
                    MatchCollection ms = re.Matches(val);
                    if (ms.Count == 0) continue;
                    Match m = ms[0];
                    return m.Groups[1].ToString();
                }
            }
            return "";
        }

        public Sharepoint(string url, string passwd) {
            // Check URL format
            // f for folder, i for item?
            Regex re = new Regex(@"^https://[a-z-]+\.sharepoint\.com/:(i|f):/g/personal/.*$");
            MatchCollection matches = re.Matches(url);
            if (matches.Count == 0) {
                throw new ArgumentException($"Invalid url: ${url}");
            }

            // Make initial request
            var req = new HttpRequestMessage(HttpMethod.Get, url);
            req.Headers.Add("Accept", "*/*");
            req.Headers.Add("User-Agent", "curl/7.83.1");

            var client_handler = new HttpClientHandler();
            client_handler.AllowAutoRedirect = false;
            HttpClient client = new HttpClient(client_handler);
            Console.WriteLine("Waiting for the FedAuth cookie...");
            var httpGetTask = client.SendAsync(req);
            if (!httpGetTask.Wait(new TimeSpan(0, 0, 5))) {
                throw new Exception($"HTTP request timed out: ${url}");
            }
            HttpResponseMessage rsp = httpGetTask.Result;
            if (rsp.StatusCode != HttpStatusCode.Found && rsp.StatusCode != HttpStatusCode.OK) {
                throw new Exception($"Unexpected HTTP status code: {rsp.StatusCode} url={url}");
            }

            // Links without password redirect us directly.
            else if (rsp.StatusCode == HttpStatusCode.Found) {
                Uri real_url = rsp.Headers.Location ?? throw new Exception("HTTP 302 not contain new location");
                // Extract server relative path
                var relpath = System.Web.HttpUtility.ParseQueryString(real_url.Query)["id"];
                // Construct API path
                string s = real_url.ToString();
                int pos = s.IndexOf("/_layouts/");
                if (pos < 0) throw new Exception("Unexpected url: " + s);
                api_ = s.Substring(0, pos) + "/_api";
                // Extract FedAuth
                fedauth_ = GetFedAuthCookieFromResponse(rsp);
            }

            // Links with password returns a input page.
            else if (rsp.StatusCode == HttpStatusCode.OK) {
                // Extract hidden form fields.
                string html = rsp.Content.ReadAsStringAsync().Result;
                System.IO.File.WriteAllText("aaa.txt", html);
                Regex input_re = new Regex(@"<input .* name=""([^""]+)"" .* value=""([^""]*)"" />");
                MatchCollection mc = input_re.Matches(html);
                Dictionary<string, string> form_values = new Dictionary<string, string>();
                foreach (Match m in mc) {
                    form_values.Add(m.Groups[1].ToString(), m.Groups[2].ToString());
                    //Console.WriteLine(m.Groups[1].ToString() + " = " + m.Groups[2].ToString());
                }
                var expected_fields = new HashSet<string>() { "SideBySideToken", "__VIEWSTATE", "__VIEWSTATEGENERATOR", "__VIEWSTATEENCRYPTED", "__EVENTVALIDATION" };
                if (!expected_fields.SetEquals(form_values.Keys)) {
                    throw new Exception("Unexpected pwd input page");
                }

                // Construct the guestaccess.aspx link
                Regex aspx_re = new Regex(@"action=""(.*guestaccess.aspx[^""]+)""");
                string part = aspx_re.Match(html).Groups[1].ToString().Replace("&amp;", "&");
                int hostname_idx = url.IndexOf("sharepoint.com");
                string guestaccess_url = $"{url.Substring(0, hostname_idx)}sharepoint.com{part}";
                //Console.WriteLine("GuestAccess: " + guestaccess_url);

                // Ask user for passwd
                while (passwd == "") {
                    Console.Write("Input password for protected link: ");
                    Console.Out.Flush();
                    passwd = Console.ReadLine() ?? throw new Exception("User cancelled");
                }
                form_values.Add("txtPassword", passwd);

                // Request w/passwd
                HttpRequestMessage guestaccess_req = new HttpRequestMessage(HttpMethod.Post, guestaccess_url);
                guestaccess_req.Headers.Add("Accept", "*/*");
                guestaccess_req.Headers.Add("User-Agent", "curl/7.83.1");
                guestaccess_req.Content = new FormUrlEncodedContent(form_values);

                HttpClient guestaccess_client = new HttpClient(new HttpClientHandler() { AllowAutoRedirect = true });
                var guestaccess_task = guestaccess_client.SendAsync(guestaccess_req);
                if (!httpGetTask.Wait(new TimeSpan(0, 0, 5))) {
                    throw new Exception($"HTTP request timed out: ${guestaccess_url}");
                }
                HttpResponseMessage guestaccess_rsp = guestaccess_task.Result;
                if (guestaccess_rsp.StatusCode != HttpStatusCode.Found && guestaccess_rsp.StatusCode != HttpStatusCode.OK) {
                    throw new Exception($"Unexpected HTTP status code: {guestaccess_rsp.StatusCode} url={guestaccess_url}");
                }

                fedauth_ = GetFedAuthCookieFromResponse(guestaccess_rsp);
                if (fedauth_ == "") {
                    throw new Exception("Wrong password.");
                }

                int pos = guestaccess_url.IndexOf("/_layouts/");
                if (pos < 0) throw new Exception("Unexpected url: " + guestaccess_url);
                api_ = guestaccess_url.Substring(0, pos) + "/_api";

            } else {
                throw new Exception("What's wrong with the compiler?");
            }

            // Check FedAuth
            if (fedauth_ == "") {
                throw new Exception("Cannot find FedAuth in response");
            }

            // Print cookie time
            string decodedfedauth = System.Text.Encoding.UTF8.GetString(System.Convert.FromBase64String(fedauth_));
            DateTime time1 = DateTime.FromFileTime(Int64.Parse(decodedfedauth.Split("|")[4].Split(",")[1]));
            DateTime time2 = DateTime.FromFileTime(Int64.Parse(decodedfedauth.Split("|")[4].Split(",")[3]));
            Console.WriteLine($"Now time:      {DateTime.Now.ToString()}");
            Console.WriteLine($"Cookie time 1: {time1.ToString()}");
            Console.WriteLine($"Cookie time 2: {time2.ToString()}");

            // Get base uid
            url = $"{api_}/web/GetSharingLinkData(@Link)?@Link='{HttpUtility.UrlEncode(url)}'";
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