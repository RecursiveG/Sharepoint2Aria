// SPDX-License-Identifier: GPL-3.0-or-later

using System.Text.Json;

namespace Sharepoint2Aria {
    public class AriaRpcClient {
        public string url_;
        public string secret_;
        public int seqnum_;

        // Only supports JSONRPC over HTTP.
        // URL must be form "http://<server>:<port>"
        public AriaRpcClient(string url, string secret) {
            url_ = url + "/jsonrpc";
            secret_ = "token:" + secret;
            seqnum_ = 0;
        }

        private JsonElement? Request(string payload) {
            HttpClient client = new HttpClient();
            HttpContent content = new StringContent(payload, System.Text.Encoding.UTF8, "application/json");
            var postTask = client.PostAsync(url_, content);
            if (!postTask.Wait(new TimeSpan(0, 0, 5))) {
                throw new Exception("Aria2 JSONRPC timed out: " + url_);
            }
            var rsp = postTask.Result;
            if (!rsp.IsSuccessStatusCode) {
                Console.WriteLine($"Error\n{payload}\n{rsp.ToString()}");
                return null;
            }
            string rsp_body = rsp.Content.ReadAsStringAsync().Result;
            var jd = JsonDocument.Parse(rsp_body);
            return jd.RootElement;
        }

        // Download a new file. Returns the GID.
        public string AddUri(string url, string? filename = null, string? dir = null, Dictionary<string, string>? cookies = null) {
            // Build cookie list
            List<string> headers = new List<string>();  // ["Cookie: FedAuth=...", ...]
            if (cookies != null) {
                foreach (var kv in cookies) {
                    headers.Add($"Cookie: {kv.Key}={kv.Value}");
                }
            }

            // Build options
            var options = new Dictionary<string, Object>();
            if (filename != null) options.Add("out", filename);
            if (dir != null) options.Add("dir", dir);
            if (headers.Count > 0) options.Add("header", headers);

            // Build params
            Object[] p;
            if (options.Count == 0) {
                p = new Object[] {
                    secret_,
                    new[] {url},
                };
            } else {
                p = new Object[] {
                    secret_,
                    new[] {url},
                    options,
                };
            }

            // Build JSON payload
            var payload = new Dictionary<string, Object> {
                ["jsonrpc"] = "2.0",
                ["method"] = "aria2.addUri",
                ["id"] = seqnum_++,
                ["params"] = p,
            };

            var jsonopt = new JsonSerializerOptions { WriteIndented = true };
            var rsp = Request(JsonSerializer.Serialize(payload, jsonopt));
            if (rsp == null) {
                return "";
            }
            var x = rsp.Value.GetProperty("result").GetString();
            if (x == null) {
                Console.WriteLine($"unexpected response {rsp.Value.ToString()}");
                return "";
            }
            return x;
        }

        // Use getGlobalStat to test connectivity.
        public void Ping() {
            // Build JSON payload
            var payload = new Dictionary<string, Object> {
                ["jsonrpc"] = "2.0",
                ["method"] = "aria2.getGlobalStat",
                ["id"] = seqnum_++,
                ["params"] = new object[] {secret_},
            };
            var rsp = Request(JsonSerializer.Serialize(payload));
            if (rsp == null) throw new Exception("Aria2 RPC not reachable.");
        }
    }
}