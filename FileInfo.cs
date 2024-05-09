// SPDX-License-Identifier: GPL-3.0-or-later

namespace Sharepoint2Aria {
    public class FileInfo {
        public string[] PathComponents;
        public string Uid;
        public string DownloadLink;
        public FileInfo(string server_relpath, string api_url, string uid) {
            if (!server_relpath.StartsWith("/")) {
                throw new Exception("Invalid server_relpath: " + server_relpath);
            }
            int domain_len = api_url.IndexOf("/personal/");
            if (domain_len < 0) domain_len = api_url.IndexOf("/sites/");
            if (domain_len < 0) {
                throw new Exception("Invalid API url: " + api_url);
            }

            PathComponents = server_relpath.Substring(1).Split("/");
            Uid = uid;
            DownloadLink = $"{api_url}/web/GetFileById('{uid}')/$value";
        }

        public static string TruncateIfTooLong(string s, bool keepExt) {
            if (System.Text.Encoding.UTF8.GetBytes(s).Length < 240) return s;
            
            string stem = s;
            string suffix = "";
            if (keepExt) {
                int dotPos = s.LastIndexOf('.');
                if (dotPos >= 0) {
                    stem = s.Substring(0, dotPos);
                    suffix = s.Substring(dotPos);
                }
            }

            for (var stem_len = stem.Length; stem_len > 0; stem_len--) {
                string ret = $"{stem.Substring(0, stem_len)}(omit){suffix}";
                if (System.Text.Encoding.UTF8.GetBytes(ret).Length < 240) return ret;
            }

            throw new Exception("File name cannot be truncated: " + s);
        }

        // Remove special character for win10.
        // Truncate utf8 > 255 strings for *nix.
        public string[] GetSafePath() {
            string[] ret = new string[PathComponents.Length];
            for (int i = 0; i < PathComponents.Length; i++) {
                string win_replaced = PathComponents[i]
                    .Replace("<", "＜")
                    .Replace(">", "＞")
                    .Replace(":", "：")
                    .Replace("\"", "”")
                    .Replace("/", "／")
                    .Replace("\\", "＼")
                    .Replace("|", "｜")
                    .Replace("?", "？")
                    .Replace("*", "＊");
                ret[i] = TruncateIfTooLong(win_replaced, i == PathComponents.Length-1);
            }
            return ret;
        }
    }
}
