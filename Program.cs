// SPDX-License-Identifier: GPL-3.0-or-later

namespace Sharepoint2Aria {
    class Program {
        static void Main(string[] args) {
            string rpc_url = "http://[::1]:6800";
            string rpc_secret = "";
            string od_url = "";
            string od_pwd = "";
            bool print_help = false;
            bool select_files = false;

            for (int i = 0; i < args.Length; i++) {
                string arg = args.ElementAt(i);
                if (arg == "--help" || arg == "-h") {
                    print_help = true;
                } else if (arg == "--select-files") {
                    select_files = true;
                } else if (arg == "--rpc-secret") {
                    rpc_secret = args.ElementAt(++i);
                } else if (arg == "--rpc-url") {
                    rpc_url = args.ElementAt(++i);
                } else if (od_url == "") {
                    od_url = arg;
                } else {
                    od_pwd = arg;
                }
            }

            if (print_help || od_url == "") {
                Console.WriteLine("Sharepoint2Aria.exe [options] <share_link> [share_password]");
                Console.WriteLine("    --rpc-url <string>      Aria2 RPC URL. Default: http://[::1]:6800");
                Console.WriteLine("    --rpc-secret <string>   Aria2 RPC secret.");
                Console.WriteLine("    --select-files          Interactively prompt for files to download.");
                return;
            }

            AriaRpcClient aria2 = new AriaRpcClient(rpc_url, rpc_secret);
            aria2.Ping();

            Sharepoint sp = new Sharepoint(od_url, od_pwd);
            Console.WriteLine($"Api:      {sp.api_}");
            Console.WriteLine($"Cookie:   FedAuth={sp.fedauth_}");
            Console.WriteLine($"UniqueId: {sp.relpath_base_uid_}");
            Console.WriteLine($"Type:     " + (sp.relpath_base_isfile_ ? "File" : "Folder"));
            var files = sp.ListFiles();
            files.Sort((a, b) => {
                string apath = string.Join("/", a.PathComponents);
                string bpath = string.Join("/", b.PathComponents);
                return apath.CompareTo(bpath);
            });

            Console.WriteLine($"\n\nFound {files.Count} files:");
            for (int i = 0; i < files.Count; i++) {
                string path = "/" + string.Join("/", files[i].PathComponents);
                Console.WriteLine($"[{i,3}] {path}");
            }

            // Select files
            int max_index = files.Count - 1;
            HashSet<int> selected_index_set = new HashSet<int>();
            if (select_files) {
                bool ok = false;
                while (!ok) {
                    Console.Write("Select files for download (e.g. 1,3,5-7 or all or exit): ");
                    Console.Out.Flush();
                    string? input = Console.ReadLine();
                    if (input == null || input == "exit") return;
                    if (input == "") continue;
                    if (input == "all") break;

                    ok = true;
                    foreach (string r in input.Split(",")) {
                        if (r.Contains("-")) {
                            var x = r.Split("-");
                            int begin = Int32.Parse(x[0]);
                            int end = Int32.Parse(x[1]);
                            if (begin < 0 || begin > max_index || end < 0 || end > max_index || end < begin) {
                                Console.WriteLine("Invalid index: " + input);
                                ok = false;
                                break;
                            }
                            for (int j = begin; j <= end; j++) {
                                selected_index_set.Add(j);
                            }
                        } else {
                            int n = Int32.Parse(r);
                            if (n < 0 || n > max_index) {
                                Console.WriteLine("Invalid index: " + input);
                                ok = false;
                                break;
                            }
                            selected_index_set.Add(n);
                        }
                    }
                }
            }
            if (selected_index_set.Count == 0) {
                Console.WriteLine("Selecting all files for download ...");
                for (int i = 0; i < files.Count; i++) {
                    selected_index_set.Add(i);
                }
            }
            List<int> selected_index = new List<int>(selected_index_set);
            selected_index.Sort();
            Console.WriteLine("Adding tasks:");
            foreach (int idx in selected_index) {
                string path = "/" + string.Join("/", files[idx].PathComponents);
                var safe_path = files[idx].GetSafePath();
                var gid = aria2.AddUri(files[idx].DownloadLink, safe_path[safe_path.Count() - 1], null, new Dictionary<string, string>() { { "FedAuth", sp.fedauth_ } });
                if (gid == null || gid == "") {
                    Console.WriteLine($"FAIL [{idx,3}] {path}");
                } else {
                    Console.WriteLine($"OK   [{idx,3}] {path}");
                }
            }
        }
    }
}
