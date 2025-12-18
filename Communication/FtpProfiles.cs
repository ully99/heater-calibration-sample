using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Net;

namespace Heater_Cal_Demo_P4.Communication
{
    public static class FtpProfiles
    {
        private static Uri BuildUri(string host, int port, bool ssl, string baseDir, string fileNameOrEmpty)
        {
            // FtpWebRequest는 ftps:// 스킴 미지원 → ftp:// + EnableSsl=true (Explicit FTPS)
            const string scheme = "ftp";

            string path = string.Empty;
            if (!string.IsNullOrWhiteSpace(baseDir))
                path = baseDir.Replace('\\', '/').Trim('/');

            if (!string.IsNullOrEmpty(fileNameOrEmpty))
            {
                string f = Uri.EscapeDataString(fileNameOrEmpty);
                path = string.IsNullOrEmpty(path) ? f : (path + "/" + f);
            }
            return new Uri($"{scheme}://{host}:{port}/{path}");
        }

        private static FtpWebRequest Create(string host, int port, bool ssl, string baseDir, string fileOrEmpty,
                                            string method, NetworkCredential cred, int timeoutMs = 10000)
        {
            var uri = BuildUri(host, port, ssl, baseDir, fileOrEmpty);
            var req = (FtpWebRequest)WebRequest.Create(uri);
            req.Method = method;
            req.Credentials = cred;
            req.UsePassive = true;
            req.UseBinary = true;
            req.KeepAlive = false;
            req.EnableSsl = ssl;
            req.Proxy = null;                    // 시스템 프록시 개입 방지
            req.Timeout = timeoutMs;
            req.ReadWriteTimeout = timeoutMs;

            return req;
        }

        // 연결 테스트 (데이터채널 없이 제어채널만 확인)
        public static Task<bool> TestAsync(string host, int port, bool ssl, string baseDir, string user, string pw)
        {
            return Task.Run(() =>
            {
                try
                {
                    var req = Create(host, port, ssl, baseDir, "", WebRequestMethods.Ftp.PrintWorkingDirectory,
                                     new NetworkCredential(user, pw));
                    using (var resp = (FtpWebResponse)req.GetResponse()) { }
                    return true;
                }
                catch { return false; }
            });
        }

        // 디렉터리 만들기(깊이별로 생성, 이미 있으면 무시)
        public static Task EnsureDirAsync(string host, int port, bool ssl, string baseDir, string user, string pw)
        {
            return Task.Run(() =>
            {
                if (string.IsNullOrWhiteSpace(baseDir)) return;

                var cred = new NetworkCredential(user, pw);
                string[] parts = baseDir.Replace('\\', '/').Trim('/').Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
                string cur = "";
                foreach (var part in parts)
                {
                    cur = string.IsNullOrEmpty(cur) ? part : (cur + "/" + part);
                    try
                    {
                        var req = Create(host, port, ssl, cur, "", WebRequestMethods.Ftp.MakeDirectory, cred);
                        using (var resp = (FtpWebResponse)req.GetResponse()) { }
                    }
                    catch (WebException ex)
                    {
                        var r = ex.Response as FtpWebResponse;
                        // 550: already exists → 무시
                        if (r == null || r.StatusCode != FtpStatusCode.ActionNotTakenFileUnavailable)
                            throw;
                    }
                }
            });
        }

        // 업로드 (json 문자열 -> 파일)
        public static Task UploadJsonAsync(string host, int port, bool ssl, string baseDir, string user, string pw,
                                           string fileName, string jsonUtf8)
        {
            return Task.Run(() =>
            {
                var bytes = Encoding.UTF8.GetBytes(jsonUtf8);
                var req = Create(host, port, ssl, baseDir, fileName, WebRequestMethods.Ftp.UploadFile, new NetworkCredential(user, pw));
                req.ContentLength = bytes.Length;

                using (var reqStream = req.GetRequestStream())
                    reqStream.Write(bytes, 0, bytes.Length);

                using (var resp = (FtpWebResponse)req.GetResponse()) { }
            });
        }

        // 다운로드 (파일 -> json 문자열)
        public static Task<string> DownloadJsonAsync(string host, int port, bool ssl, string baseDir, string user, string pw, string fileName)
        {
            return Task.Run(() =>
            {
                var req = Create(host, port, ssl, baseDir, fileName, WebRequestMethods.Ftp.DownloadFile, new NetworkCredential(user, pw));
                using (var resp = (FtpWebResponse)req.GetResponse())
                using (var s = resp.GetResponseStream())
                using (var ms = new MemoryStream())
                {
                    s.CopyTo(ms);
                    return Encoding.UTF8.GetString(ms.ToArray());
                }
            });
        }

        // 목록 (파일명들)
        public static Task<string[]> ListAsync(string host, int port, bool ssl, string baseDir, string user, string pw)
        {
            return Task.Run(() =>
            {
                var req = Create(host, port, ssl, baseDir, "", WebRequestMethods.Ftp.ListDirectory, new NetworkCredential(user, pw));
                using (var resp = (FtpWebResponse)req.GetResponse())
                using (var sr = new StreamReader(resp.GetResponseStream(), Encoding.ASCII, true))
                {
                    var all = sr.ReadToEnd();
                    var lines = all.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    return lines;
                }
            });
        }
        public static Task DeleteAsync(string host, int port, bool ssl, string baseDir, string user, string pw, string fileName)
        {
            return Task.Run(() =>
            {
                var req = Create(host, port, ssl, baseDir, fileName, WebRequestMethods.Ftp.DeleteFile, new NetworkCredential(user, pw));
                using (var resp = (FtpWebResponse)req.GetResponse()) { }
            });
        }
    }
}
