using DocumentFormat.OpenXml.Drawing.Diagrams;
using ICSharpCode.SharpZipLib.Zip;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyTemplate.MyClass
{
    /// <summary>
    /// Dispose() が呼ばれても、親ストリーム（_innerStream）を閉じないためのラッパ
    /// StreamWriter や BinaryWriter は Dispose() 時に基底ストリームも閉じる仕様なので
    /// ZipOutputStream を直接渡すと CloseEntry() 前に ZipOutputStream が閉じられてしまい
    /// "Cannot access a closed file." 例外が発生する
    /// </summary>
    public class  NonClosingStreamWrapper : Stream
    {
        private readonly Stream _innerStream;
        public NonClosingStreamWrapper(Stream inner) => _innerStream = inner;

        public override bool CanRead => _innerStream.CanRead;
        public override bool CanSeek => _innerStream.CanSeek;
        public override bool CanWrite => _innerStream.CanWrite;
        public override long Length => _innerStream.Length;
        public override long Position
        {
            get => _innerStream.Position;
            set => _innerStream.Position = value;
        }

        public override void Flush() => _innerStream.Flush();
        public override int Read(byte[] buffer, int offset, int count) => _innerStream.Read(buffer, offset, count);
        public override long Seek(long offset, SeekOrigin origin) => _innerStream.Seek(offset, origin);
        public override void SetLength(long value) => _innerStream.SetLength(value);
        public override void Write(byte[] buffer, int offset, int count) => _innerStream?.Write(buffer, offset, count);

        protected override void Dispose(bool disposing)
        {
            // 親ストリームは閉じない（Disposeしない）
            // _innerStream.Dispose(); は呼ばない
        }

    }


    public class MyArchiveWriter : IDisposable
    {
        private readonly bool _useZip;
        private readonly string _baseDirectory;
        private readonly Stream _fileStream;
        private ZipOutputStream _zipStream;
        private bool _entryOpen;
        private readonly bool _failIfExists;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="outputPath"></param>
        /// <param name="useZip"></param>
        /// <param name="failIfExists"></param>
        public MyArchiveWriter(string outputPath, bool useZip, bool failIfExists = true)
        {
            _useZip = useZip;
            _failIfExists = failIfExists;

            if (_useZip)
            {
                var mode = failIfExists ? FileMode.CreateNew : FileMode.Create;
                _fileStream = new FileStream(outputPath, mode, FileAccess.Write);
                _zipStream = new ZipOutputStream(_fileStream);
                _zipStream.SetLevel(9); // 最高圧縮率
            }
            else
            {
                _baseDirectory = Path.GetDirectoryName(outputPath) ?? ".";
            }
        }

        /// <summary>
        /// 新しいエントリーを作成し書き込み用のストリームを返す
        /// </summary>
        /// <param name="entryName"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException"></exception>
        public Stream CreateEntry(string entryName, string? password = null)
        {
            if (_useZip)
            {
                if (_entryOpen) throw new InvalidOperationException("前のエントリーを閉じてください。");

                entryName = entryName.Replace('\\', '/'); // Zipではスラッシュを使用
                var entry = new ZipEntry(entryName)
                {
                    DateTime = DateTime.Now,
                };

                _zipStream.Password = password; // パスワード設定
                _zipStream.PutNextEntry(entry);
                _entryOpen = true;

                // ZipOutputStreamはDispose()時に親ストリームを閉じるため、ラップして閉じないようにする
                return new NonClosingStreamWrapper(_zipStream);
            }
            else
            {
                // 通常の場合は指定パスにファイルを作成
                var filePath = Path.Combine(_baseDirectory, entryName);
                var dirPath = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(dirPath))
                    Directory.CreateDirectory(dirPath); // ディレクトリが存在しない場合は作成

                var mode = _failIfExists ? FileMode.CreateNew : FileMode.Create;
                return new FileStream(filePath, mode, FileAccess.Write);
            }
        }

        /// <summary>
        /// ZIPエントリーを閉じる
        /// </summary>
        public void CloseEntry()
        {
            if (_useZip || _entryOpen)
            {
                _zipStream.CloseEntry();
                _entryOpen = false;
            }
        }

        /// <summary>
        /// ファイルを書き込む
        /// </summary>
        /// <param name="entryName"></param>
        /// <param name="sourceFilePath"></param>
        /// <param name="password"></param>
        public void WriteFile(string entryName, string sourceFilePath, string? password = null)
        {
            using (var entryStream = CreateEntry(entryName, password))
            using (var fs = File.OpenRead(sourceFilePath))
            {
                fs.CopyTo(entryStream);
            }
            CloseEntry();
        }

        /// <summary>
        /// Byte配列を書き込む
        /// </summary>
        /// <param name="entryName"></param>
        /// <param name="data"></param>
        /// <param name="password"></param>
        /// <exception cref="ArgumentNullException"></exception>
        public void WriteBytes(string entryName, byte[] data, string? password = null)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));

            using (var entryStream = CreateEntry(entryName, password))
            {
                entryStream.Write(data, 0, data.Length);
            }
            CloseEntry();
        }

        /// <summary>
        /// 空のフォルダを作成
        /// </summary>
        /// <param name="folderName"></param>
        /// <exception cref="InvalidOperationException"></exception>
        public void CreateFolder(string folderName)
        {
            if (_useZip)
            {
                if (_entryOpen) throw new InvalidOperationException("前のエントリーを閉じてください。");

                folderName = folderName.Replace('\\', '/'); // Zipではスラッシュを使用

                if (!folderName.EndsWith("/"))
                {
                    folderName += "/"; // フォルダ名はスラッシュで終わる必要がある
                }

                var entry = new ZipEntry(folderName)
                {
                    DateTime = DateTime.Now,
                };
                _zipStream.PutNextEntry(entry);
                _zipStream.CloseEntry();
            }
            else
            {
                // 通常の場合はディレクトリを作成
                var dirPath = Path.Combine(_baseDirectory, folderName);
                Directory.CreateDirectory(dirPath);
            }
        }

        /// <summary>
        /// 指定したフォルダの内容を追加する
        /// デフォルトでフォルダ構造を維持
        /// </summary>
        /// <param name="folderPath">追加するフォルダのパス</param>
        /// <param name="basePathInZip">ZIP内のフォルダパス（空ならルート）</param>
        /// <param name="includeSubDirs">サブフォルダも含めるか</param>
        /// <param name="password"></param>
        /// <exception cref="DirectoryNotFoundException"></exception>
        public void AddFolder(string folderPath, string basePathInZip = "", bool includeSubDirs = true, string? password = null)
        {
            if (!Directory.Exists(folderPath)) throw new DirectoryNotFoundException($"指定されたフォルダが存在しません: {folderPath}");

            var searchOption = includeSubDirs ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            var files = Directory.GetFiles(folderPath, "*", searchOption);

            foreach (var file in files)
            {
                var relativePath = Path.GetRelativePath(folderPath, file);
                if (!string.IsNullOrEmpty(basePathInZip))
                {
                    relativePath = Path.Combine(basePathInZip, relativePath);
                    relativePath = relativePath.Replace('\\', '/'); // Zipではスラッシュを使用
                    WriteFile(relativePath, file, password);
                }
            }
        }

        /// <summary>
        /// 開放
        /// </summary>
        public void Dispose()
        {
            if (_useZip)
            {
                if (_entryOpen) CloseEntry();
                _zipStream?.Finish();
                _zipStream?.Close();
                _zipStream?.Dispose();
            }
        }
    }
}
