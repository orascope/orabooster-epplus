/*******************************************************************************
 * orabooster-epplus: this file is new in the fork (item B5 of issue #8). It is not part of EPPlus 4.5 or EPPlusFree and has
 * no code from EPPlus 5 or later. Written by OraScope Co.,Ltd. 2026, under the license of the library (LGPL-3.0).
 *
 * Saving and loading with asynchronous input and output.
 *
 * What is asynchronous: every read from and write to the stream or file of the caller. A server that does not allow
 * synchronous writes to its response (ASP.NET Core) can save to it, and a slow disk or network does not hold a thread.
 * What is not: making the package. The parts are written and compressed in memory on the calling thread, exactly as Save
 * does it, and a loaded package is read from memory. A program that must not wait for that (a window) can run the call on
 * another thread itself; the library does not start threads.
 *******************************************************************************/
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeOpenXml
{
    public sealed partial class ExcelPackage
    {
        const int AsyncBufferSize = 81920;

        /// <summary>
        /// Makes the package in memory, as <see cref="Save()"/> does (encrypted when a password is set), and gives its bytes.
        /// A stream of the caller that the package was made over is not written here: it is given back to be written to
        /// asynchronously.
        /// </summary>
        private byte[] BuildInMemory(out Stream streamOfCaller)
        {
            streamOfCaller = null;
            if (_stream != null && !(_stream is MemoryStream))
            {
                streamOfCaller = _stream;      // new ExcelPackage(stream) or new ExcelPackage(stream, template)
                _stream = new MemoryStream();
            }
            FileInfo file = File;
            File = null;                       // Save() then writes the whole package, encrypted or not, to _stream
            try
            {
                Save();
            }
            finally
            {
                File = file;
            }
            byte[] bytes = ((MemoryStream)_stream).ToArray();
            if (streamOfCaller != null)
            {
                _stream.Dispose();
                _stream = streamOfCaller;      // as after Save(): the stream of the caller is the stream of the package
            }
            return bytes;
        }

        private static async Task WriteAllAsync(Stream stream, byte[] bytes, CancellationToken cancellationToken)
        {
            if (!stream.CanWrite) throw new Exception("Can not write to outputstream");
            await stream.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
            await stream.FlushAsync(cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Saves the package like <see cref="Save()"/>, with asynchronous writes: to <see cref="File"/> when the package has
        /// one, to the stream the package was made over when it has one. With neither, the package is only made and can be
        /// read from <see cref="Stream"/>. The package is closed after it has been saved.
        /// </summary>
        /// <param name="cancellationToken">Stops the save before the package is made, or while it is written.</param>
        /// <remarks>
        /// The package is made in memory on the calling thread; only the writes are asynchronous.
        /// </remarks>
        public async Task SaveAsync(CancellationToken cancellationToken = default(CancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            FileInfo file = File;
            try
            {
                Stream streamOfCaller;
                byte[] bytes = BuildInMemory(out streamOfCaller);
                if (streamOfCaller != null) await WriteAllAsync(streamOfCaller, bytes, cancellationToken).ConfigureAwait(false);
                if (file != null)
                {
                    using (var fs = new FileStream(file.FullName, FileMode.Create, FileAccess.Write, FileShare.None, AsyncBufferSize, true))
                    {
                        await WriteAllAsync(fs, bytes, cancellationToken).ConfigureAwait(false);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                if (file == null) throw;
                throw new InvalidOperationException(string.Format("Error saving file {0}", file.FullName), ex); // as Save() does
            }
        }

        /// <summary>
        /// Saves the package like <see cref="Save(string)"/>, with asynchronous writes. See <see cref="SaveAsync(CancellationToken)"/>.
        /// </summary>
        /// <param name="password">The password to encrypt the workbook with. It replaces the password of <see cref="Encryption"/>.</param>
        /// <param name="cancellationToken">Stops the save before the package is made, or while it is written.</param>
        public Task SaveAsync(string password, CancellationToken cancellationToken = default(CancellationToken))
        {
            Encryption.Password = password;
            return SaveAsync(cancellationToken);
        }

        /// <summary>
        /// Saves the workbook to a file like <see cref="SaveAs(FileInfo)"/>, writing the file asynchronously.
        /// The package is closed after it has been saved.
        /// </summary>
        /// <param name="file">The file. It is made new.</param>
        /// <param name="cancellationToken">Stops the save before the package is made, or while it is written.</param>
        public Task SaveAsAsync(FileInfo file, CancellationToken cancellationToken = default(CancellationToken))
        {
            File = file;
            return SaveAsync(cancellationToken);
        }

        /// <summary>
        /// Saves the workbook to a file like <see cref="SaveAs(FileInfo, string)"/>, writing the file asynchronously.
        /// </summary>
        /// <param name="file">The file. It is made new.</param>
        /// <param name="password">The password to encrypt the workbook with. It replaces the password of <see cref="Encryption"/>.</param>
        /// <param name="cancellationToken">Stops the save before the package is made, or while it is written.</param>
        public Task SaveAsAsync(FileInfo file, string password, CancellationToken cancellationToken = default(CancellationToken))
        {
            File = file;
            Encryption.Password = password;
            return SaveAsync(cancellationToken);
        }

        /// <summary>
        /// Copies the package to a stream like <see cref="SaveAs(Stream)"/>, with asynchronous writes only: a stream that does not
        /// allow synchronous writes (the response of a web server) can be used. The package is closed after it has been saved.
        /// </summary>
        /// <param name="outputStream">The stream to write the package to. It is written from its position and is not closed.</param>
        /// <param name="cancellationToken">Stops the save before the package is made, or while it is written.</param>
        public async Task SaveAsAsync(Stream outputStream, CancellationToken cancellationToken = default(CancellationToken))
        {
            if (outputStream == null) throw new ArgumentNullException("outputStream");
            cancellationToken.ThrowIfCancellationRequested();
            File = null;
            Stream streamOfCaller;
            byte[] bytes = BuildInMemory(out streamOfCaller);
            if (streamOfCaller != null) await WriteAllAsync(streamOfCaller, bytes, cancellationToken).ConfigureAwait(false);
            if (outputStream != streamOfCaller && outputStream != _stream)
            {
                await WriteAllAsync(outputStream, bytes, cancellationToken).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Copies the package to a stream like <see cref="SaveAs(Stream, string)"/>, with asynchronous writes only.
        /// </summary>
        /// <param name="outputStream">The stream to write the package to.</param>
        /// <param name="password">The password to encrypt the workbook with. It replaces the password of <see cref="Encryption"/>.</param>
        /// <param name="cancellationToken">Stops the save before the package is made, or while it is written.</param>
        public Task SaveAsAsync(Stream outputStream, string password, CancellationToken cancellationToken = default(CancellationToken))
        {
            Encryption.Password = password;
            return SaveAsAsync(outputStream, cancellationToken);
        }

        /// <summary>
        /// Loads a package from a stream like <see cref="Load(Stream)"/>, reading the stream asynchronously.
        /// </summary>
        /// <param name="input">The stream. It is read from its start when it can seek, else from its position, to its end.</param>
        /// <param name="cancellationToken">Stops the load while the stream is read.</param>
        /// <remarks>The stream is read to memory asynchronously; the package is then opened from memory on the calling thread.</remarks>
        public Task LoadAsync(Stream input, CancellationToken cancellationToken = default(CancellationToken))
        {
            return LoadAsync(input, null, cancellationToken);
        }

        /// <summary>
        /// Loads an encrypted package from a stream like <see cref="Load(Stream, string)"/>, reading the stream asynchronously.
        /// </summary>
        /// <param name="input">The stream.</param>
        /// <param name="password">The password to decrypt the document. Null for a package that is not encrypted.</param>
        /// <param name="cancellationToken">Stops the load while the stream is read.</param>
        public async Task LoadAsync(Stream input, string password, CancellationToken cancellationToken = default(CancellationToken))
        {
            if (input == null) throw new ArgumentNullException("input");
            if (!input.CanRead) throw new Exception("Can not read from inputstream");
            cancellationToken.ThrowIfCancellationRequested();
            var memory = new MemoryStream();
            if (input.CanSeek) input.Seek(0, SeekOrigin.Begin); // as Load does
            await input.CopyToAsync(memory, AsyncBufferSize, cancellationToken).ConfigureAwait(false);
            memory.Position = 0;
            if (password == null) Load(memory);
            else Load(memory, password);
        }

        /// <summary>
        /// Loads a package from a file, reading the file asynchronously. The file becomes <see cref="File"/> of the package, as
        /// with <c>new ExcelPackage(file)</c>: <see cref="Save()"/> and <see cref="SaveAsync(CancellationToken)"/> write to it.
        /// </summary>
        /// <param name="file">The file. It must be there.</param>
        /// <param name="cancellationToken">Stops the load while the file is read.</param>
        public Task LoadAsync(FileInfo file, CancellationToken cancellationToken = default(CancellationToken))
        {
            return LoadAsync(file, null, cancellationToken);
        }

        /// <summary>
        /// Loads an encrypted package from a file, reading the file asynchronously. See <see cref="LoadAsync(FileInfo, CancellationToken)"/>.
        /// </summary>
        /// <param name="file">The file. It must be there.</param>
        /// <param name="password">The password to decrypt the document. Null for a package that is not encrypted.</param>
        /// <param name="cancellationToken">Stops the load while the file is read.</param>
        public async Task LoadAsync(FileInfo file, string password, CancellationToken cancellationToken = default(CancellationToken))
        {
            if (file == null) throw new ArgumentNullException("file");
            using (var fs = new FileStream(file.FullName, FileMode.Open, FileAccess.Read, FileShare.Read, AsyncBufferSize, true))
            {
                await LoadAsync(fs, password, cancellationToken).ConfigureAwait(false);
            }
            File = file;
        }
    }
}
