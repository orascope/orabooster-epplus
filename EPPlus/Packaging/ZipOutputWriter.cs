/*******************************************************************************
 * This file is part of orabooster-epplus, a fork of EPPlusFree 4.5.3.8
 * (https://github.com/rimland/EPPlus), itself a continuation of EPPlus 4.5.3.3.
 *
 * GNU Lesser General Public License v3.0 (see license.md / lgpl-3.0.txt).
 *
 * 2026-11-24 OraScope Co.,Ltd.: new file. Writes the package with
 * System.IO.Compression.ZipArchive (native zlib) instead of the managed
 * DotNetZip (Ionic) ZipOutputStream. Saving was dominated by the managed deflate.
 *******************************************************************************/
using System;
using System.IO;
using System.IO.Compression;

namespace OfficeOpenXml.Packaging
{
    /// <summary>
    /// A write-only stream with the same usage as Ionic's ZipOutputStream:
    /// set <see cref="CompressionLevel"/>, call <see cref="PutNextEntry"/>, then write the entry.
    /// The previous entry is closed when the next one starts or when this stream is closed.
    /// </summary>
    internal sealed class ZipOutputWriter : Stream
    {
        readonly ZipArchive _archive;
        Stream _entry;
        bool _closed;

        internal ZipOutputWriter(Stream a_output, bool a_leaveOpen)
        {
            _archive = new ZipArchive(a_output, ZipArchiveMode.Create, a_leaveOpen);
        }

        /// <summary>Compression of the entries started after this is set.</summary>
        internal OfficeOpenXml.CompressionLevel CompressionLevel { get; set; } = OfficeOpenXml.CompressionLevel.Default;

        /// <summary>Kept for the callers written for Ionic. The native deflater picks its own buffer size.</summary>
        internal int CodecBufferSize { get; set; }

        internal void PutNextEntry(string a_name)
        {
            CloseEntry();
            // Part names in the package are relative URIs ("/xl/workbook.xml"); zip entry names have no leading slash.
            string l_name = a_name.StartsWith("/", StringComparison.Ordinal) ? a_name.Substring(1) : a_name;
            ZipArchiveEntry l_entry = _archive.CreateEntry(l_name, Map(CompressionLevel));
            _entry = l_entry.Open();
        }

        static System.IO.Compression.CompressionLevel Map(OfficeOpenXml.CompressionLevel a_level)
        {
            if (a_level == OfficeOpenXml.CompressionLevel.None) return System.IO.Compression.CompressionLevel.NoCompression;
            if (a_level <= OfficeOpenXml.CompressionLevel.Level3) return System.IO.Compression.CompressionLevel.Fastest;
            return System.IO.Compression.CompressionLevel.Optimal; // zlib level 6, the same as Ionic's Default
        }

        void CloseEntry()
        {
            if (_entry != null)
            {
                _entry.Dispose();
                _entry = null;
            }
        }

        Stream Current
        {
            get
            {
                if (_entry == null) throw new InvalidOperationException("PutNextEntry must be called before writing.");
                return _entry;
            }
        }

        public override void Write(byte[] buffer, int offset, int count) { Current.Write(buffer, offset, count); }
        public override void Write(ReadOnlySpan<byte> buffer) { Current.Write(buffer); }
        public override void WriteByte(byte value) { Current.WriteByte(value); }
        public override void Flush() { _entry?.Flush(); }

        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => !_closed;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) { throw new NotSupportedException(); }
        public override long Seek(long offset, SeekOrigin origin) { throw new NotSupportedException(); }
        public override void SetLength(long value) { throw new NotSupportedException(); }

        protected override void Dispose(bool disposing)
        {
            if (disposing && !_closed)
            {
                _closed = true;
                CloseEntry();
                _archive.Dispose(); // writes the central directory
            }
            base.Dispose(disposing);
        }
    }
}
