using System;
using System.Collections.Generic;
using System.Text;
using Yarhl.FileFormat;
using Yarhl.IO;
using Yarhl.Media.Text;

namespace Erebonia.Converters
{
    public class Po2BinaryEasy : IConverter<BinaryFormat, Po>
    {
        public Po PoPassed { get; set; }
        public Po Convert(BinaryFormat source)
        {
            return PoPassed;
        }
    }
}
