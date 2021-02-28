using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;
using Yarhl.FileFormat;

namespace Erebonia.Formats
{
    class XLS : IFormat
    {
        public XLS()
        {
            Entries = new Collection<Entry>();
        }

        public class Entry
        {
            public uint ID { get; set; }
            public uint Row { get; set; }
            public uint Column { get; set; }
            public uint[] Cell {
                get => new uint[] { Row, Column };
            }
            public string Text { get; set; }
        }

        public Collection<Entry> Entries { get; }
    }
}
