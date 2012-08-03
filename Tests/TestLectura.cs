// -----------------------------------------------------------------------
// <copyright file="TestLectura.cs" company="Microsoft">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

namespace Tests
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using NUnit.Framework;
    using NPOI.HSSF.UserModel;
    using System.IO;
    
    [TestFixture]
    public class TestLectura
    {
        [Test()]
        public void CargarArchivo()
        {
            System.Diagnostics.Debugger.Launch();
            var flujo=File.OpenRead("C:\\output.xls");
            HSSFWorkbook libro = new HSSFWorkbook(flujo, false);
            var hoja = libro.GetSheetAt(0);
            var fila = hoja.GetRow(1);
            var celda=fila.GetCell(2);
        }
    }
}
