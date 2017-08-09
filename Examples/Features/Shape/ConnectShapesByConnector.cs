using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Shape
{
    public class ConnectShapesByConnector : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Spread.Drawing.IShape ShapeBegin = worksheet.Shapes.AddShape(GrapeCity.Documents.Spread.Drawing.AutoShapeType.Rectangle, 1, 1, 100, 100);
            GrapeCity.Documents.Spread.Drawing.IShape EndBegin = worksheet.Shapes.AddShape(GrapeCity.Documents.Spread.Drawing.AutoShapeType.Rectangle, 200, 200, 100, 100);
            GrapeCity.Documents.Spread.Drawing.IShape ConnectorShape = worksheet.Shapes.AddConnector(GrapeCity.Documents.Spread.Drawing.ConnectorType.Straight, 1, 1, 101, 101);

            //connect shapes by connector shape.
            ConnectorShape.ConnectorFormat.BeginConnect(ShapeBegin, 3);
            ConnectorShape.ConnectorFormat.EndConnect(EndBegin, 0);
        }

        public override bool ShowViewer
        {
            get
            {
                return false;
            }
        }
    }
}
