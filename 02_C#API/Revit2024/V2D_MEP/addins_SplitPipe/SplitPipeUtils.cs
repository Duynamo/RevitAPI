using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace AlphaBIM
{
    // ─── Selection Filter: chỉ chọn Pipe ───────────────────────────────────
    public class PipeSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem.Category != null &&
                   elem.Category.Id.Value == (long)BuiltInCategory.OST_PipeCurves;
        }

        public bool AllowReference(Reference reference, XYZ position) => false;
    }

    // ─── Tiện ích cắt pipe ─────────────────────────────────────────────────
    public static class SplitPipeUtils
    {
        /// <summary>
        /// Cho người dùng pick một Pipe từ canvas Revit.
        /// </summary>
        public static Pipe PickPipe(UIDocument uidoc)
        {
            Reference refObj = uidoc.Selection.PickObject(
                ObjectType.Element,
                new PipeSelectionFilter(),
                "Chọn một Pipe cần cắt");

            return uidoc.Document.GetElement(refObj.ElementId) as Pipe;
        }

        /// <summary>
        /// Chia đoạn thẳng thành các điểm cách đều nhau theo <paramref name="segmentLength"/> (feet).
        /// Trả về danh sách điểm bắt đầu từ startPoint (bao gồm chính nó),
        /// mỗi bước tăng thêm segmentLength.
        /// </summary>
        public static List<XYZ> DivideLineSegment(
            XYZ startPoint, XYZ endPoint, double segmentLength)
        {
            var points = new List<XYZ>();
            double totalLength = startPoint.DistanceTo(endPoint);
            XYZ direction = (endPoint - startPoint).Normalize();

            XYZ current = startPoint;
            points.Add(current);

            while (current.DistanceTo(startPoint) + segmentLength <= totalLength)
            {
                current = current + direction * segmentLength;
                points.Add(current);
            }
            return points;
        }

        /// <summary>
        /// Kiểm tra điểm có nằm trên curve không (với ngưỡng 1e-6).
        /// </summary>
        public static bool IsPointOnCurve(Curve curve, XYZ point)
        {
            IntersectionResult result = curve.Project(point);
            return result != null && result.Distance < 1e-6;
        }

        /// <summary>
        /// Cắt pipe tại nhiều điểm theo thứ tự.
        /// Sử dụng PlumbingUtils.BreakCurve cho mỗi điểm.
        /// </summary>
        public static List<Pipe> SplitPipeAtPoints(
            Document doc, Pipe pipe, List<XYZ> points)
        {
            var newPipes = new List<Pipe>();
            Pipe currentPipe = pipe;
            Pipe originalPipe = pipe;

            foreach (XYZ point in points)
            {
                LocationCurve pipeLocation = currentPipe.Location as LocationCurve;
                if (pipeLocation == null) continue;

                Curve pipeCurve = pipeLocation.Curve;
                if (pipeCurve == null) continue;

                ElementId newPipeId;
                if (IsPointOnCurve(pipeCurve, point))
                {
                    newPipeId = PlumbingUtils.BreakCurve(doc, currentPipe.Id, point);
                }
                else
                {
                    // Điểm không nằm trên pipe hiện tại → thử với pipe gốc
                    newPipeId = PlumbingUtils.BreakCurve(doc, originalPipe.Id, point);
                }

                if (newPipeId != null && newPipeId != ElementId.InvalidElementId)
                {
                    Pipe newPipe = doc.GetElement(newPipeId) as Pipe;
                    if (newPipe != null)
                    {
                        newPipes.Add(newPipe);
                        currentPipe = newPipe;
                    }
                }
            }

            return newPipes;
        }

        /// <summary>
        /// Tìm đúng 2 connector hở gần điểm splitPt (sau khi BreakCurve).
        /// tolerance tính theo feet (~15 mm = 0.05 ft).
        /// </summary>
        public static List<Connector> GetOpenConnectorsNearPoint(
            Document doc, XYZ splitPt, double tolerance = 0.05)
        {
            var openConns = new List<Connector>();

            var allPipes = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_PipeCurves)
                .OfClass(typeof(Pipe))
                .Cast<Pipe>();

            foreach (Pipe p in allPipes)
            {
                try
                {
                    foreach (Connector conn in p.ConnectorManager.Connectors)
                    {
                        if (conn.IsConnected) continue;
                        if (conn.Origin.DistanceTo(splitPt) <= tolerance)
                            openConns.Add(conn);
                    }
                }
                catch { /* bỏ qua pipe lỗi */ }
            }

            return openConns;
        }

        /// <summary>
        /// Tạo union fitting (mặt bích) tại mỗi split point bằng NewUnionFitting.
        /// Pipe type cần có Union fitting trong Routing Preferences.
        /// </summary>
        public static List<FamilyInstance> CreateFlangesAtSplitPoints(
            Document doc, List<XYZ> splitPoints)
        {
            var flanges = new List<FamilyInstance>();

            foreach (XYZ splitPt in splitPoints)
            {
                try
                {
                    List<Connector> openConns = GetOpenConnectorsNearPoint(doc, splitPt);

                    if (openConns.Count >= 2)
                    {
                        FamilyInstance union = doc.Create.NewUnionFitting(
                            openConns[0], openConns[1]);
                        if (union != null) flanges.Add(union);
                    }
                    else
                    {
                        // Fallback: tolerance lớn hơn (~30 mm)
                        List<Connector> openConns2 =
                            GetOpenConnectorsNearPoint(doc, splitPt, tolerance: 0.1);
                        if (openConns2.Count >= 2)
                        {
                            FamilyInstance union = doc.Create.NewUnionFitting(
                                openConns2[0], openConns2[1]);
                            if (union != null) flanges.Add(union);
                        }
                    }
                }
                catch { /* bỏ qua nếu tạo fitting lỗi */ }
            }

            return flanges;
        }
    }
}
