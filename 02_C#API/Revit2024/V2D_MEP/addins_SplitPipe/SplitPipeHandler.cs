using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;

namespace AlphaBIM
{
    /// <summary>
    /// ExternalEvent handler – thực thi trên Revit main thread.
    /// Nhận tham số từ WPF window, thực hiện split + tạo flange.
    /// </summary>
    public class SplitPipeHandler : IExternalEventHandler
    {
        // ── Tham số nhận từ UI ───────────────────────────────────────────
        public Pipe   SelectedPipe    { get; set; }
        public double SegmentLengthFt { get; set; }
        public int    SplitK          { get; set; }
        public bool   SortByX         { get; set; }
        public bool   SortByY         { get; set; }
        public bool   SortByZ         { get; set; }
        public bool   SortByMin       { get; set; }
        public bool   CreateFlange    { get; set; }

        // ── Kết quả trả về (tuỳ chọn) ───────────────────────────────────
        public bool   Success         { get; private set; }
        public string ResultMessage   { get; private set; }

        public string GetName() => "SplitPipe";

        public void Execute(UIApplication app)
        {
            Success       = false;
            ResultMessage = string.Empty;

            var doc = app.ActiveUIDocument.Document;

            try
            {
                if (SelectedPipe == null)
                {
                    ResultMessage = "Không có pipe nào được chọn.";
                    return;
                }

                // ── Lấy connector origins & sort ──────────────────────────
                var origins = new List<XYZ>();
                foreach (Connector c in SelectedPipe.ConnectorManager.Connectors)
                    origins.Add(c.Origin);

                List<XYZ> sorted;
                if (SortByX)
                    sorted = SortByMin
                        ? origins.OrderBy(p => p.X).ToList()
                        : origins.OrderByDescending(p => p.X).ToList();
                else if (SortByY)
                    sorted = SortByMin
                        ? origins.OrderBy(p => p.Y).ToList()
                        : origins.OrderByDescending(p => p.Y).ToList();
                else
                    sorted = SortByMin
                        ? origins.OrderBy(p => p.Z).ToList()
                        : origins.OrderByDescending(p => p.Z).ToList();

                XYZ startPt = sorted[0];
                XYZ endPt   = sorted[1];

                // ── Tính điểm cắt ─────────────────────────────────────────
                List<XYZ> allPts = SplitPipeUtils.DivideLineSegment(
                    startPt, endPt, SegmentLengthFt);

                // Bỏ điểm đầu (chính startPt), lấy tối đa SplitK điểm
                List<XYZ> splitPoints = allPts.Count > 1
                    ? allPts.Skip(1).Take(SplitK).ToList()
                    : new List<XYZ>();

                if (splitPoints.Count == 0)
                {
                    ResultMessage = "Không có điểm cắt nào.";
                    return;
                }

                // ── Transaction ───────────────────────────────────────────
                using (var tx = new Transaction(doc, "Split Pipe"))
                {
                    tx.Start();

                    SplitPipeUtils.SplitPipeAtPoints(doc, SelectedPipe, splitPoints);

                    if (CreateFlange)
                        SplitPipeUtils.CreateFlangesAtSplitPoints(doc, splitPoints);

                    tx.Commit();
                }

                Success = true;
            }
            catch (Exception ex)
            {
                ResultMessage = ex.Message;
                // Hiện MessageBox trên Revit dispatcher
                System.Windows.MessageBox.Show(
                    $"Lỗi khi cắt pipe:\n{ex.Message}",
                    "Lỗi",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }
    }
}
