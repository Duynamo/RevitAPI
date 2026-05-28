
#region Namespaces

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Application = Autodesk.Revit.ApplicationServices.Application;
using Form      = System.Windows.Forms.Form;

#endregion

namespace AlphaBIM
{
    // ══════════════════════════════════════════════════════════════════
    //  BỘ LỌC: Chỉ cho phép chọn MEPCurve (Pipe / Duct)
    // ══════════════════════════════════════════════════════════════════
    public class MEPCurveFilter : ISelectionFilter
    {
        public bool AllowElement(Element e)  => e is MEPCurve;
        public bool AllowReference(Reference r, XYZ p) => false;
    }

    // ══════════════════════════════════════════════════════════════════
    //  UI: Cửa sổ xoay Elbow (WinForms, TopMost)
    // ══════════════════════════════════════════════════════════════════
    public class ElbowRotationForm : Form
    {
        // ── Controls ──────────────────────────────────────────────────
        private Label       _lblAngle;
        private TextBox     _txtAngle;
        private Button      _btnApply;
        private Button      _btnClose;
        private Label       _lblTotal;

        // ── State ─────────────────────────────────────────────────────
        public  double ApplyDelta   { get; private set; }   // góc delta vừa apply (radian)
        public  bool   ShouldApply  { get; private set; }   // cờ khi Apply được nhấn
        private double _totalDeg    = 0;                    // tổng góc đã xoay (độ)

        public ElbowRotationForm()
        {
            BuildUI();
        }

        private void BuildUI()
        {
            // ── Form ─────────────────────────────────────────────────
            Text            = "AlphaBIM – Xoay Co (Elbow)";
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            TopMost         = true;
            StartPosition   = FormStartPosition.Manual;
            Location        = new System.Drawing.Point(100, 200);
            Width           = 280;
            Height          = 175;
            ShowInTaskbar   = false;

            // ── Label Góc ────────────────────────────────────────────
            _lblAngle = new Label
            {
                Text     = "Góc quay thêm (độ):",
                Left     = 12, Top  = 14,
                Width    = 180, Height = 20
            };

            // ── TextBox nhập góc ─────────────────────────────────────
            _txtAngle = new TextBox
            {
                Left  = 12, Top    = 36,
                Width = 240, Height = 24,
                Text  = "90"
            };

            // ── Label tổng ───────────────────────────────────────────
            _lblTotal = new Label
            {
                Text     = "Tổng đã xoay: 0°",
                Left     = 12, Top  = 66,
                Width    = 240, Height = 20,
                ForeColor = System.Drawing.Color.Gray
            };

            // ── Nút Apply ────────────────────────────────────────────
            _btnApply = new Button
            {
                Text   = "▶  Quay / Áp dụng",
                Left   = 12,  Top    = 92,
                Width  = 115, Height = 30
            };
            _btnApply.Click += BtnApply_Click;

            // ── Nút Close ────────────────────────────────────────────
            _btnClose = new Button
            {
                Text   = "✕  Đóng",
                Left   = 137, Top    = 92,
                Width  = 115, Height = 30,
                DialogResult = DialogResult.Cancel
            };
            _btnClose.Click += (s, e) => Close();

            Controls.AddRange(new Control[]
                { _lblAngle, _txtAngle, _lblTotal, _btnApply, _btnClose });

            AcceptButton = _btnApply;
            CancelButton = _btnClose;
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            if (!double.TryParse(_txtAngle.Text.Replace(',', '.'),
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out double deg))
            {
                MessageBox.Show("Vui lòng nhập số hợp lệ.", "Lỗi nhập liệu",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ApplyDelta  = deg * Math.PI / 180.0;   // đổi sang Radian
            ShouldApply = true;
            _totalDeg  += deg;
            _lblTotal.Text = $"Tổng đã xoay: {_totalDeg:F1}°";

            // Reset cờ sau một tick để vòng lặp bên ngoài bắt được
            System.Threading.Tasks.Task.Delay(50)
                  .ContinueWith(_ => ShouldApply = false);
        }

        // Cho phép vòng lặp ngoài đợi người dùng tương tác
        public void ProcessEventsOnce()
        {
            System.Windows.Forms.Application.DoEvents();
        }
    }

    // ══════════════════════════════════════════════════════════════════
    //  LỆNH CHÍNH: Create Elbow
    // ══════════════════════════════════════════════════════════════════
    [Transaction(TransactionMode.Manual)]
    public class CreateElbowCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData,
            ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument    uidoc = uiapp.ActiveUIDocument;
            Application   app   = uiapp.Application;
            Document      doc   = uidoc.Document;

            // Khi chạy bằng Add-in Manager thì comment 2 dòng bên dưới để tránh lỗi
            string dllFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            AssemblyLoader.LoadAllRibbonAssemblies(dllFolder);

            try
            {
                // ══════════════════════════════════════════════════════
                //  BƯỚC 1: Chọn vị trí đầu ống
                // ══════════════════════════════════════════════════════

                // 1a. Người dùng click chọn một điểm trong không gian
                XYZ pickedPoint = uidoc.Selection.PickPoint(
                    ObjectSnapTypes.Endpoints | ObjectSnapTypes.Midpoints |
                    ObjectSnapTypes.Nearest,
                    "BƯỚC 1 – Click chọn vị trí gần đầu Pipe / Duct cần đặt Co:");

                // 1b. Người dùng chọn đoạn ống (chỉ MEPCurve)
                Reference pickRef = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    new MEPCurveFilter(),
                    "BƯỚC 1 – Chọn Pipe / Duct chứa connector muốn đặt Co:");

                MEPCurve mepCurve = doc.GetElement(pickRef.ElementId) as MEPCurve;
                if (mepCurve == null)
                {
                    message = "Phần tử được chọn không hợp lệ.";
                    return Result.Failed;
                }

                // ══════════════════════════════════════════════════════
                //  BƯỚC 2: Tìm Connector gần điểm được chọn nhất
                // ══════════════════════════════════════════════════════
                Connector nearestConn = FindNearestConnector(mepCurve, pickedPoint);

                if (nearestConn == null)
                {
                    TaskDialog.Show("Lỗi",
                        "Không tìm thấy Connector trên phần tử được chọn.");
                    return Result.Failed;
                }

                if (nearestConn.IsConnected)
                {
                    TaskDialog.Show("Connector đã kết nối",
                        "Connector gần nhất đã được kết nối với phần tử khác.\n" +
                        "Vui lòng chọn điểm gần đầu hở của ống.");
                    return Result.Failed;
                }

                // ══════════════════════════════════════════════════════
                //  BƯỚC 3: Tạo Elbow Fitting
                // ══════════════════════════════════════════════════════

                // 3a. Lấy FamilySymbol Elbow từ RoutingPreferenceManager
                FamilySymbol elbowSymbol = GetElbowSymbol(doc, mepCurve);

                if (elbowSymbol == null)
                {
                    TaskDialog.Show("Thiếu Family",
                        "Không tìm thấy Family Elbow trong Routing Preference của ống.\n" +
                        "Vui lòng cấu hình Routing Preference cho loại ống này.");
                    return Result.Failed;
                }

                // 3b. Tạo ống stub ngắn vuông góc để làm connector thứ 2
                //     (NewElbowFitting yêu cầu 2 connector hở từ 2 MEPCurve)
                FamilyInstance elbowFitting = null;
                ElementId      stubId       = ElementId.InvalidElementId;

                using (Transaction txCreate = new Transaction(doc, "AlphaBIM – Create Elbow"))
                {
                    txCreate.Start();
                    try
                    {
                        // 3b-i. Kích hoạt FamilySymbol nếu chưa active
                        if (!elbowSymbol.IsActive)
                            elbowSymbol.Activate();

                        // 3b-ii. Tạo ống stub ngắn vuông góc với connector
                        Connector stubConn = CreateStubSegment(
                            doc, mepCurve, nearestConn, out stubId);

                        if (stubConn == null)
                            throw new InvalidOperationException(
                                "Không thể tạo đoạn ống tham chiếu để đặt Co.");

                        // 3b-iii. Kết nối 2 connector
                        if (!nearestConn.IsConnectedTo(stubConn))
                            nearestConn.ConnectTo(stubConn);

                        // 3b-iv. Tạo Elbow Fitting
                        elbowFitting = doc.Create.NewElbowFitting(nearestConn, stubConn);

                        if (elbowFitting == null)
                            throw new InvalidOperationException(
                                "Revit không thể tạo Elbow Fitting tại vị trí này.");

                        txCreate.Commit();
                    }
                    catch (Exception exCreate)
                    {
                        txCreate.RollBack();
                        TaskDialog.Show("Lỗi tạo Co",
                            $"Chi tiết: {exCreate.Message}\n\n" +
                            "Gợi ý: Kiểm tra Routing Preference và kích thước ống.");
                        return Result.Failed;
                    }
                }

                // ══════════════════════════════════════════════════════
                //  BƯỚC 4: Vòng lặp UI Xoay Elbow (TopMost WinForms)
                // ══════════════════════════════════════════════════════

                // Trục xoay: đường thẳng qua tâm connector, hướng BasisZ
                XYZ  axisOrigin    = nearestConn.Origin;
                XYZ  axisDirection = nearestConn.CoordinateSystem.BasisZ;
                Line rotationAxis  = Line.CreateUnbound(axisOrigin, axisDirection);

                using (ElbowRotationForm rotForm = new ElbowRotationForm())
                {
                    rotForm.Show();   // non-modal để Revit vẫn respond

                    while (rotForm.Visible)
                    {
                        rotForm.ProcessEventsOnce();

                        if (rotForm.ShouldApply)
                        {
                            double delta = rotForm.ApplyDelta;

                            using (Transaction txRot =
                                new Transaction(doc, "AlphaBIM – Rotate Elbow"))
                            {
                                txRot.Start();
                                try
                                {
                                    ElementTransformUtils.RotateElement(
                                        doc, elbowFitting.Id, rotationAxis, delta);
                                    txRot.Commit();
                                }
                                catch
                                {
                                    txRot.RollBack();
                                }
                            }

                            uidoc.RefreshActiveView();
                        }

                        System.Threading.Thread.Sleep(20); // giảm tải CPU
                    }
                }

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;   // ESC — không báo lỗi
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        // ──────────────────────────────────────────────────────────────
        //  HÀM PHỤ 1: Tìm Connector gần pickedPoint nhất
        // ──────────────────────────────────────────────────────────────
        private Connector FindNearestConnector(MEPCurve mepCurve, XYZ pickedPoint)
        {
            Connector nearest  = null;
            double    minDist  = double.MaxValue;

            foreach (Connector c in mepCurve.ConnectorManager.Connectors)
            {
                double dist = c.Origin.DistanceTo(pickedPoint);
                if (dist < minDist)
                {
                    minDist = dist;
                    nearest = c;
                }
            }

            return nearest;
        }

        // ──────────────────────────────────────────────────────────────
        //  HÀM PHỤ 2: Lấy FamilySymbol Elbow từ RoutingPreferenceManager
        // ──────────────────────────────────────────────────────────────
        private FamilySymbol GetElbowSymbol(Document doc, MEPCurve mepCurve)
        {
            try
            {
                ElementType elemType = doc.GetElement(mepCurve.GetTypeId()) as ElementType;

                RoutingPreferenceManager rpm = null;

                if (elemType is PipeType pipeType)
                    rpm = pipeType.RoutingPreferenceManager;
                else if (elemType is DuctType ductType)
                    rpm = ductType.RoutingPreferenceManager;

                if (rpm == null) return null;

                int ruleCount = rpm.GetNumberOfRules(RoutingPreferenceRuleGroupType.Elbows);
                if (ruleCount == 0) return null;

                // Lấy rule đầu tiên (mặc định)
                RoutingPreferenceRule rule =
                    rpm.GetRule(RoutingPreferenceRuleGroupType.Elbows, 0);

                return doc.GetElement(rule.MEPPartId) as FamilySymbol;
            }
            catch
            {
                return null;
            }
        }

        // ──────────────────────────────────────────────────────────────
        //  HÀM PHỤ 3: Tạo đoạn ống stub ngắn vuông góc với nearestConn
        //  Trả về connector hở của stub gần với nearestConn
        // ──────────────────────────────────────────────────────────────
        private Connector CreateStubSegment(Document doc,
            MEPCurve original, Connector nearestConn, out ElementId stubId)
        {
            stubId = ElementId.InvalidElementId;

            // Hướng dọc theo ống (BasisZ của connector = chiều ra khỏi ống)
            XYZ flowDir = nearestConn.CoordinateSystem.BasisZ;

            // Lấy vector vuông góc với flowDir (dùng BasisX của connector)
            XYZ perpDir = nearestConn.CoordinateSystem.BasisX;
            if (perpDir.IsZeroLength()) perpDir = XYZ.BasisX;

            // Điểm đầu stub = tâm connector + 1 chút theo perpDir
            const double stubLength = 1.0; // 1 foot (~30 cm) — đủ để elbow fit
            XYZ startPt = nearestConn.Origin;
            XYZ endPt   = startPt + perpDir * stubLength;

            // Lấy level của ống gốc
            Level level = doc.GetElement(original.ReferenceLevel?.Id
                          ?? Level.Create(doc, 0).Id) as Level;

            // Tạo Pipe hoặc Duct stub cùng loại với ống gốc
            MEPCurve stub = null;

            if (original is Pipe)
            {
                stub = Pipe.Create(doc,
                    original.GetTypeId(),
                    original.ReferenceLevel.Id,
                    nearestConn,
                    endPt);
            }
            else if (original is Duct)
            {
                stub = Duct.Create(doc,
                    original.GetTypeId(),
                    original.ReferenceLevel.Id,
                    nearestConn,
                    endPt);
            }

            if (stub == null) return null;

            stubId = stub.Id;

            // Trả về connector của stub gần với nearestConn (đầu kết nối)
            Connector stubNear   = null;
            double    minDist    = double.MaxValue;

            foreach (Connector c in stub.ConnectorManager.Connectors)
            {
                double d = c.Origin.DistanceTo(nearestConn.Origin);
                if (d < minDist)
                {
                    minDist  = d;
                    stubNear = c;
                }
            }

            return stubNear;
        }
    }
}
