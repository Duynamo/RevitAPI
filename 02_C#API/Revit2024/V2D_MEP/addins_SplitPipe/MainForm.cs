using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
// Alias tránh xung đột tên với Autodesk.Revit.DB
using WinForms  = System.Windows.Forms;
using Drawing   = System.Drawing;

namespace AlphaBIM
{
    /// <summary>
    /// Form WinForms chính để điều khiển thao tác Split Pipe.
    /// </summary>
    public class MainForm : WinForms.Form
    {
        // ── Revit context ────────────────────────────────────────────────
        private readonly UIDocument _uidoc;
        private readonly Document   _doc;
        private Pipe                _selPipe;

        // ── Controls ─────────────────────────────────────────────────────
        private WinForms.Button      _btt_pickPipe;
        private WinForms.Button      _btt_SPLIT;
        private WinForms.Button      _btt_CANCEL;
        private WinForms.Label       _lb_FVC;

        private WinForms.GroupBox    _grb_inputData;
        private WinForms.Label       _lb_Length;
        private WinForms.TextBox     _txb_Length;
        private WinForms.Label       _lb_splitNumber;
        private WinForms.TextBox     _txb_K;

        private WinForms.GroupBox    _grb_sortConn;
        private WinForms.RadioButton _rbt_pX;
        private WinForms.RadioButton _rbt_pY;
        private WinForms.RadioButton _rbt_pZ;

        private WinForms.GroupBox    _grb_MinMax;
        private WinForms.RadioButton _rbt_sortByMin;
        private WinForms.RadioButton _rbt_sortByMax;

        private WinForms.CheckBox    _chk_createFlange;

        // ── Constructor ──────────────────────────────────────────────────
        public MainForm(UIDocument uidoc, Document doc)
        {
            _uidoc = uidoc;
            _doc   = doc;
            InitializeComponent();
        }

        // ══════════════════════════════════════════════════════════════════
        // InitializeComponent
        // ══════════════════════════════════════════════════════════════════
        private void InitializeComponent()
        {
            // ─── tính kích thước form theo màn hình ─────────────────────
            Drawing.Rectangle workArea = WinForms.Screen.PrimaryScreen.WorkingArea;
            int sw = workArea.Width  / 4;   // screen_width  ~  1/4 màn hình
            int sh = workArea.Height / 2;   // screen_height ~  1/2 màn hình

            // ── khởi tạo controls ─────────────────────────────────────
            _btt_pickPipe     = new WinForms.Button();
            _grb_inputData    = new WinForms.GroupBox();
            _lb_Length        = new WinForms.Label();
            _txb_Length       = new WinForms.TextBox();
            _lb_splitNumber   = new WinForms.Label();
            _txb_K            = new WinForms.TextBox();
            _grb_sortConn     = new WinForms.GroupBox();
            _rbt_pX           = new WinForms.RadioButton();
            _rbt_pY           = new WinForms.RadioButton();
            _rbt_pZ           = new WinForms.RadioButton();
            _grb_MinMax       = new WinForms.GroupBox();
            _rbt_sortByMin    = new WinForms.RadioButton();
            _rbt_sortByMax    = new WinForms.RadioButton();
            _chk_createFlange = new WinForms.CheckBox();
            _btt_SPLIT        = new WinForms.Button();
            _btt_CANCEL       = new WinForms.Button();
            _lb_FVC           = new WinForms.Label();

            _grb_sortConn.SuspendLayout();
            _grb_inputData.SuspendLayout();
            _grb_MinMax.SuspendLayout();
            SuspendLayout();

            // ── Form ────────────────────────────────────────────────────
            ClientSize      = new Drawing.Size(sw, sh);
            StartPosition   = WinForms.FormStartPosition.CenterScreen;
            FormBorderStyle = WinForms.FormBorderStyle.Fixed3D;
            Name            = "MainForm";
            Text            = "Split Pipe";
            TopMost         = true;
            Load           += MainForm_Load;

            // ── btt_pickPipe ─────────────────────────────────────────────
            _btt_pickPipe.Font              = new Drawing.Font("Meiryo UI", 10.2f, Drawing.FontStyle.Bold, Drawing.GraphicsUnit.Point, 128);
            _btt_pickPipe.ForeColor         = Drawing.Color.Red;
            _btt_pickPipe.Location          = new Drawing.Point((int)(sw * 0.3), (int)(sh * 0.05));
            _btt_pickPipe.Name              = "btt_pickPipe";
            _btt_pickPipe.Size              = new Drawing.Size((int)(sw * 0.35), (int)(sh * 0.1));
            _btt_pickPipe.Text              = "PICK PIPE";
            _btt_pickPipe.UseVisualStyleBackColor = true;
            _btt_pickPipe.Click            += BttPickPipe_Click;

            // ── grb_inputData ────────────────────────────────────────────
            _grb_inputData.Font             = new Drawing.Font("Meiryo UI", 9f, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, 128);
            _grb_inputData.Location         = new Drawing.Point((int)(sw * 0.03), (int)(sh * 0.2));
            _grb_inputData.Name             = "grb_inputData";
            _grb_inputData.Size             = new Drawing.Size((int)(sw * 0.94), (int)(sh * 0.25));
            _grb_inputData.Text             = "Input";
            _grb_inputData.TabStop          = false;

            // ── lb_Length ────────────────────────────────────────────────
            _lb_Length.Font                 = new Drawing.Font("Meiryo UI", 7.8f, Drawing.FontStyle.Bold, Drawing.GraphicsUnit.Point, 128);
            _lb_Length.Location             = new Drawing.Point((int)(sw * 0.05), (int)(sh * 0.05));
            _lb_Length.Name                 = "lb_Length";
            _lb_Length.Size                 = new Drawing.Size((int)(sw * 0.3), (int)(sh * 0.05));
            _lb_Length.Text                 = "Length (mm):";
            _lb_Length.TextAlign            = Drawing.ContentAlignment.MiddleCenter;

            // ── txb_Length ───────────────────────────────────────────────
            _txb_Length.Font                = new Drawing.Font("Meiryo UI", 7.2f, Drawing.FontStyle.Bold, Drawing.GraphicsUnit.Point, 128);
            _txb_Length.Location            = new Drawing.Point((int)(sw * 0.05), (int)(sh * 0.1));
            _txb_Length.Name                = "txb_Length";
            _txb_Length.Size               = new Drawing.Size((int)(sw * 0.35), (int)(sh * 0.07));
            _txb_Length.Text                = "3000";

            // ── lb_splitNumber ───────────────────────────────────────────
            _lb_splitNumber.Font            = new Drawing.Font("Meiryo UI", 7.8f, Drawing.FontStyle.Bold, Drawing.GraphicsUnit.Point, 128);
            _lb_splitNumber.Location        = new Drawing.Point((int)(sw * 0.55), (int)(sh * 0.05));
            _lb_splitNumber.Name            = "lb_splitNumber";
            _lb_splitNumber.Size            = new Drawing.Size((int)(sw * 0.15), (int)(sh * 0.05));
            _lb_splitNumber.Text            = "K:";
            _lb_splitNumber.TextAlign       = Drawing.ContentAlignment.MiddleCenter;

            // ── txb_K ────────────────────────────────────────────────────
            _txb_K.Font                     = new Drawing.Font("Meiryo UI", 7.2f, Drawing.FontStyle.Bold, Drawing.GraphicsUnit.Point, 128);
            _txb_K.Location                 = new Drawing.Point((int)(sw * 0.55), (int)(sh * 0.1));
            _txb_K.Name                     = "txb_K";
            _txb_K.Size                     = new Drawing.Size((int)(sw * 0.35), (int)(sh * 0.07));
            _txb_K.Text                     = "1000";
            _txb_K.TextChanged             += TxbK_TextChanged;

            // ── add controls vào grb_inputData ──────────────────────────
            _grb_inputData.Controls.Add(_lb_Length);
            _grb_inputData.Controls.Add(_txb_Length);
            _grb_inputData.Controls.Add(_lb_splitNumber);
            _grb_inputData.Controls.Add(_txb_K);

            // ── grb_sortConn ─────────────────────────────────────────────
            _grb_sortConn.Font              = new Drawing.Font("Meiryo UI", 9f, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, 128);
            _grb_sortConn.Location          = new Drawing.Point((int)(sw * 0.03), (int)(sh * 0.5));
            _grb_sortConn.Name              = "grb_sortConn";
            _grb_sortConn.Size              = new Drawing.Size((int)(sw * 0.45), (int)(sh * 0.3));
            _grb_sortConn.Text              = "Sort Connector by";
            _grb_sortConn.TabStop           = false;

            // ── rbt_pX ───────────────────────────────────────────────────
            _rbt_pX.Font                    = new Drawing.Font("Meiryo UI", 7.8f, Drawing.FontStyle.Bold, Drawing.GraphicsUnit.Point, 128);
            _rbt_pX.ForeColor               = Drawing.Color.Red;
            _rbt_pX.Location                = new Drawing.Point((int)(sw * 0.05), (int)(sh * 0.05));
            _rbt_pX.Name                    = "rbt_pX";
            _rbt_pX.Size                    = new Drawing.Size((int)(sw * 0.35), (int)(sh * 0.07));
            _rbt_pX.Text                    = "pX";
            _rbt_pX.Checked                 = true;

            // ── rbt_pY ───────────────────────────────────────────────────
            _rbt_pY.Font                    = new Drawing.Font("Meiryo UI", 7.8f, Drawing.FontStyle.Bold, Drawing.GraphicsUnit.Point, 128);
            _rbt_pY.ForeColor               = Drawing.Color.Red;
            _rbt_pY.Location                = new Drawing.Point((int)(sw * 0.05), (int)(sh * 0.12));
            _rbt_pY.Name                    = "rbt_pY";
            _rbt_pY.Size                    = new Drawing.Size((int)(sw * 0.35), (int)(sh * 0.07));
            _rbt_pY.Text                    = "pY";

            // ── rbt_pZ ───────────────────────────────────────────────────
            _rbt_pZ.Font                    = new Drawing.Font("Meiryo UI", 7.8f, Drawing.FontStyle.Bold, Drawing.GraphicsUnit.Point, 128);
            _rbt_pZ.ForeColor               = Drawing.Color.Red;
            _rbt_pZ.Location                = new Drawing.Point((int)(sw * 0.05), (int)(sh * 0.19));
            _rbt_pZ.Name                    = "rbt_pZ";
            _rbt_pZ.Size                    = new Drawing.Size((int)(sw * 0.35), (int)(sh * 0.07));
            _rbt_pZ.Text                    = "pZ";

            _grb_sortConn.Controls.Add(_rbt_pX);
            _grb_sortConn.Controls.Add(_rbt_pY);
            _grb_sortConn.Controls.Add(_rbt_pZ);

            // ── grb_MinMax ───────────────────────────────────────────────
            _grb_MinMax.Font                = new Drawing.Font("Meiryo UI", 9f, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, 128);
            _grb_MinMax.Location            = new Drawing.Point((int)(sw * 0.52), (int)(sh * 0.5));
            _grb_MinMax.Name                = "grb_MinMax";
            _grb_MinMax.Size                = new Drawing.Size((int)(sw * 0.45), (int)(sh * 0.3));
            _grb_MinMax.Text                = "MinMax";
            _grb_MinMax.TabStop             = false;

            // ── rbt_sortByMin ─────────────────────────────────────────────
            _rbt_sortByMin.Font             = new Drawing.Font("Meiryo UI", 7.8f, Drawing.FontStyle.Bold, Drawing.GraphicsUnit.Point, 128);
            _rbt_sortByMin.ForeColor        = Drawing.Color.Fuchsia;
            _rbt_sortByMin.Location         = new Drawing.Point((int)(sw * 0.05), (int)(sh * 0.05));
            _rbt_sortByMin.Name             = "rbt_sortByMin";
            _rbt_sortByMin.Size             = new Drawing.Size((int)(sw * 0.35), (int)(sh * 0.07));
            _rbt_sortByMin.Text             = "Sort By Min?";
            _rbt_sortByMin.Checked          = true;

            // ── rbt_sortByMax ─────────────────────────────────────────────
            _rbt_sortByMax.Font             = new Drawing.Font("Meiryo UI", 7.8f, Drawing.FontStyle.Bold, Drawing.GraphicsUnit.Point, 128);
            _rbt_sortByMax.ForeColor        = Drawing.Color.Fuchsia;
            _rbt_sortByMax.Location         = new Drawing.Point((int)(sw * 0.05), (int)(sh * 0.12));
            _rbt_sortByMax.Name             = "rbt_sortByMax";
            _rbt_sortByMax.Size             = new Drawing.Size((int)(sw * 0.35), (int)(sh * 0.07));
            _rbt_sortByMax.Text             = "Sort By Max?";

            _grb_MinMax.Controls.Add(_rbt_sortByMin);
            _grb_MinMax.Controls.Add(_rbt_sortByMax);

            // ── chk_createFlange ─────────────────────────────────────────
            _chk_createFlange.Font          = new Drawing.Font("Meiryo UI", 8.0f, Drawing.FontStyle.Bold, Drawing.GraphicsUnit.Point, 128);
            _chk_createFlange.ForeColor     = Drawing.Color.DarkBlue;
            _chk_createFlange.Location      = new Drawing.Point((int)(sw * 0.03), (int)(sh * 0.83));
            _chk_createFlange.Name          = "chk_createFlange";
            _chk_createFlange.Size          = new Drawing.Size((int)(sw * 0.38), (int)(sh * 0.1));
            _chk_createFlange.Text          = "Create Flange";
            _chk_createFlange.UseVisualStyleBackColor = true;
            _chk_createFlange.Checked       = false;

            // ── btt_SPLIT ─────────────────────────────────────────────────
            _btt_SPLIT.Font                 = new Drawing.Font("Meiryo UI", 10.2f, Drawing.FontStyle.Bold, Drawing.GraphicsUnit.Point, 128);
            _btt_SPLIT.ForeColor            = Drawing.Color.Red;
            _btt_SPLIT.Location             = new Drawing.Point((int)(sw * 0.45), (int)(sh * 0.85));
            _btt_SPLIT.Name                 = "btt_SPLIT";
            _btt_SPLIT.Size                 = new Drawing.Size((int)(sw * 0.25), (int)(sh * 0.1));
            _btt_SPLIT.Text                 = "SPLIT";
            _btt_SPLIT.UseVisualStyleBackColor = true;
            _btt_SPLIT.Click               += BttSplit_Click;

            // ── btt_CANCEL ────────────────────────────────────────────────
            _btt_CANCEL.Font                = new Drawing.Font("Meiryo UI", 10.2f, Drawing.FontStyle.Bold, Drawing.GraphicsUnit.Point, 128);
            _btt_CANCEL.ForeColor           = Drawing.Color.Red;
            _btt_CANCEL.Location            = new Drawing.Point((int)(sw * 0.72), (int)(sh * 0.85));
            _btt_CANCEL.Name                = "btt_CANCEL";
            _btt_CANCEL.Size                = new Drawing.Size((int)(sw * 0.25), (int)(sh * 0.1));
            _btt_CANCEL.Text                = "CANCEL";
            _btt_CANCEL.UseVisualStyleBackColor = true;
            _btt_CANCEL.Click              += BttCancel_Click;

            // ── lb_FVC ────────────────────────────────────────────────────
            _lb_FVC.Font                    = new Drawing.Font("Meiryo UI", 4.8f, Drawing.FontStyle.Bold, Drawing.GraphicsUnit.Point, 128);
            _lb_FVC.ForeColor               = Drawing.Color.Blue;
            _lb_FVC.Location                = new Drawing.Point((int)(sw * 0.05), (int)(sh * 0.9));
            _lb_FVC.Name                    = "lb_FVC";
            _lb_FVC.Size                    = new Drawing.Size((int)(sw * 0.15), (int)(sh * 0.05));
            _lb_FVC.Text                    = "@FVC";
            _lb_FVC.TextAlign               = Drawing.ContentAlignment.MiddleCenter;

            // ── Thêm controls vào Form ────────────────────────────────────
            Controls.Add(_btt_pickPipe);
            Controls.Add(_grb_inputData);
            Controls.Add(_grb_sortConn);
            Controls.Add(_grb_MinMax);
            Controls.Add(_chk_createFlange);
            Controls.Add(_btt_SPLIT);
            Controls.Add(_btt_CANCEL);
            Controls.Add(_lb_FVC);

            _grb_sortConn.ResumeLayout(false);
            _grb_inputData.ResumeLayout(false);
            _grb_inputData.PerformLayout();
            _grb_MinMax.ResumeLayout(false);
            ResumeLayout(false);
        }

        // ══════════════════════════════════════════════════════════════════
        // Event Handlers
        // ══════════════════════════════════════════════════════════════════

        private void MainForm_Load(object sender, EventArgs e) { }

        private void TxbK_TextChanged(object sender, EventArgs e) { }

        /// <summary>
        /// Ẩn form, cho phép người dùng pick pipe trong Revit, rồi hiện lại form.
        /// </summary>
        private void BttPickPipe_Click(object sender, EventArgs e)
        {
            try
            {
                Hide();
                _selPipe = SplitPipeUtils.PickPipe(_uidoc);
                Show();
                BringToFront();

                if (_selPipe != null)
                    _btt_pickPipe.Text = $"✔ {_selPipe.Id}";
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                Show();
                BringToFront();
            }
            catch (Exception ex)
            {
                Show();
                BringToFront();
                WinForms.MessageBox.Show(ex.Message, "Lỗi chọn Pipe",
                    WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// Thực hiện cắt pipe và (tuỳ chọn) tạo mặt bích.
        /// </summary>
        private void BttSplit_Click(object sender, EventArgs e)
        {
            if (_selPipe == null)
            {
                WinForms.MessageBox.Show("Vui lòng chọn một Pipe trước!", "Thiếu đầu vào",
                    WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
                return;
            }

            // ── Đọc đầu vào ─────────────────────────────────────────────
            if (!double.TryParse(_txb_Length.Text.Trim(), out double lengthMm) || lengthMm <= 0)
            {
                WinForms.MessageBox.Show("Nhập chiều dài hợp lệ (mm).", "Đầu vào không hợp lệ",
                    WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
                return;
            }

            if (!int.TryParse(_txb_K.Text.Trim(), out int splitK) || splitK <= 0)
            {
                WinForms.MessageBox.Show("Nhập số đoạn K hợp lệ.", "Đầu vào không hợp lệ",
                    WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
                return;
            }

            // Chuyển mm → feet
            double segmentLengthFt = lengthMm / 304.8;

            // ── Lấy connector endpoints và sort ──────────────────────────
            var connectors = new List<Connector>();
            foreach (Connector c in _selPipe.ConnectorManager.Connectors)
                connectors.Add(c);

            List<XYZ> originConns = connectors.Select(c => c.Origin).ToList();

            List<XYZ> sortConns;
            if (_rbt_pX.Checked)
                sortConns = _rbt_sortByMin.Checked
                    ? originConns.OrderBy(p => p.X).ToList()
                    : originConns.OrderByDescending(p => p.X).ToList();
            else if (_rbt_pY.Checked)
                sortConns = _rbt_sortByMin.Checked
                    ? originConns.OrderBy(p => p.Y).ToList()
                    : originConns.OrderByDescending(p => p.Y).ToList();
            else // pZ
                sortConns = _rbt_sortByMin.Checked
                    ? originConns.OrderBy(p => p.Z).ToList()
                    : originConns.OrderByDescending(p => p.Z).ToList();

            XYZ startPt = sortConns[0];
            XYZ endPt   = sortConns[1];

            // ── Tính các điểm cắt ─────────────────────────────────────────
            List<XYZ> allPoints = SplitPipeUtils.DivideLineSegment(startPt, endPt, segmentLengthFt);

            // Lấy splitK điểm (bỏ qua điểm đầu)
            List<XYZ> splitPoints = allPoints.Count > 1
                ? allPoints.Skip(1).Take(splitK).ToList()
                : new List<XYZ>();

            if (splitPoints.Count == 0)
            {
                WinForms.MessageBox.Show(
                    "Không có điểm cắt nào được tạo ra.\nKiểm tra lại chiều dài và giá trị K.",
                    "Thông báo",
                    WinForms.MessageBoxButtons.OK,
                    WinForms.MessageBoxIcon.Information);
                return;
            }

            // ── Thực hiện cắt pipe trong transaction ──────────────────────
            using (Transaction tx = new Transaction(_doc, "Split Pipe"))
            {
                tx.Start();
                try
                {
                    SplitPipeUtils.SplitPipeAtPoints(_doc, _selPipe, splitPoints);

                    // ── Tạo mặt bích nếu checkbox được chọn ──────────────
                    if (_chk_createFlange.Checked)
                    {
                        SplitPipeUtils.CreateFlangesAtSplitPoints(_doc, splitPoints);
                    }

                    tx.Commit();
                }
                catch (Exception ex)
                {
                    tx.RollBack();
                    WinForms.MessageBox.Show($"Lỗi khi cắt pipe:\n{ex.Message}", "Lỗi",
                        WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
                    return;
                }
            }

            // Đóng form sau khi hoàn thành
            Close();
        }

        private void BttCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
