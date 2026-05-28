using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;

namespace AlphaBIM
{
    public partial class SplitPipeWindow : Window
    {
        private readonly UIDocument    _uidoc;
        private readonly Document      _doc;
        private readonly SplitPipeHandler _handler;
        private readonly ExternalEvent _exEvent;
        private Pipe _selPipe;

        public SplitPipeWindow(UIDocument uidoc, Document doc,
                               SplitPipeHandler handler, ExternalEvent exEvent)
        {
            _uidoc   = uidoc;
            _doc     = doc;
            _handler = handler;
            _exEvent = exEvent;
            InitializeComponent();
        }

        // ─── Pick Pipe ───────────────────────────────────────────────────────
        // KEY FIX: ẩn window TRƯỚC khi PickObject → PickObject nhận focus từ Revit
        private void BtnPickPipe_Click(object sender, RoutedEventArgs e)
        {
            Hide();   // ẩn WPF window → trả control về Revit
            try
            {
                var filter = new PipeSelectionFilter();
                Reference refObj = _uidoc.Selection.PickObject(
                    Autodesk.Revit.UI.Selection.ObjectType.Element,
                    filter,
                    "Chọn một Pipe cần cắt");

                _selPipe = _doc.GetElement(refObj.ElementId) as Pipe;

                if (_selPipe != null)
                {
                    txb_pipeInfo.Text   = $"✔  Đã chọn Pipe  ID: {_selPipe.Id}";
                    btn_pickPipe.Content = $"✔   PIPE  ID: {_selPipe.Id}";
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // user nhấn Esc – không làm gì
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Lỗi chọn Pipe",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            finally
            {
                Show();        // hiện lại window
                Activate();    // đưa window lên foreground
            }
        }

        // ─── Split ───────────────────────────────────────────────────────────
        private void BtnSplit_Click(object sender, RoutedEventArgs e)
        {
            if (_selPipe == null)
            {
                MessageBox.Show("Vui lòng chọn một Pipe trước!",
                    "Thiếu đầu vào", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!double.TryParse(txb_Length.Text.Trim(), out double mm) || mm <= 0)
            {
                MessageBox.Show("Nhập chiều dài hợp lệ (mm).",
                    "Đầu vào không hợp lệ", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!int.TryParse(txb_K.Text.Trim(), out int k) || k <= 0)
            {
                MessageBox.Show("Nhập số đoạn K hợp lệ.",
                    "Đầu vào không hợp lệ", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Truyền tham số vào handler
            _handler.SelectedPipe     = _selPipe;
            _handler.SegmentLengthFt  = mm / 304.8;
            _handler.SplitK           = k;
            _handler.SortByX          = rbt_pX.IsChecked == true;
            _handler.SortByY          = rbt_pY.IsChecked == true;
            _handler.SortByZ          = rbt_pZ.IsChecked == true;
            _handler.SortByMin        = rbt_sortByMin.IsChecked == true;
            _handler.CreateFlange     = chk_createFlange.IsChecked == true;

            // ExternalEvent thực thi trên Revit main thread
            _exEvent.Raise();
            Close();
        }

        // ─── Cancel ──────────────────────────────────────────────────────────
        private void BtnCancel_Click(object sender, RoutedEventArgs e) => Close();
    }
}
