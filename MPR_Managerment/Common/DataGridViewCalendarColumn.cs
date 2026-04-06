using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MPR_Managerment.Common
{
    using System;
    using System.Windows.Forms;

    public class DataGridViewCalendarColumn : DataGridViewColumn
    {
        public DataGridViewCalendarColumn() : base(new DataGridViewCalendarCell())
        {
        }

        public override DataGridViewCell CellTemplate
        {
            get => base.CellTemplate;
            set
            {
                if (value != null && !value.GetType().IsAssignableFrom(typeof(DataGridViewCalendarCell)))
                    throw new InvalidCastException("Phải là DataGridViewCalendarCell");
                base.CellTemplate = value;
            }
        }
    }

    public class DataGridViewCalendarCell : DataGridViewTextBoxCell
    {
        public DataGridViewCalendarCell() : base()
        {
            // Định dạng mặc định khi hiển thị
            this.Style.Format = "dd-MM-yyyy";
        }

        public override void InitializeEditingControl(int rowIndex, object initialValue, DataGridViewCellStyle dataGridViewCellStyle)
        {
            // Gọi base để khởi tạo control chỉnh sửa
            base.InitializeEditingControl(rowIndex, initialValue, dataGridViewCellStyle);

            // Lấy EditingControl từ DataGridView quản lý Cell này
            // Sử dụng this.DataGridView để truy cập trực tiếp vào Grid chứa Cell
            CalendarEditingControl ctl = DataGridView.EditingControl as CalendarEditingControl;

            if (ctl != null)
            {
                // Kiểm tra giá trị của Cell
                if (this.Value == null || this.Value == DBNull.Value || string.IsNullOrEmpty(this.Value.ToString()))
                {
                    ctl.Value = DateTime.Now;
                }
                else
                {
                    // Ép kiểu giá trị hiện tại của ô vào DateTimePicker
                    if (DateTime.TryParse(this.Value.ToString(), out DateTime dateValue))
                    {
                        ctl.Value = dateValue;
                    }
                    else
                    {
                        ctl.Value = DateTime.Now;
                    }
                }
            }
        }

        public override Type EditType => typeof(CalendarEditingControl);
        public override Type ValueType => typeof(DateTime);
        public override object DefaultNewRowValue => DateTime.Now;
    }

    // Control hiển thị lịch khi nhấn vào ô
    public class CalendarEditingControl : DateTimePicker, IDataGridViewEditingControl
    {
        DataGridView dataGridView;
        private bool valueChanged = false;
        int rowIndex;

        public CalendarEditingControl()
        {
            this.Format = DateTimePickerFormat.Custom;
            this.CustomFormat = "dd-MM-yyyy";
        }

        public object EditingControlFormattedValue
        {
            get => this.Value.ToString("dd-MM-yyyy");
            set { if (value is String) this.Value = DateTime.Parse((String)value); }
        }

        public object GetEditingControlFormattedValue(DataGridViewDataErrorContexts context) => EditingControlFormattedValue;
        public void ApplyCellStyleToEditingControl(DataGridViewCellStyle dataGridViewCellStyle)
        {
            this.Font = dataGridViewCellStyle.Font;
            this.CalendarForeColor = dataGridViewCellStyle.ForeColor;
            this.CalendarMonthBackground = dataGridViewCellStyle.BackColor;
        }

        public int EditingControlRowIndex { get => rowIndex; set => rowIndex = value; }
        public bool EditingControlWantsInputKey(Keys key, bool dataGridViewWantsInputKey)
        {
            switch (key & Keys.KeyCode)
            {
                case Keys.Left:
                case Keys.Up:
                case Keys.Down:
                case Keys.Right:
                case Keys.Home:
                case Keys.End:
                case Keys.PageDown:
                case Keys.PageUp:
                    return true;
                default:
                    return !dataGridViewWantsInputKey;
            }
        }

        public void PrepareEditingControlForEdit(bool selectAll) { }
        public bool RepositionEditingControlOnValueChange => false;
        public DataGridView EditingControlDataGridView { get => dataGridView; set => dataGridView = value; }
        public bool EditingControlValueChanged { get => valueChanged; set => valueChanged = value; }
        public Cursor EditingPanelCursor => base.Cursor;

        protected override void OnValueChanged(EventArgs eventargs)
        {
            valueChanged = true;
            this.EditingControlDataGridView.NotifyCurrentCellDirty(true);
            base.OnValueChanged(eventargs);
        }
    }
}
