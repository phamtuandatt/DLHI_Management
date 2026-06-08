using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using MPR_Managerment.Forms.WarehouseGUI;

namespace MPR_Managerment.Helpers
{
    public class DataGridViewFilterHelper
    {
        private readonly DataGridView _dgv;
        private readonly DataTable _table;
        private readonly BindingSource _bindingSource;
        private readonly Dictionary<string, HashSet<string>> _columnFilters = new Dictionary<string, HashSet<string>>();

        public DataGridViewFilterHelper(DataGridView dgv, DataTable table, BindingSource bindingSource)
        {
            _dgv = dgv;
            _table = table;
            _bindingSource = bindingSource;
        }

        public void EnableFilter(bool enable)
        {
            if (enable)
            {
                _dgv.ColumnHeaderMouseClick -= Dgv_ColumnHeaderMouseClick;
                _dgv.ColumnHeaderMouseClick += Dgv_ColumnHeaderMouseClick;
            }
            else
            {
                _dgv.ColumnHeaderMouseClick -= Dgv_ColumnHeaderMouseClick;
            }
            _dgv.Invalidate();
        }

        public void ClearAllFilters()
        {
            _columnFilters.Clear();
            ApplyFilter();
            _dgv.Invalidate();
        }

        private void Dgv_ColumnHeaderMouseClick(object? sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex != -1 || e.ColumnIndex < 0) return;

            var colName = _dgv.Columns[e.ColumnIndex].Name;
            if (colName == "Chon") return;
            if (_table == null || _table.Rows.Count == 0) return;

            var currentSelected = _columnFilters.ContainsKey(colName)
                ? _columnFilters[colName].ToList()
                : _table.AsEnumerable().Select(r => r[colName]?.ToString() ?? "").Distinct().ToList();

            var filterForm = new FilterMenuForm(colName, _table, currentSelected);
            filterForm.Location = Cursor.Position;
            filterForm.ShowDialog();

            if (!filterForm.Canceled)
            {
                if (filterForm.IsSorted)
                {
                    _dgv.Sort(_dgv.Columns[colName], filterForm.SortDirection);
                }
                else
                {
                    _columnFilters[colName] = filterForm.SelectedValues.ToHashSet();
                    ApplyFilter();
                }
            }
        }

        private void ApplyFilter()
        {
            if (_table == null || _bindingSource == null) return;

            var filters = new List<string>();
            foreach (var kv in _columnFilters)
            {
                var colName = kv.Key;
                var selectedValues = kv.Value;
                if (selectedValues == null || selectedValues.Count == 0) continue;

                if (selectedValues.Count == _table.AsEnumerable().Select(r => r[colName]?.ToString() ?? "").Distinct().Count())
                    continue;

                var exprs = selectedValues.Select(v => $"ISNULL([{colName}], '') = '{v.Replace("'", "''")}'").ToList();
                filters.Add("(" + string.Join(" OR ", exprs) + ")");
            }

            _bindingSource.RemoveFilter();
            _bindingSource.Filter = filters.Count > 0 ? string.Join(" AND ", filters) : null;
        }
    }
}