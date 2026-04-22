from django.shortcuts import render, redirect, get_object_or_404
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.http import HttpResponse, JsonResponse
from .models import RFQEntry
from .forms import RFQEntryForm

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import datetime
import json
import re

# Maps Excel column headers → model field names
BULK_COLUMN_MAP = {
    'Supplier Code':                    'supplier_code',
    'Supplier Name':                    'supplier_name',
    'Part No':                          'part_no',
    'Part Description':                 'part_description',
    'Order Qty':                        'order_qty',
    'UOM':                              'uom',
    'Unit Price':                       'unit_price',
    'Currency':                         'currency',
    'PIC':                              'pic',
    'Contact Email':                    'contact_email',
    'Contact Secondary Email':          'contact_secondary_email',
    'Lead Time (days)':                 'lead_time_days',
    'Ship Lead Time (days)':            'ship_lead_time_days',
    'UOM (Quote)':                      'quote_uom',
    'COO':                              'coo',
    'Currency (Quote)':                 'quote_currency',
    'Unit Price 1':                     'unit_price_1',
    'MOQ-1':                            'moq_1',
    'Unit Price 2':                     'unit_price_2',
    'MOQ-2':                            'moq_2',
    'Unit Price 3':                     'unit_price_3',
    'MOQ-3':                            'moq_3',
    'Lot Size':                         'lot_size',
    'HTS Code':                         'hts_code',
    'ECCN/EAR99':                       'eccn_ear99',
    'Manufacture Part Number':          'manufacture_part_number',
    'Manufacturer Name':                'manufacturer_name',
    'Manufacturer Address':             'manufacturer_address',
    'Item Weight (kg)':                 'item_weight_kg',
    'Volume Weight (kg)':               'volume_weight_kg',
    'Russian Steel Confirmation':       'russian_steel_confirmation',
    'Hazmat (Y/N)':                     'hazmat',
    'UN# · SDS/MSDS':                   'un_sds_msds',
    'Product Regulation':               'product_regulation',
    'EOL Status':                       'eol_status',
    'Alternative Part(s)':              'alternative_parts',
    'Alternative Part No':              'alternative_part_no',
    'Mfg Address + Postal (CN Only)':   'mfg_address_postal_cn',
    'UFLPA Compliance (CN Only)':       'uflpa_compliance',
    'UFLPA Start Date':                 'uflpa_start_date',
    'UFLPA Expiry Date':                'uflpa_expiry_date',
    'USMCA Certificate (CA/MX Only)':   'usmca_certificate',
    'USMCA Start Date':                 'usmca_start_date',
    'USMCA Expiry Date':                'usmca_expiry_date',
    'Status':                           'status',
    'Comments':                         'comments',
}

# Known alternative spellings / truncations in supplier Excel templates
HEADER_ALIASES = {
    'Contact Secondar Email':   'Contact Secondary Email',
    'Contact Secondar email':   'Contact Secondary Email',
    'Lead Time':                'Lead Time (days)',
    'Ship Lead Time':           'Ship Lead Time (days)',
    'MOQ -1':                   'MOQ-1',
    'MOQ -2':                   'MOQ-2',
    'MOQ -3':                   'MOQ-3',
    'MOQ - 1':                  'MOQ-1',
    'MOQ - 2':                  'MOQ-2',
    'MOQ - 3':                  'MOQ-3',
    # Merged two-row headers (row1 + row2 combined text)
    'Manufacturer Address (Street|City|ZIP|Country)': 'Manufacturer Address',
    'Start Date':               'UFLPA Start Date',
    'Expiry Date':              'UFLPA Expiry Date',
    'Start Date (MM/DD/YYYY)':  'UFLPA Start Date',
    'Expiry Date (MM/DD/YYYY)': 'UFLPA Expiry Date',
    # Columns that may appear without their parenthetical suffix
    'UFLPA Compliance':                     'UFLPA Compliance (CN Only)',
    'UFLPA Compliance Statement':           'UFLPA Compliance (CN Only)',
    'UFLPA Compliance Statement (CN Only)': 'UFLPA Compliance (CN Only)',
    'Mfg Address + Postal':                 'Mfg Address + Postal (CN Only)',
    'USMCA Certificate':                    'USMCA Certificate (CA/MX Only)',
}

# When a header resolves to a field that was already assigned,
# map it to its second-occurrence counterpart (column-order dependent).
DUPLICATE_FIELD_MAP = {
    'uom':                'quote_uom',
    'currency':           'quote_currency',
    'uflpa_start_date':   'usmca_start_date',
    'uflpa_expiry_date':  'usmca_expiry_date',
}

DECIMAL_FIELDS = {
    'order_qty', 'unit_price', 'unit_price_1', 'moq_1',
    'unit_price_2', 'moq_2', 'unit_price_3', 'moq_3',
    'lot_size', 'item_weight_kg', 'volume_weight_kg',
}
INT_FIELDS = {'lead_time_days', 'ship_lead_time_days'}
DATE_FIELDS = {'uflpa_start_date', 'uflpa_expiry_date', 'usmca_start_date', 'usmca_expiry_date'}


def _normalize_header(text):
    """Collapse newlines / extra whitespace into a single space, strip
    dropdown-arrow indicator characters (▼ ▾ ▽) that Excel embeds."""
    if not text:
        return ''
    t = str(text)
    # Strip common dropdown indicator characters
    for ch in '\u25bc\u25be\u25bd\u25b6\u2193':   # ▼ ▾ ▽ ▶ ↓
        t = t.replace(ch, '')
    return ' '.join(t.split()).strip()


def _resolve_header(text):
    """Return the model field name for a header string, or None."""
    norm = _normalize_header(text)
    if norm in BULK_COLUMN_MAP:
        return BULK_COLUMN_MAP[norm]
    # Check aliases (maps alt text → canonical key)
    canonical = HEADER_ALIASES.get(norm)
    if canonical and canonical in BULK_COLUMN_MAP:
        return BULK_COLUMN_MAP[canonical]
    # Fallback: normalise spaces around hyphens (e.g. "MOQ -1" → "MOQ-1")
    collapsed = re.sub(r'\s*-\s*', '-', norm)
    if collapsed in BULK_COLUMN_MAP:
        return BULK_COLUMN_MAP[collapsed]
    canonical2 = HEADER_ALIASES.get(collapsed)
    if canonical2 and canonical2 in BULK_COLUMN_MAP:
        return BULK_COLUMN_MAP[canonical2]
    return None


def _coerce(field, value):
    if value is None or str(value).strip() == '':
        # Numeric/date fields accept NULL; CharFields must use empty string
        if field in DECIMAL_FIELDS or field in INT_FIELDS or field in DATE_FIELDS:
            return None
        return ''
    if field in DECIMAL_FIELDS:
        try:
            return float(str(value).replace(',', ''))
        except (ValueError, TypeError):
            return None
    if field in INT_FIELDS:
        try:
            return int(float(str(value).replace(',', '')))
        except (ValueError, TypeError):
            return None
    if field in DATE_FIELDS:
        if isinstance(value, (datetime.date, datetime.datetime)):
            return value if isinstance(value, datetime.date) else value.date()
        s = str(value).strip()
        if not s: return None
        # Try DD/MM/YYYY first (as requested by user)
        try:
            return datetime.datetime.strptime(s, '%d/%m/%Y').date()
        except ValueError:
            # Fallback to MM/DD/YYYY
            try:
                return datetime.datetime.strptime(s, '%m/%d/%Y').date()
            except ValueError:
                # Fallback to ISO YYYY-MM-DD
                try:
                    return datetime.datetime.strptime(s, '%Y-%m-%d').date()
                except ValueError:
                    return None
    return str(value).strip()


def _entry_to_dict(entry):
    def fd(d): return d.isoformat() if d else ''
    def fn(v): return str(v) if v is not None else ''
    return {
        'pk': entry.pk,
        'supplier_code': entry.supplier_code or '',
        'supplier_name': entry.supplier_name or '',
        'part_no': entry.part_no or '',
        'part_description': entry.part_description or '',
        'order_qty': fn(entry.order_qty),
        'uom': entry.uom or '',
        'unit_price': fn(entry.unit_price),
        'currency': entry.currency or '',
        'pic': entry.pic or '',
        'contact_email': entry.contact_email or '',
        'contact_secondary_email': entry.contact_secondary_email or '',
        'lead_time_days': fn(entry.lead_time_days),
        'ship_lead_time_days': fn(entry.ship_lead_time_days),
        'quote_uom': entry.quote_uom or '',
        'coo': entry.coo or '',
        'quote_currency': entry.quote_currency or '',
        'unit_price_1': fn(entry.unit_price_1),
        'moq_1': fn(entry.moq_1),
        'unit_price_2': fn(entry.unit_price_2),
        'moq_2': fn(entry.moq_2),
        'unit_price_3': fn(entry.unit_price_3),
        'moq_3': fn(entry.moq_3),
        'lot_size': fn(entry.lot_size),
        'hts_code': entry.hts_code or '',
        'eccn_ear99': entry.eccn_ear99 or '',
        'manufacture_part_number': entry.manufacture_part_number or '',
        'manufacturer_name': entry.manufacturer_name or '',
        'manufacturer_address': entry.manufacturer_address or '',
        'item_weight_kg': fn(entry.item_weight_kg),
        'volume_weight_kg': fn(entry.volume_weight_kg),
        'russian_steel_confirmation': entry.russian_steel_confirmation or '',
        'hazmat': entry.hazmat or '',
        'un_sds_msds': entry.un_sds_msds or '',
        'product_regulation': entry.product_regulation or '',
        'eol_status': entry.eol_status or '',
        'alternative_parts': entry.alternative_parts or '',
        'alternative_part_no': entry.alternative_part_no or '',
        'mfg_address_postal_cn': entry.mfg_address_postal_cn or '',
        'uflpa_compliance': entry.uflpa_compliance or '',
        'uflpa_start_date': fd(entry.uflpa_start_date),
        'uflpa_expiry_date': fd(entry.uflpa_expiry_date),
        'usmca_certificate': entry.usmca_certificate or '',
        'usmca_start_date': fd(entry.usmca_start_date),
        'usmca_expiry_date': fd(entry.usmca_expiry_date),
        'status': entry.status or '',
        'comments': entry.comments or '',
    }


@login_required
def rfq_list(request):
    form = RFQEntryForm()
    edit_form = RFQEntryForm(prefix='edit')
    total_count = RFQEntry.objects.count()
    return render(request, 'tracker/rfq_list.html', {'total_count': total_count, 'form': form, 'edit_form': edit_form})


@login_required
def rfq_data(request):
    from django.db.models import Q
    from django.core.paginator import Paginator

    q = request.GET.get('q', '').strip()
    try:
        page = int(request.GET.get('page', 1))
    except (ValueError, TypeError):
        page = 1
    page_size = request.GET.get('page_size', '20')
    try:
        sort_col = int(request.GET.get('sort_col', -1))
    except (ValueError, TypeError):
        sort_col = -1
    sort_dir = request.GET.get('sort_dir', 'asc')

    qs = RFQEntry.objects.all()

    if q:
        qs = qs.filter(
            Q(supplier_code__icontains=q) | Q(supplier_name__icontains=q) |
            Q(part_no__icontains=q) | Q(part_description__icontains=q) |
            Q(manufacture_part_number__icontains=q) | Q(manufacturer_name__icontains=q) |
            Q(pic__icontains=q) | Q(coo__icontains=q) | Q(hts_code__icontains=q) |
            Q(eccn_ear99__icontains=q) | Q(status__icontains=q) | Q(comments__icontains=q)
        )

    COL_FIELD = {
        0: 'pk', 1: 'supplier_code', 2: 'supplier_name', 3: 'part_no',
        4: 'part_description', 5: 'order_qty', 6: 'uom', 7: 'unit_price',
        8: 'currency', 9: 'pic', 10: 'contact_email', 11: 'contact_secondary_email',
        12: 'lead_time_days', 13: 'ship_lead_time_days', 14: 'quote_uom',
        15: 'coo', 16: 'quote_currency', 17: 'unit_price_1', 18: 'moq_1',
        19: 'unit_price_2', 20: 'moq_2', 21: 'unit_price_3', 22: 'moq_3',
        23: 'lot_size', 24: 'hts_code', 25: 'eccn_ear99',
        26: 'manufacture_part_number', 27: 'manufacturer_name',
        28: 'manufacturer_address', 29: 'item_weight_kg', 30: 'volume_weight_kg',
        31: 'russian_steel_confirmation', 32: 'hazmat', 33: 'un_sds_msds',
        34: 'product_regulation', 35: 'eol_status', 36: 'alternative_parts',
        37: 'alternative_part_no', 38: 'mfg_address_postal_cn',
        39: 'uflpa_compliance', 40: 'uflpa_start_date', 41: 'uflpa_expiry_date',
        42: 'usmca_certificate', 43: 'usmca_start_date', 44: 'usmca_expiry_date',
        45: 'status', 46: 'comments',
    }

    field = COL_FIELD.get(sort_col)
    if field:
        order = field if sort_dir == 'asc' else '-' + field
        qs = qs.order_by(order)

    total = qs.count()

    if page_size == 'all':
        entries = list(qs)
        pages = 1
        page = 1
    else:
        try:
            page_size_int = int(page_size)
        except (ValueError, TypeError):
            page_size_int = 20
        paginator = Paginator(qs, page_size_int)
        page_obj = paginator.get_page(page)
        entries = list(page_obj.object_list)
        pages = paginator.num_pages
        page = page_obj.number

    return JsonResponse({
        'entries': [_entry_to_dict(e) for e in entries],
        'total': total,
        'page': page,
        'pages': pages,
    })


@login_required
def rfq_add(request):
    if request.method == 'POST':
        form = RFQEntryForm(request.POST)
        if form.is_valid():
            form.save()
            messages.success(request, 'RFQ entry added successfully.')
            return redirect('rfq_list')
        else:
            messages.error(request, 'Please correct the errors below.')
            total_count = RFQEntry.objects.count()
            edit_form = RFQEntryForm(prefix='edit')
            return render(request, 'tracker/rfq_list.html', {'total_count': total_count, 'form': form, 'edit_form': edit_form})
    return redirect('rfq_list')


@login_required
def rfq_delete(request, pk):
    entry = get_object_or_404(RFQEntry, pk=pk)
    if request.method == 'POST':
        entry.delete()
        messages.success(request, f'Entry for "{entry.supplier_name} — {entry.part_no}" deleted.')
    return redirect('rfq_list')


@login_required
def rfq_edit(request, pk):
    entry = get_object_or_404(RFQEntry, pk=pk)
    if request.method == 'POST':
        form = RFQEntryForm(request.POST, instance=entry)
        if form.is_valid():
            form.save()
            messages.success(request, 'Entry updated successfully.')
            return redirect('rfq_list')
    else:
        form = RFQEntryForm(instance=entry)
    return render(request, 'tracker/rfq_edit.html', {'form': form, 'entry': entry})


@login_required
def rfq_export(request):
    entries = RFQEntry.objects.all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'RFQ Tracker'

    headers = [
        'Supplier Code', 'Supplier Name', 'Part No', 'Part Description',
        'Order Qty', 'UOM', 'Unit Price', 'Currency',
        'PIC', 'Contact Email', 'Contact Secondary Email',
        'Lead Time (days)', 'Ship Lead Time (days)',
        'UOM (Quote)', 'COO', 'Currency (Quote)',
        'Unit Price 1', 'MOQ-1', 'Unit Price 2', 'MOQ-2', 'Unit Price 3', 'MOQ-3',
        'Lot Size', 'HTS Code', 'ECCN/EAR99',
        'Manufacture Part Number', 'Manufacturer Name',
        'Manufacturer Address', 'Item Weight (kg)', 'Volume Weight (kg)',
        'Russian Steel Confirmation', 'Hazmat (Y/N)', 'UN# · SDS/MSDS',
        'Product Regulation', 'EOL Status',
        'Alternative Part(s)', 'Alternative Part No',
        'Mfg Address + Postal (CN Only)', 'UFLPA Compliance (CN Only)',
        'UFLPA Start Date', 'UFLPA Expiry Date',
        'USMCA Certificate (CA/MX Only)', 'USMCA Start Date', 'USMCA Expiry Date',
        'Status', 'Comments',
    ]

    header_fill = PatternFill(start_color='003366', end_color='003366', fill_type='solid')
    header_font = Font(bold=True, color='FFFFFF', size=11)
    mfr_fill = PatternFill(start_color='FFFBEA', end_color='FFFBEA', fill_type='solid')

    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for row_idx, entry in enumerate(entries, start=2):
        row = [
            entry.supplier_code,
            entry.supplier_name,
            entry.part_no,
            entry.part_description,
            float(entry.order_qty) if entry.order_qty is not None else '',
            entry.uom,
            float(entry.unit_price) if entry.unit_price is not None else '',
            entry.currency,
            entry.pic,
            entry.contact_email,
            entry.contact_secondary_email,
            entry.lead_time_days or '',
            entry.ship_lead_time_days or '',
            entry.quote_uom,
            entry.coo,
            entry.quote_currency,
            float(entry.unit_price_1) if entry.unit_price_1 is not None else '',
            float(entry.moq_1) if entry.moq_1 is not None else '',
            float(entry.unit_price_2) if entry.unit_price_2 is not None else '',
            float(entry.moq_2) if entry.moq_2 is not None else '',
            float(entry.unit_price_3) if entry.unit_price_3 is not None else '',
            float(entry.moq_3) if entry.moq_3 is not None else '',
            float(entry.lot_size) if entry.lot_size is not None else '',
            entry.hts_code,
            entry.eccn_ear99,
            entry.manufacture_part_number,
            entry.manufacturer_name,
            entry.manufacturer_address,
            float(entry.item_weight_kg) if entry.item_weight_kg is not None else '',
            float(entry.volume_weight_kg) if entry.volume_weight_kg is not None else '',
            entry.russian_steel_confirmation,
            entry.hazmat,
            entry.un_sds_msds,
            entry.product_regulation,
            entry.eol_status,
            entry.alternative_parts,
            entry.alternative_part_no,
            entry.mfg_address_postal_cn,
            entry.uflpa_compliance,
            entry.uflpa_start_date,
            entry.uflpa_expiry_date,
            entry.usmca_certificate,
            entry.usmca_start_date,
            entry.usmca_expiry_date,
            entry.status,
            entry.comments,
        ]
        for col_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            # Highlight Manufacture Part Number column (col 27)
            if col_idx == 27:
                cell.fill = mfr_fill

    # Auto-fit column widths
    for col_idx, header in enumerate(headers, start=1):
        col_letter = get_column_letter(col_idx)
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val:
                max_len = max(max_len, len(str(val)))
        ws.column_dimensions[col_letter].width = min(max_len + 4, 50)

    ws.freeze_panes = 'A2'
    ws.row_dimensions[1].height = 20

    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = 'attachment; filename="RFQ_Tracker_Export.xlsx"'
    wb.save(response)
    return response


@login_required
def rfq_bulk_upload(request):
    if request.method != 'POST':
        return redirect('rfq_list')

    files = request.FILES.getlist('excel_files')
    if not files:
        messages.error(request, 'No files selected.')
        return redirect('rfq_list')

    total_added = 0
    errors = []

    for f in files:
        try:
            wb = openpyxl.load_workbook(f, data_only=True)
            ws = wb.active

            # ── Build header → column-index map ──────────────────────
            # Handles: embedded newlines in cells, two-row merged headers,
            # and known typos / truncations via HEADER_ALIASES.
            headers = {}
            row1_cells = list(ws[1])
            row2_cells = list(ws[2]) if ws.max_row >= 2 else []
            data_start_row = 2          # default: data begins right after row 1
            uses_row2_header = False     # flipped if any column needs row-2 text

            for ci, cell in enumerate(row1_cells):
                # Normalise: replaces \n, dropdown arrows, excess whitespace
                r1 = _normalize_header(cell.value)
                field = _resolve_header(r1)

                if not field and ci < len(row2_cells) and row2_cells[ci].value:
                    # Try combining row-1 + row-2 (e.g. "Lead Time" + "(days)")
                    r2 = _normalize_header(row2_cells[ci].value)
                    combined = f"{r1} {r2}".strip()
                    field = _resolve_header(combined)
                    if field:
                        uses_row2_header = True

                if field:
                    # Handle duplicate headers (e.g. two "UOM" columns)
                    if field in headers:
                        field = DUPLICATE_FIELD_MAP.get(field)
                    if field and field not in headers:
                        headers[field] = ci + 1       # store 1-based index

            if uses_row2_header:
                data_start_row = 3

            if 'supplier_code' not in headers and 'supplier_name' not in headers:
                errors.append(f"{f.name}: could not find recognisable column headers.")
                continue

            added = 0
            for row in ws.iter_rows(min_row=data_start_row, values_only=True):
                # Skip completely empty rows
                if all(v is None or str(v).strip() == '' for v in row):
                    continue

                kwargs = {}
                for field, col_idx in headers.items():
                    raw = row[col_idx - 1] if col_idx - 1 < len(row) else None
                    kwargs[field] = _coerce(field, raw)

                # supplier_code and part_no are required
                if not kwargs.get('supplier_code') and not kwargs.get('supplier_name'):
                    continue

                RFQEntry.objects.create(**kwargs)
                added += 1

            total_added += added

        except Exception as e:
            errors.append(f"{f.name}: {e}")

    if total_added:
        messages.success(request, f'Bulk upload complete — {total_added} row(s) added.')
    if errors:
        for err in errors:
            messages.error(request, err)

    return redirect('rfq_list')


@login_required
def rfq_clear_all(request):
    """Delete every RFQEntry (POST only)."""
    if request.method == 'POST':
        count, _ = RFQEntry.objects.all().delete()
        messages.success(request, f'All data cleared — {count} record(s) deleted.')
    return redirect('rfq_list')


@login_required
def rfq_patch(request, pk):
    """Inline-update a single dropdown field via AJAX."""
    if request.method != 'POST':
        return JsonResponse({'ok': False}, status=405)
    entry = get_object_or_404(RFQEntry, pk=pk)
    try:
        data = json.loads(request.body)
        field = data.get('field', '')
        value = data.get('value', '')
    except (json.JSONDecodeError, AttributeError):
        return JsonResponse({'ok': False, 'error': 'Invalid JSON'}, status=400)
    PATCHABLE = {
        'russian_steel_confirmation', 'hazmat', 'un_sds_msds',
        'product_regulation', 'eol_status', 'uflpa_compliance', 'usmca_certificate',
        'status', 'comments',
    }
    if field not in PATCHABLE:
        return JsonResponse({'ok': False, 'error': 'Field not allowed'}, status=400)
    setattr(entry, field, value)
    entry.save(update_fields=[field])
    return JsonResponse({'ok': True})


@login_required
def rfq_edit_json(request, pk):
    """Load (GET) or save (POST) an entry via AJAX — used by the edit modal."""
    entry = get_object_or_404(RFQEntry, pk=pk)
    if request.method == 'POST':
        form = RFQEntryForm(request.POST, instance=entry, prefix='edit')
        if form.is_valid():
            saved = form.save()
            return JsonResponse({'ok': True, 'entry': _entry_to_dict(saved)})
        errors = {k: [str(e) for e in v] for k, v in form.errors.items()}
        return JsonResponse({'ok': False, 'errors': errors}, status=400)
    return JsonResponse({'ok': True, 'entry': _entry_to_dict(entry)})
