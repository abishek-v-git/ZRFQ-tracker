from django import forms
from .models import RFQEntry

_text = lambda ph='': {'class': 'form-control form-control-sm', 'placeholder': ph}
_num  = lambda step='1': {'class': 'form-control form-control-sm', 'step': step}
_date = lambda: {'class': 'form-control form-control-sm', 'type': 'date'}
_area = lambda rows=2: {'class': 'form-control form-control-sm', 'rows': rows}
_sel  = lambda: {'class': 'form-select form-select-sm'}


class RFQEntryForm(forms.ModelForm):
    class Meta:
        model = RFQEntry
        exclude = ['created_at']
        widgets = {
            # Supplier
            'supplier_code':             forms.TextInput(attrs=_text()),
            'supplier_name':             forms.TextInput(attrs=_text()),
            'part_no':                   forms.TextInput(attrs=_text()),
            'part_description':          forms.Textarea(attrs=_area(2)),
            'order_qty':                 forms.NumberInput(attrs=_num('0.01')),
            'uom':                       forms.TextInput(attrs=_text('e.g. EA, PC, KG')),
            'unit_price':                forms.NumberInput(attrs=_num('0.0001')),
            'currency':                  forms.TextInput(attrs=_text('e.g. USD, EUR')),
            # Contact
            'pic':                       forms.TextInput(attrs=_text()),
            'contact_email':             forms.TextInput(attrs=_text()),
            'contact_secondary_email':   forms.TextInput(attrs=_text()),
            # Lead times
            'lead_time_days':            forms.NumberInput(attrs=_num()),
            'ship_lead_time_days':       forms.NumberInput(attrs=_num()),
            # Quote
            'quote_uom':                 forms.TextInput(attrs=_text()),
            'coo':                       forms.TextInput(attrs=_text('e.g. US, CN, DE')),
            'quote_currency':            forms.TextInput(attrs=_text('e.g. USD')),
            'unit_price_1':              forms.NumberInput(attrs=_num('0.0001')),
            'moq_1':                     forms.NumberInput(attrs=_num('0.01')),
            'unit_price_2':              forms.NumberInput(attrs=_num('0.0001')),
            'moq_2':                     forms.NumberInput(attrs=_num('0.01')),
            'unit_price_3':              forms.NumberInput(attrs=_num('0.0001')),
            'moq_3':                     forms.NumberInput(attrs=_num('0.01')),
            'lot_size':                  forms.NumberInput(attrs=_num('0.01')),
            # Product codes
            'hts_code':                  forms.TextInput(attrs=_text()),
            'eccn_ear99':                forms.TextInput(attrs=_text('e.g. EAR99')),
            # Manufacturer
            'manufacture_part_number':   forms.TextInput(attrs=_text()),
            'manufacturer_name':         forms.TextInput(attrs=_text()),
            'manufacturer_address':      forms.Textarea(attrs=_area(2)),
            'item_weight_kg':            forms.NumberInput(attrs=_num('0.0001')),
            'volume_weight_kg':          forms.NumberInput(attrs=_num('0.0001')),
            # Compliance
            'russian_steel_confirmation': forms.Select(attrs=_sel(), choices=[
                ('', '—'), ('Yes', 'Yes'), ('No', 'No'),
            ]),
            'hazmat':                    forms.Select(attrs=_sel(), choices=[
                ('', '—'), ('Yes', 'Yes'), ('No', 'No'),
            ]),
            'un_sds_msds':               forms.Select(attrs=_sel(), choices=[
                ('', '—'), ('Yes', 'Yes'), ('No', 'No'),
            ]),
            'product_regulation':        forms.Select(attrs=_sel(), choices=[
                ('', '—'), ('FDA Compliance', 'FDA Compliance'), ('FCC Certification', 'FCC Certification'), 
                ('CE', 'CE'), ('Others', 'Others'), ('Not Applicable', 'Not Applicable'),
            ]),
            'eol_status':                forms.Select(attrs=_sel(), choices=[
                ('', '—'), ('Yes', 'Yes'), ('No', 'No'),
            ]),
            'alternative_parts':         forms.Textarea(attrs=_area(2)),
            'alternative_part_no':       forms.TextInput(attrs=_text()),
            # CN-specific
            'mfg_address_postal_cn':     forms.Textarea(attrs=_area(2)),
            'uflpa_compliance':          forms.Select(attrs=_sel(), choices=[
                ('', '—'), ('Yes', 'Yes'), ('No', 'No'),
            ]),
            'uflpa_start_date':          forms.DateInput(attrs=_date()),
            'uflpa_expiry_date':         forms.DateInput(attrs=_date()),
            # USMCA
            'usmca_certificate':         forms.Select(attrs=_sel(), choices=[
                ('', '—'), ('Yes', 'Yes'), ('No', 'No'),
            ]),
            'usmca_start_date':          forms.DateInput(attrs=_date()),
            'usmca_expiry_date':         forms.DateInput(attrs=_date()),
            'status': forms.Select(attrs=_sel(), choices=[
                ('', '—'), ('No Response Yet', 'No Response Yet'), 
                ('Partially Data Received', 'Partially Data Received'), 
                ('Completed', 'Completed'), ('Others', 'Others'),
            ]),
            'comments':                  forms.Textarea(attrs=_area(2)),
        }
