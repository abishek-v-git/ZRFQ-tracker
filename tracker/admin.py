from django.contrib import admin
from .models import RFQEntry


@admin.register(RFQEntry)
class RFQEntryAdmin(admin.ModelAdmin):
    list_display = [
        'supplier_code', 'supplier_name', 'part_no', 'part_description',
        'order_qty', 'uom', 'currency', 'unit_price', 'manufacture_part_number',
        'manufacturer_name', 'eol_status',
    ]
    search_fields = ['supplier_name', 'supplier_code', 'part_no', 'manufacturer_name']
    list_filter = ['currency', 'uom', 'eol_status']
