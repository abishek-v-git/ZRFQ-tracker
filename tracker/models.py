from django.db import models


class RFQEntry(models.Model):

    # ── Supplier ──────────────────────────────────────────────────────────────
    supplier_code = models.CharField(max_length=100, verbose_name='Supplier Code')
    supplier_name = models.CharField(max_length=255, verbose_name='Supplier Name')
    part_no = models.CharField(max_length=100, verbose_name='Part No')
    part_description = models.TextField(blank=True, verbose_name='Part Description')
    order_qty = models.DecimalField(max_digits=14, decimal_places=2, null=True, blank=True, verbose_name='Order Qty')
    uom = models.CharField(max_length=20, blank=True, verbose_name='UOM')
    unit_price = models.DecimalField(max_digits=16, decimal_places=4, null=True, blank=True, verbose_name='Unit Price')
    currency = models.CharField(max_length=10, blank=True, verbose_name='Currency')

    # ── Contact ───────────────────────────────────────────────────────────────
    pic = models.CharField(max_length=255, blank=True, verbose_name='PIC')
    contact_email = models.CharField(max_length=255, blank=True, verbose_name='Contact Email')
    contact_secondary_email = models.CharField(max_length=255, blank=True, verbose_name='Contact Secondary Email')

    # ── Lead Times ────────────────────────────────────────────────────────────
    lead_time_days = models.IntegerField(null=True, blank=True, verbose_name='Lead Time (days)')
    ship_lead_time_days = models.IntegerField(null=True, blank=True, verbose_name='Ship Lead Time (days)')

    # ── Supplier Quote ────────────────────────────────────────────────────────
    quote_uom = models.CharField(max_length=20, blank=True, verbose_name='UOM')
    coo = models.CharField(max_length=100, blank=True, verbose_name='COO')
    quote_currency = models.CharField(max_length=10, blank=True, verbose_name='Currency')
    unit_price_1 = models.DecimalField(max_digits=16, decimal_places=4, null=True, blank=True, verbose_name='Unit Price 1')
    moq_1 = models.DecimalField(max_digits=14, decimal_places=2, null=True, blank=True, verbose_name='MOQ-1')
    unit_price_2 = models.DecimalField(max_digits=16, decimal_places=4, null=True, blank=True, verbose_name='Unit Price 2')
    moq_2 = models.DecimalField(max_digits=14, decimal_places=2, null=True, blank=True, verbose_name='MOQ-2')
    unit_price_3 = models.DecimalField(max_digits=16, decimal_places=4, null=True, blank=True, verbose_name='Unit Price 3')
    moq_3 = models.DecimalField(max_digits=14, decimal_places=2, null=True, blank=True, verbose_name='MOQ-3')
    lot_size = models.DecimalField(max_digits=14, decimal_places=2, null=True, blank=True, verbose_name='Lot Size')

    # ── Product Codes ─────────────────────────────────────────────────────────
    hts_code = models.CharField(max_length=50, blank=True, verbose_name='HTS Code')
    eccn_ear99 = models.CharField(max_length=50, blank=True, verbose_name='ECCN/EAR99')

    # ── Manufacturer ──────────────────────────────────────────────────────────
    manufacture_part_number = models.CharField(max_length=100, blank=True, verbose_name='Manufacture Part Number')
    manufacturer_name = models.CharField(max_length=255, blank=True, verbose_name='Manufacturer Name')
    manufacturer_address = models.TextField(blank=True, verbose_name='Manufacturer Address (Street|City|ZIP|Country)')
    item_weight_kg = models.DecimalField(max_digits=10, decimal_places=4, null=True, blank=True, verbose_name='Item Weight (kg)')
    volume_weight_kg = models.DecimalField(max_digits=10, decimal_places=4, null=True, blank=True, verbose_name='Volume Weight (kg)')

    # ── Compliance ────────────────────────────────────────────────────────────
    russian_steel_confirmation = models.CharField(max_length=50, blank=True, verbose_name='Russian Steel Confirmation')
    hazmat = models.CharField(max_length=10, blank=True, verbose_name='Hazmat (Y/N)')
    un_sds_msds = models.CharField(max_length=255, blank=True, verbose_name='UN# · SDS/MSDS')
    product_regulation = models.CharField(max_length=255, blank=True, verbose_name='Product Regulation')
    eol_status = models.CharField(max_length=50, blank=True, verbose_name='EOL Status')
    alternative_parts = models.TextField(blank=True, verbose_name='Alternative Part(s)')
    alternative_part_no = models.CharField(max_length=255, blank=True, verbose_name='Alternative Part No')

    # ── China-specific ────────────────────────────────────────────────────────
    mfg_address_postal_cn = models.TextField(blank=True, verbose_name='Mfg Address + Postal (CN Only)')
    uflpa_compliance = models.CharField(max_length=100, blank=True, verbose_name='UFLPA Compliance Statement (CN Only)')
    uflpa_start_date = models.DateField(null=True, blank=True, verbose_name='UFLPA Start Date')
    uflpa_expiry_date = models.DateField(null=True, blank=True, verbose_name='UFLPA Expiry Date')

    # ── USMCA ─────────────────────────────────────────────────────────────────
    usmca_certificate = models.CharField(max_length=100, blank=True, verbose_name='USMCA Certificate (CA/MX Only)')
    usmca_start_date = models.DateField(null=True, blank=True, verbose_name='USMCA Start Date')
    usmca_expiry_date = models.DateField(null=True, blank=True, verbose_name='USMCA Expiry Date')

    # ── Workflow ──────────────────────────────────────────────────────────────
    status = models.CharField(max_length=255, blank=True, verbose_name='Status')
    comments = models.CharField(max_length=1000, blank=True, verbose_name='Comments')

    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['id']
        verbose_name = 'RFQ Entry'
        verbose_name_plural = 'RFQ Entries'

    def __str__(self):
        return f"{self.supplier_name} — {self.part_no}"
