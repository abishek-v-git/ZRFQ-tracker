from django.db import migrations, models


def backfill_rfq_sent(apps, schema_editor):
    RFQEntry = apps.get_model('tracker', 'RFQEntry')
    RFQEntry.objects.filter(rfq_sent='').update(rfq_sent='No')


class Migration(migrations.Migration):

    dependencies = [
        ('tracker', '0009_add_rfq_sent'),
    ]

    operations = [
        migrations.AlterField(
            model_name='rfqentry',
            name='rfq_sent',
            field=models.CharField(blank=True, default='No', max_length=10, verbose_name='RFQ Sent'),
        ),
        migrations.RunPython(backfill_rfq_sent, migrations.RunPython.noop),
    ]
