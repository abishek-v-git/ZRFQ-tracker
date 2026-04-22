from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('tracker', '0006_change_comments_to_textfield'),
    ]

    operations = [
        migrations.RunSQL(
            sql='ALTER TABLE tracker_rfqentry DROP COLUMN IF EXISTS priority;',
            reverse_sql='ALTER TABLE tracker_rfqentry ADD COLUMN priority character varying(255) NOT NULL DEFAULT \'\';',
        ),
    ]
